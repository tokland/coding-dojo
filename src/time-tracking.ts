import { Struct } from "./domain/entities/generic/Struct";
import { Future } from "./domain/entities/generic/Future";
import _ from "./domain/entities/generic/Collection";

// DOMAIN: entities

class Manager extends Struct<{
    id: Id;
    name: string;
    email: string;
}>() {}

class TimeTracking extends Struct<{
    id: Id;
    managerId: Id;
    day: CalendarDay;
    hours: number;
    description: string;
    isApproved: boolean;
}>() {
    approve() {
        return this._update({ isApproved: true });
    }

    static getManagerIds(timeTrackings: TimeTracking[]): Id[] {
        return _(timeTrackings)
            .map(timeTracking => timeTracking.managerId)
            .uniq()
            .value();
    }

    static approve(timeTrackings: TimeTracking[]): TimeTracking[] {
        return timeTrackings.map(timeTracking => timeTracking.approve());
    }
}

class CalendarDay extends Struct<{
    year: number;
    month: number;
    day: number;
}>() {}

class CalendarMonth extends Struct<{
    year: number;
    month: number;
}>() {}

class Message extends Struct<{
    recipients: string[];
    title: string;
    text: string;
}>() {}

type Id = string;

type Async<Data> = Future<Error, Data>;

// DOMAIN: repositories

interface ManagerRepository {
    getByIds(ids: Id[]): Async<Manager[]>;
}

interface TimeTrackingRepository {
    get(options: { managerId: Id; months: CalendarMonth[] }): Async<TimeTracking[]>;
    save(timeTrackings: TimeTracking[]): Async<void>;
}

interface MessageRepository {
    send(message: Message): Async<void>;
}

// DOMAIN: use cases

export class GetTimeTrackingUseCase {
    constructor(private repositories: { timeTracking: TimeTrackingRepository }) {}

    execute(managerId: Id, month: CalendarMonth): Async<TimeTracking[]> {
        return this.repositories.timeTracking.get({ managerId, months: [month] });
    }
}

export class ApproveManagerTimeTrackingUseCase {
    constructor(
        private repositories: {
            timeTracking: TimeTrackingRepository;
            manager: ManagerRepository;
            message: MessageRepository;
        }
    ) {}

    execute(timeTrackings: TimeTracking[]): Async<void> {
        const timeTrackingsApproved = TimeTracking.approve(timeTrackings);

        return this.repositories.timeTracking
            .save(timeTrackingsApproved)
            .flatMap(() => this.notifyManagers(timeTrackings));
    }

    // Demo: promise-like syntax with Future.block
    executeUsingNewAsyncBlock(timeTrackings: TimeTracking[]): Async<void> {
        return Future.block(async $ => {
            const timeTrackingsApproved = TimeTracking.approve(timeTrackings);
            await $(this.repositories.timeTracking.save(timeTrackingsApproved));
            await $(this.notifyManagers(timeTrackings));
        });
    }

    // Create private methods for long actions so the `execute` method stays readable

    private notifyManagers(timeTrackings: TimeTracking[]): Async<void> {
        return Future.block(async $ => {
            const managerIds = TimeTracking.getManagerIds(timeTrackings);
            const managers = await $(this.repositories.manager.getByIds(managerIds));
            const sendMessages$ = managers.map(manager => this.notifyManager(manager));
            await $(Future.sequential(sendMessages$));
        });
    }

    private notifyManager(manager: Manager): Async<void> {
        return this.repositories.message.send(
            Message.create({
                recipients: [manager.email],
                title: "Time tracking approved",
                text: `Hello ${manager.name}, your time tracking for have been approved`,
            })
        );
    }
}

// DATA: repository implementations
