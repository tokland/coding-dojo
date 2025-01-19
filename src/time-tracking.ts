import _ from "./domain/entities/generic/Collection";
import { Struct } from "./domain/entities/generic/Struct";
import { Future } from "./domain/entities/generic/Future";

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

    static approve(timeTrackings: TimeTracking[]): TimeTracking[] {
        return timeTrackings.map(timeTracking => timeTracking.approve());
    }

    static getManagerIds(timeTrackings: TimeTracking[]): Id[] {
        return _(timeTrackings)
            .map(timeTracking => timeTracking.managerId)
            .uniq()
            .value();
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
    getById(ids: Id): Async<Manager>;
}

interface TimeTrackingRepository {
    getByManager(options: { managerId: Id; months: CalendarMonth[] }): Async<TimeTracking[]>;
    save(timeTrackings: TimeTracking[]): Async<void>;
}

interface MessageRepository {
    send(message: Message): Async<void>;
}

// DOMAIN: use cases

export class GetTimeTrackingUseCase {
    constructor(private repositories: { timeTracking: TimeTrackingRepository }) {}

    execute(options: { managerId: Id; month: CalendarMonth }): Async<TimeTracking[]> {
        return this.repositories.timeTracking.getByManager({
            managerId: options.managerId,
            months: [options.month],
        });
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
    executeUsingFutureBlock(timeTrackings: TimeTracking[]): Async<void> {
        return Future.block(async $ => {
            const timeTrackingsApproved = TimeTracking.approve(timeTrackings);
            await $(this.repositories.timeTracking.save(timeTrackingsApproved));
            await $(this.notifyManagers(timeTrackings));
        });
    }

    // Create private methods for actions so the `execute` method stays readable

    private notifyManagers(timeTrackings: TimeTracking[]): Async<void> {
        return Future.block(async $ => {
            for (const managerId of TimeTracking.getManagerIds(timeTrackings)) {
                const manager = await $(this.repositories.manager.getById(managerId));
                const timeTrackingsForManager = timeTrackings.filter(tt => tt.managerId === managerId);
                await $(this.notifyManager(manager, timeTrackingsForManager));
            }
        });
    }

    private notifyManager(manager: Manager, timeTrackings: TimeTracking[]): Async<void> {
        return this.repositories.message.send(
            Message.create({
                recipients: [manager.email],
                title: "Time trackings approved",
                text: `Hello ${manager.name}, time tracking approved: ${timeTrackings.length}`,
            })
        );
    }
}

// DATA: repository implementations
