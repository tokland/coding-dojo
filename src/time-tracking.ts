import { Future } from "./domain/entities/generic/Future";
import { Struct } from "./domain/entities/generic/Struct";

// DOMAIN: entities

type Id = string;

type Async<T> = Future<Error, T>;

class Manager extends Struct<{
    id: Id;
    name: string;
    email: string;
}>() {}

class TimeTracking extends Struct<{
    id: Id;
    manager: { id: Id }; // or managerId: Id ?
    day: CalendarDay;
    hours: number;
    description: string;
    isApproved: boolean;
}>() {
    approve() {
        return this._update({ isApproved: true });
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

// DOMAIN: repositories

interface ManagerRepository {
    getById(id: Id): Async<Manager>;
}

interface TimeTrackingRepository {
    get(options: { managerId: Id; month: CalendarMonth }): Async<TimeTracking[]>;
    // or should we have a approve method instead?
    save(timeTrackings: TimeTracking[]): Async<void>;
}

interface MessageRepository {
    send(message: Message): Async<void>;
}

// DOMAIN: use cases

export class ApproveManagerTimeTrackingUseCase {
    constructor(
        private repositories: {
            timeTracking: TimeTrackingRepository;
            manager: ManagerRepository;
            message: MessageRepository;
        }
    ) {}

    execute(options: { month: CalendarMonth; managerId: Id }): Async<void> {
        return this.repositories.timeTracking
            .get(options)
            .map(timeTrackings => this.approveTimeTrackings(timeTrackings))
            .flatMap(timeTrackingsApproved =>
                this.repositories.timeTracking.save(timeTrackingsApproved)
            )
            .flatMap(() => this.repositories.manager.getById(options.managerId))
            .flatMap(manager => this.sendMessageToManager(manager, { month: options.month }));
    }

    // As a demo on how to use Async.block (which allows a promise-like syntax)
    executeUsingNewAsyncBlock(options: { month: CalendarMonth; managerId: Id }): Async<void> {
        return Future.block(async $ => {
            const timeTrackings = await $(this.repositories.timeTracking.get(options));
            const timeTrackingsApproved = this.approveTimeTrackings(timeTrackings);
            await $(this.repositories.timeTracking.save(timeTrackingsApproved));
            const manager = await $(this.repositories.manager.getById(options.managerId));
            await $(this.sendMessageToManager(manager, { month: options.month }));
        });
    }

    // We create private methods for longer actions so the `execute` method stays readable

    private approveTimeTrackings(timeTrackings: TimeTracking[]): TimeTracking[] {
        return timeTrackings
            .filter(timeTracking => !timeTracking.isApproved)
            .map(timeTracking => timeTracking.approve());
    }

    private sendMessageToManager(manager: Manager, options: { month: CalendarMonth }): Async<void> {
        return this.repositories.message.send(
            new Message({
                recipients: [manager.email],
                title: "Time tracking approved",
                text: `Your time tracking for ${options.month} has been approved`,
            })
        );
    }
}

// DATA: repository implementations (NOT DONE)
