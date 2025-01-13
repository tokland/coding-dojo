import { Struct } from "./domain/entities/generic/Struct";
import { Future } from "./domain/entities/generic/Future";

// DOMAIN: entities

type Id = string;

type Async<Data> = Future<Error, Data>;

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
}>() {
    toDate() {
        return new Date(this.year, this.month - 1, 1);
    }
}

// or maybe Notification?
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
    get(options: { managerId: Id; months: CalendarMonth[] }): Async<TimeTracking[]>;
    save(timeTrackings: TimeTracking[]): Async<void>; // or should we have approve() instead?
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
        return this.getTimeTrackings(options)
            .map(timeTrackings => this.approveTimeTrackings(timeTrackings))
            .flatMap(timeTrackingsApproved => this.repositories.timeTracking.save(timeTrackingsApproved))
            .flatMap(() => this.repositories.manager.getById(options.managerId))
            .flatMap(manager => this.sendMessageToManager(manager, { month: options.month }));
    }

    // Demo: promise-like syntax with Future.block
    executeUsingNewAsyncBlock(options: { month: CalendarMonth; managerId: Id }): Async<void> {
        return Future.block(async $ => {
            const timeTrackings = await $(this.getTimeTrackings(options));
            const timeTrackingsApproved = this.approveTimeTrackings(timeTrackings);
            await $(this.repositories.timeTracking.save(timeTrackingsApproved));
            const manager = await $(this.repositories.manager.getById(options.managerId));
            await $(this.sendMessageToManager(manager, { month: options.month }));
        });
    }

    // Create private methods for long actions so the `execute` method stays readable

    private getTimeTrackings(options: { month: CalendarMonth; managerId: Id }): Async<TimeTracking[]> {
        return this.repositories.timeTracking.get({
            managerId: options.managerId,
            months: [options.month],
        });
    }

    private approveTimeTrackings(timeTrackings: TimeTracking[]): TimeTracking[] {
        return timeTrackings
            .filter(timeTracking => !timeTracking.isApproved)
            .map(timeTracking => timeTracking.approve());
    }

    private sendMessageToManager(manager: Manager, options: { month: CalendarMonth }): Async<void> {
        const monthName = options.month
            .toDate()
            .toLocaleString("default", { month: "long", year: "numeric" });

        return this.repositories.message.send(
            Message.create({
                recipients: [manager.email],
                title: "Time tracking approved",
                text: `Your time tracking for ${monthName} has been approved`,
            })
        );
    }
}

// DATA: repository implementations
