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
}>() {}

// naming: Notification?
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
    // save() instead of approve() allows the approval workflow to live in the domain layer
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
        return this.getTimeTrackings(options)
            .map(timeTrackings => this.approveTimeTrackings(timeTrackings))
            .flatMap(timeTrackingsApproved => this.repositories.timeTracking.save(timeTrackingsApproved))
            .flatMap(() => this.sendApprovalMessageToManager(options));
    }

    // Demo: promise-like syntax with Future.block
    executeUsingNewAsyncBlock(options: { month: CalendarMonth; managerId: Id }): Async<void> {
        return Future.block(async $ => {
            const timeTrackings = await $(this.getTimeTrackings(options));
            const timeTrackingsApproved = this.approveTimeTrackings(timeTrackings);
            await $(this.repositories.timeTracking.save(timeTrackingsApproved));
            await $(this.sendApprovalMessageToManager(options));
        });
    }

    // Create private methods for long actions so the `execute` method stays readable

    private getTimeTrackings(options: { month: CalendarMonth; managerId: Id }): Async<TimeTracking[]> {
        return this.repositories.timeTracking.get({
            managerId: options.managerId,
            months: [options.month],
        });
    }

    // This could also be in the entity, as a static method
    private approveTimeTrackings(timeTrackings: TimeTracking[]): TimeTracking[] {
        return timeTrackings
            .filter(timeTracking => !timeTracking.isApproved)
            .map(timeTracking => timeTracking.approve());
    }

    private sendApprovalMessageToManager(options: { managerId: Id; month: CalendarMonth }): Async<void> {
        return this.repositories.manager.getById(options.managerId).flatMap(manager => {
            return this.repositories.message.send(
                Message.create({
                    recipients: [manager.email],
                    title: "Time tracking approved",
                    text: `Hello ${manager.name}, your time tracking for have been approved`,
                })
            );
        });
    }
}

// DATA: repository implementations
