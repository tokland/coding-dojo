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

    static forManager(timeTrackings: TimeTracking[], managerId: Id): TimeTracking[] {
        return timeTrackings.filter(timeTracking => timeTracking.managerId === managerId);
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

class Message extends Struct<{
    recipients: string[];
    title: string;
    text: string;
}>() {}

type Id = string;

type Async<Data> = Future<Error, Data>;

// DOMAIN: repositories

interface ManagerRepository {
    getById(id: Id): Async<Manager>;
    getByIds(id: Id[]): Async<Manager[]>;
}

interface TimeTrackingRepository {
    getAll(): Async<TimeTracking[]>;
    save(timeTrackings: TimeTracking[]): Async<void>;
}

interface MessageRepository {
    send(message: Message): Async<void>;
}

// DOMAIN: use cases

export class GetTimeTrackingsUseCase {
    constructor(private repositories: { timeTracking: TimeTrackingRepository; manager: ManagerRepository }) {}

    execute(): Async<TimeTrackingAndManager[]> {
        return this.repositories.timeTracking.getAll().flatMap(timeTrackings => {
            return this.joinTrackingsAndManagers(timeTrackings);
        });
    }

    private joinTrackingsAndManagers(timeTrackings: TimeTracking[]): Async<TimeTrackingAndManager[]> {
        const managerIds = _(timeTrackings)
            .map(timeTracking => timeTracking.managerId)
            .uniq()
            .value();

        return this.repositories.manager.getByIds(managerIds).map(managers => {
            const managersById = _(managers).keyBy(manager => manager.id);

            return _(timeTrackings)
                .compactMap(timeTracking => {
                    const manager = managersById.get(timeTracking.managerId);
                    return manager ? { timeTracking, manager } : undefined;
                })
                .value();
        });
    }
}

type TimeTrackingAndManager = {
    timeTracking: TimeTracking;
    manager: Manager;
};

export class ApproveTimeTrackingsUseCase {
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

    // Demo: we can use promise-like syntax with Future.block
    executeUsingFutureBlock(timeTrackings: TimeTracking[]): Async<void> {
        return Future.block(async $ => {
            const timeTrackingsApproved = TimeTracking.approve(timeTrackings);
            await $(this.repositories.timeTracking.save(timeTrackingsApproved));
            await $(this.notifyManagers(timeTrackings));
        });
    }

    // Create private methods for actions so the `execute` method stays readable

    private notifyManagers(timeTrackings: TimeTracking[]): Async<void> {
        // Future.block is nice for loops of async operations (but it forces sequential execution)
        return Future.block(async $ => {
            for (const managerId of TimeTracking.getManagerIds(timeTrackings)) {
                const manager = await $(this.repositories.manager.getById(managerId));
                const timeTrackingsForManager = TimeTracking.forManager(timeTrackings, managerId);
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

// DATA: repository implementations (DHIS2)

import { D2Api } from "./types/d2-api";
import { apiToFuture } from "./data/api-futures";

export class ManagerD2Repository implements ManagerRepository {
    getById(_id: Id): Async<Manager> {
        // Get the tracked entity: api.tracker.trackedEntities.get( trackedEntity.id == id )
        // Map the tracked entity (id, attributes) to a manager object.
        return Future.error(new Error("Not implemented"));
    }

    getByIds(id: Id[]): Async<Manager[]> {
        // Get the tracked entities: api.tracker.trackedEntities.get( trackedEntity.id IN ids )
        // Map the tracked entities (id, attributes) to manager objects.
        return Future.error(new Error("Not implemented"));
    }
}

export class TimeTrackingD2Repository implements TimeTrackingRepository {
    getAll(): Async<TimeTracking[]> {
        // Get the events for the time trackings:
        //   api.tracker.events.get( program.id == timeTrackingProgram.id )
        // map the events (which contains data values) to time trackings.
        return Future.error(new Error("Not implemented"));
    }

    save(_timeTrackings: TimeTracking[]): Async<void> {
        // Get events (it may not exist) for the trackings:
        //   api.tracker.events.get( event.id IN timeTracking.ids )
        // Update the existing events with the timeTracking to save, and build from zero for the non-existent.
        // Post the updated+new events: api.tracker.post({ events })
        return Future.error(new Error("Not implemented"));
    }
}

// This one is very simple, let's fully implement it
export class MessageD2Repository implements MessageRepository {
    constructor(private api: D2Api) {}

    send(message: Message): Async<void> {
        return apiToFuture(
            this.api.email.sendMessage({
                recipients: message.recipients,
                subject: message.title,
                text: message.text,
            })
        );
    }
}
