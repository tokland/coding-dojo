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

    /* Collection methods */

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

    execute(): Async<TimeTracking[]> {
        return this.repositories.timeTracking.getAll();
    }
}

class GetManagersUseCase {
    constructor(private managerRepository: ManagerRepository) {}

    execute(managerIds: Id[]): Async<Manager[]> {
        return this.managerRepository.getByIds(managerIds);
    }
}

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
import React from "react";

export class ManagerD2Repository implements ManagerRepository {
    getById(_id: Id): Async<Manager> {
        // Get the tracked entity: api.tracker.trackedEntities.get( trackedEntity.id == id )
        // Map the tracked entity (id, attributes) to a manager object.
        return Future.error(new Error("Not implemented"));
    }

    getByIds(_id: Id[]): Async<Manager[]> {
        // Get the tracked entities: api.tracker.trackedEntities.get( trackedEntity.id IN ids )
        // Map the tracked entities (id, attributes) to manager objects.
        return Future.error(new Error("Not implemented"));
    }
}

export class TimeTrackingD2Repository implements TimeTrackingRepository {
    getAll(): Async<TimeTracking[]> {
        // Get the events for the trackings: api.tracker.events.get( program.id == timeTrackingProgram.id )
        // map the events (which contains data values) to time trackings.
        return Future.error(new Error("Not implemented"));
    }

    save(_timeTrackings: TimeTracking[]): Async<void> {
        // Get events (it may not exist) for the trackings:
        //   api.tracker.events.get( event.id IN timeTracking.ids )
        // Update the existing events with the timeTracking to save, or build for the non-existent.
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

// COMPOSITION ROOT

type CompositionRoot = {
    getTimeTrackingsUseCase: GetTimeTrackingsUseCase;
    getManagersUseCase: GetManagersUseCase;
};

declare const compositionRoot: CompositionRoot;

// VIEW

// Rows containing time trackings with manager information

type TimeTrackingViewModelProps = {
    timeTracking: TimeTracking;
    manager: Manager;
};

export class TimeTrackingViewModel extends Struct<TimeTrackingViewModelProps>() {
    get id(): string {
        return [this.timeTracking.id, this.manager.id].join("-");
    }
}

// VIEW (loading custom hook)

type LoadState<Data> =
    | { type: "loading" } //
    | { type: "loaded"; data: Data }
    | { type: "error"; error: Error };

type TimeTrackingsLoadState = LoadState<{ timeTrackings: TimeTrackingViewModel[] }>;

declare const useCompositionRoot: () => CompositionRoot;

export function useTimeTrackings(): TimeTrackingsLoadState {
    const compositionRoot = useCompositionRoot();
    const [state, setState] = React.useState<TimeTrackingsLoadState>({ type: "loading" });

    React.useEffect(() => {
        return getTimeTrackings(compositionRoot).run(
            timeTrackings => setState({ type: "loaded", data: { timeTrackings: timeTrackings } }),
            error => setState({ type: "error", error: error })
        );
    }, [setState]);

    return state;
}

function getTimeTrackings(compositionRoot: CompositionRoot): Async<TimeTrackingViewModel[]> {
    return compositionRoot.getTimeTrackingsUseCase.execute().flatMap(timeTrackings => {
        const managerIds = TimeTracking.getManagerIds(timeTrackings);

        return compositionRoot.getManagersUseCase.execute(managerIds).map(managers => {
            return buildTimeTrackingViewModels(timeTrackings, managers);
        });
    });
}

function buildTimeTrackingViewModels(
    timeTrackings: TimeTracking[],
    managers: Manager[]
): TimeTrackingViewModel[] {
    const managersById = _(managers).keyBy(manager => manager.id);

    return _(timeTrackings)
        .compactMap(timeTracking => {
            const manager = managersById.get(timeTracking.managerId);

            return manager
                ? TimeTrackingViewModel.create({ timeTracking: timeTracking, manager: manager })
                : undefined; // Skip time trackings without a manager
        })
        .value();
}
