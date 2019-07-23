import Configuration from "./Configuration";

export enum ExpirationTypes {
    ByEvery_Minutes,
    BySpecificDateTime,
    ByRecurrenceOnEvery_,
    ByItemChange
}

export enum RecurrenceTypes {// local datetime only
    Evening5pm,
    Morning10am,
    MondayMorning9am,
    FridayAfternoon1pm,
    NextDay10am
}

export enum ExpirableInBackground {
    Allowed,
    Suppressed
}

export interface IExpirationAction {
    actionCreatorFunction?: Function;
}

export interface IExpirationSetting {
    expirationType: ExpirationTypes;
    value: number | Date | RecurrenceTypes; // # of minutes, Datetime value, RecurrentTypes
    expirableInBackground: ExpirableInBackground;
}

export interface IExpirationBase {
    setting: IExpirationSetting;
    action: IExpirationAction;
    exprationInUTC: number;  // for all expirabletypes, eventually resolve to next expiration time
    usedSetInterval?: boolean;
    target?: any;
}

// global actions queue
class ActionsQueue {
    private flushMinInterval: number = Configuration.actionQueueFastPollingInterval; // at most 1 flush per 15 seconds.
    private lastFlushingTime: number;
    private actionQueue: IExpirationBase[];
    private expirablesMonitoringArray: IExpirationBase[];
    private static instance: ActionsQueue;

    private constructor() {
        this.lastFlushingTime = Date.now();
        this.actionQueue = [];
        this.expirablesMonitoringArray = [];

        const checkOfficeJSReadyInterval = window.setInterval(
            () => {
                if (Configuration.isCurrentItemReady) {
                    // start the execution engine timer
                    window.setInterval(this.timedPolling.bind(this), this.flushMinInterval);
                    // Clear the interval
                    window.clearInterval(checkOfficeJSReadyInterval);
                }
            },
            10);
    }

    public static getInstance() {
        if (!ActionsQueue.instance) {
            ActionsQueue.instance = new ActionsQueue();
        }

        return ActionsQueue.instance;
    }

    private timedPolling(): void {
        if (Configuration.isCurrentItemReady) {
            this.timedQueuing();
            this.timedExecution();
        }
    }

    private timedQueuing(): void {
        const nowInUTC = ExpirableUtils.nowInUTC();
        const newExpirablesArray: IExpirationBase[] = [];
        this.expirablesMonitoringArray.map(
            (value: IExpirationBase, index: number, arr: IExpirationBase[]) => {
                if (!!value.exprationInUTC && value.exprationInUTC < nowInUTC) {
                    newExpirablesArray.push(ExpirableUtils.ExpireMeAndReschedule(value));
                } else {
                    newExpirablesArray.push(value);
                }
            }
        );

        this.expirablesMonitoringArray = newExpirablesArray;
    }

    private timedExecution(): void {
        if (this.shouldFlushActionQueue()) {
            this.tryExecuteQueuedActions();
        }
    }

    private shouldFlushActionQueue(): boolean {
        return (Date.now() - this.lastFlushingTime) > this.flushMinInterval;
    }

    private compareActions(aExpirable: IExpirationBase, bExpirable: IExpirationBase): number {
        const a: IExpirationAction = aExpirable.action;
        const b: IExpirationAction = bExpirable.action;

        // a simple comparison of length to mimic function comparison.
        return (String(a.actionCreatorFunction).length - String(b.actionCreatorFunction).length);
    }

    public addToExpirablesMonitoringArray(expirable: IExpirationBase): void {
        this.expirablesMonitoringArray.push(expirable);
    }

    public enqueue(expirable: IExpirationBase): void {
        this.actionQueue.push(expirable);
    }

    public tryExecuteQueuedActions(): boolean {
        let dedupedActionQueue: IExpirationBase[] = [];
        this.actionQueue.sort(this.compareActions);
        dedupedActionQueue = this.actionQueue.filter((value: IExpirationBase, index: number, arr: IExpirationBase[]) => {
            if (index < 1) {
                return true;
            } else {
                return this.compareActions(value, arr[index - 1]) !== 0;
            }
        });

        dedupedActionQueue.forEach(a => {
            a.action.actionCreatorFunction.apply(a.target, []);
        });

        this.lastFlushingTime = Date.now();
        this.actionQueue = [];
        return true;
    }
}

class ExpirableUtils {
    public static nowInUTC(): number {
        return new Date(Date.now()).getTime();
    }

    public static IsExpired<T extends IExpirationBase>(expirable: T): boolean {
        const nowInUTC: number = new Date(Date.now()).getTime();
        return nowInUTC > expirable.exprationInUTC;
    }

    public static setDay(date: Date, dayOfWeek: number): Date {
        if (dayOfWeek >= 0 && dayOfWeek <= 6) {
            date = new Date(date.getTime());
            date.setDate(date.getDate() + (dayOfWeek + 7 - date.getDay()) % 7);
        }

        return date;
    }

    public static nextTimeToRecur(recurrenceType: RecurrenceTypes): number {
        let nextTimeInUTC: number = 0;
        const today = new Date();
        switch (recurrenceType) {
            case RecurrenceTypes.NextDay10am:
                nextTimeInUTC = new Date(today.getTime() + Configuration.oneDay).setHours(10, 0, 0, 0);
                break;
            case RecurrenceTypes.Morning10am:
                let morning10am = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 10, 0, 0, 0);
                if (today > morning10am) {
                    morning10am = new Date(morning10am.getTime() + Configuration.oneDay);
                }

                nextTimeInUTC = morning10am.getTime();
                break;
            case RecurrenceTypes.Evening5pm:
                let evening5pm = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 17, 0, 0, 0);
                if (today > evening5pm) {
                    evening5pm = new Date(evening5pm.getTime() + Configuration.oneDay);
                }

                nextTimeInUTC = evening5pm.getTime();
                break;
            case RecurrenceTypes.FridayAfternoon1pm:
                let fridayOfThisWeek = ExpirableUtils.setDay(today, 5); // set to Friday
                fridayOfThisWeek = new Date(fridayOfThisWeek.getFullYear(), fridayOfThisWeek.getMonth(), fridayOfThisWeek.getDate(), 13, 0, 0, 0);
                if (today > fridayOfThisWeek) {
                    fridayOfThisWeek = new Date(fridayOfThisWeek.getTime() + Configuration.oneWeek);
                }

                nextTimeInUTC = fridayOfThisWeek.getTime();
                break;
            case RecurrenceTypes.MondayMorning9am:
                let mondayOfThisWeek = ExpirableUtils.setDay(today, 1); // set to Monday
                mondayOfThisWeek = new Date(mondayOfThisWeek.getFullYear(), mondayOfThisWeek.getMonth(), mondayOfThisWeek.getDate(), 9, 0, 0, 0);
                if (today > mondayOfThisWeek) {
                    mondayOfThisWeek = new Date(mondayOfThisWeek.getTime() + Configuration.oneWeek);
                }

                nextTimeInUTC = mondayOfThisWeek.getTime();
                break;
            default:
                break;
        }

        return nextTimeInUTC;
    }

    // Trigger actions
    public static QueueActions<T extends IExpirationBase>(expirable: T): T {
        // at this point, action Queue must have established.
        const actionsQueue: ActionsQueue = ActionsQueue.getInstance();
        actionsQueue.enqueue(expirable);
        return expirable;
    }

    // schedule the next expiration
    public static Schedule<T extends IExpirationBase>(expirable: T): T {
        let timeoutInMilliSeconds: number = 0;
        const millisecondsInOneMinute = Configuration.oneMinuteInMilliSeconds; //  Constants.fastPaceOneMinute_ForTesting;
        const nowInUTC: number = ExpirableUtils.nowInUTC();
        switch (expirable.setting.expirationType) {
            case ExpirationTypes.ByEvery_Minutes:
            case ExpirationTypes.ByItemChange:
                timeoutInMilliSeconds = (expirable.setting.value as number) * millisecondsInOneMinute;
                break;
            case ExpirationTypes.BySpecificDateTime:
                const expireTimeInUTC: number = (expirable.setting.value as Date).getTime();
                timeoutInMilliSeconds = expireTimeInUTC - nowInUTC;
                break;
            case ExpirationTypes.ByRecurrenceOnEvery_:
                timeoutInMilliSeconds = ExpirableUtils.nextTimeToRecur(expirable.setting.value as RecurrenceTypes) - nowInUTC;
                break;
            default:
                break;
        }

        if (timeoutInMilliSeconds > 0) {
            expirable.exprationInUTC = nowInUTC + timeoutInMilliSeconds;
            if (!!expirable.usedSetInterval) {
                return expirable;
            }

            if (expirable.setting.expirationType === ExpirationTypes.ByEvery_Minutes
                || expirable.setting.expirationType === ExpirationTypes.ByRecurrenceOnEvery_) {
                expirable.usedSetInterval = true;
                //caller needs to put the expirable object returned in this case into ActionQueue monitoring array
                // different cases might require adding to different arrays. caller has the info.
                // queueing logic also shouldn't be added into this class.
            } else {
                window.setTimeout(ExpirableUtils.ExpireMeAndReschedule, timeoutInMilliSeconds, expirable);
            }
        } else {
            // if it's itemchange, register to itemchange event for this creator function
            if (expirable.setting.expirationType === ExpirationTypes.ByItemChange) {
                ExpirableUtils.registerWithItemChangedEvent(expirable);
            }
        }

        return expirable;
    }

    public static registerWithItemChangedEvent<T extends IExpirationBase>(expirable: T) {
        const MAX_RETRY_COUNT: number = 10;
        let retryCount = 0;
        const timeoutId = window.setInterval(
            (expirable: T) => {
                if (Configuration.isCurrentItemReady && !!Office.context.mailbox.addHandlerAsync) {
                    Office.context.mailbox.addHandlerAsync(
                        Office.EventType.ItemChanged,
                        (eventArgs) => {
                            // This is the event handler.

                            if (!eventArgs.initialData) {
                                return; // don't refresh when folder change. item change will be issued again when the item is selected in new folder.
                            }

                            expirable.action.actionCreatorFunction.apply(expirable.target, []);
                        },
                        (asyncResult) => {
                            // This is the event registration handler. This will be called ONCE when the event is registered.
                            // remove the interval
                            window.clearInterval(timeoutId);
                        });
                } else {
                    if (retryCount <= MAX_RETRY_COUNT) {
                        retryCount++;
                    } else {
                        window.clearInterval(timeoutId);
                    }
                }
            },
            3000,
            expirable);
    }

    // expire the current object
    // depending on type of expiration, set another expiration point,
    // return timeoutsetting object to caller, so it can be updated if needed and save back to data array.
    public static ExpireMeAndReschedule<T extends IExpirationBase>(expirable: T): T {

        if (Configuration.isCurrentItemReady &&
            (!Configuration.isInBackgroundMode || expirable.setting.expirableInBackground === ExpirableInBackground.Allowed) &&
            ExpirableUtils.IsExpired(expirable)) {
            // Queuing actions
            try {
                expirable = ExpirableUtils.QueueActions(expirable);
            } catch (e) {
                // no-op
            }
        }

        // schedule again if necessary
        expirable = ExpirableUtils.Schedule(expirable);

        return expirable;
    }
}

// decorate the class
export function expirable(): Function;

// decorate the creator with other expiration types.
// enforcing type and val consistency.
// any new types defined in ExpriationTypes would need to add a new function def here.
export function expirable(
    expirationType: ExpirationTypes.ByEvery_Minutes,
    settingVal: number,
    expirableInBackground?: ExpirableInBackground): Function;

export function expirable(
    expirationType: ExpirationTypes.ByRecurrenceOnEvery_,
    settingVal: RecurrenceTypes,
    expirableInBackground?: ExpirableInBackground): Function;

export function expirable(
    expirationType: ExpirationTypes.BySpecificDateTime,
    settingVal: Date,
    expirableInBackground?: ExpirableInBackground): Function;

// decorate the creator with ExpirationTypes.ByItemChange type, without settingVal
export function expirable(expirationType: ExpirationTypes.ByItemChange);

export function expirable(
    expirationType?: ExpirationTypes,
    settingVal?: number | Date | RecurrenceTypes,
    expirableInBackground: ExpirableInBackground = ExpirableInBackground.Suppressed): Function {

    // set the expiration number to -1 for itemchange type, so it will fire as soon as item is changed.
    if (!settingVal && !expirationType && expirationType === ExpirationTypes.ByItemChange) {
        settingVal = -1;
    }

    return (target?: any, propertyKey?: string | symbol, descriptor?: TypedPropertyDescriptor<any>): TypedPropertyDescriptor<any> | void | Function => {

        if (!target) {
            // We're being used to declare the parameter
            return null;
        }

        // decorator on class
        // change constructor here.
        if (!propertyKey) {
/* tslint:disable:no-invalid-this */
            const wrapper = function (...args: any[]) {
                target.apply(this, args);
                // update expirable objects' references to target for this class instance.
                if (this.$__Expirables__$) {
                    for (const key of Object.keys(this.$__Expirables__$)) {
                        this.$__Expirables__$[key].target = this;
                    }
                }

                return this;
            };
/* tslint:enable:no-invalid-this */

            const prototype = wrapper.prototype = target.prototype;

            if (prototype.$__Expirables__$) {
                for (const key of Object.keys(prototype.$__Expirables__$)) {
                    // first schedule the call regardless it's in current workflow or not.
                    prototype.$__Expirables__$[key] = ExpirableUtils.Schedule(prototype.$__Expirables__$[key]);
                    // initial queue to expirable array in actionqueue.
                    ActionsQueue.getInstance().addToExpirablesMonitoringArray(prototype.$__Expirables__$[key]);
                }
            }

            return wrapper;
        }

        // allocate expirable object array.
        if (!target.$__Expirables__$) {
            target.$__Expirables__$ = [];
        }

        // initialize expirable object
        const expirable: IExpirationBase = {
            action: {
                actionCreatorFunction: descriptor.value
            },
            setting: {
                expirationType: expirationType,
                value: settingVal,
                expirableInBackground: expirableInBackground
            },
            exprationInUTC: 0
        };

        expirable.target = target;
        target.$__Expirables__$.push(expirable);
    };
}
