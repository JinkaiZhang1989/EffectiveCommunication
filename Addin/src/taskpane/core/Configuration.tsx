export default class Configuration {
    public static epochTicksDiff: number = 621355968000000000; // the number of .net ticks at the unix epoch
    public static ticksPerMillisecond: number = 10000;         // there are 10000 .net ticks per millisecond

    public static oneSecondInMilliseconds: number = 1000;

    public static oneMinuteInSeconds: number = 60;
    public static oneMinuteInMilliSeconds: number = Configuration.oneMinuteInSeconds * Configuration.oneSecondInMilliseconds;

    public static fiveMinutesInMilliSeconds: number = 5 * Configuration.oneMinuteInMilliSeconds;
    public static fifteenMinutesInMilliSeconds: number = 15 * Configuration.oneMinuteInMilliSeconds;
    public static halfHourInMilliSeconds: number = 30 * Configuration.oneMinuteInMilliSeconds;
    public static halfHourInHours: number = 0.5;
    public static threeQuarterHourInHours: number = 0.75;
    public static oneHourInHours: number = 1;
    public static oneHourInMinutes: number = 60;
    public static oneHourInSeconds: number = Configuration.oneHourInMinutes * Configuration.oneMinuteInSeconds;
    public static oneHourInMilliSeconds: number = Configuration.oneHourInSeconds * Configuration.oneSecondInMilliseconds;
    public static twoHoursInMilliSeconds: number = 2 * Configuration.oneHourInMilliSeconds;

    public static oneDayInHours: number = 24;
    public static oneDayInMinutes: number = Configuration.oneDayInHours * Configuration.oneHourInMinutes;
    public static oneDayInSeconds: number = Configuration.oneDayInHours * Configuration.oneHourInSeconds;
    public static oneDay: number = Configuration.oneDayInHours * Configuration.oneHourInMilliSeconds; // 1 day in milliseconds.

    public static oneWeekInDays: number = 7;
    public static oneWeek: number = 7 * Configuration.oneDay;      // 7 days in milliseconds

    public static fourteenDaysInSeconds: number = 14 * Configuration.oneDayInSeconds;
    public static fourteenDays: number = 14 * Configuration.oneDay;    // 14 days in milliseconds

    public static oneMonthInDays: number = 30;
    public static oneMonth: number = 30 * Configuration.oneDay; // 1 month in milliseconds (~30 days)

    public static oneQuarterYear: number = 91 * Configuration.oneDay; // 3 months in milliseconds
    public static oneYear: number = 365 * Configuration.oneDay; // 1 year in milliseconds
    
    // expirable constants
    public static actionQueueFastPollingInterval: number = 6 * Configuration.oneSecondInMilliseconds; // at most 1 flush per 6 seconds.

    public static officejsHasBeenInitialized: boolean = false;

    public static get isCurrentItemReady(): boolean {
        // we cannot refer to anything under Office, since that variable is not declared yet.
        return Configuration.officejsHasBeenInitialized
            && !!Office
            && !!Office.context
            && !!Office.context.mailbox
            && !!Office.context.mailbox.item;
    }

    public static get isInBackgroundMode(): boolean {
        return Configuration.getQueryVariable("extpoint") === "bgaddin";
    }

    private static getQueryVariable(queryKey: string): string {
        const queryString = window.location.search.substring(1);
        if (!!queryString) {
            const variables: string[] = queryString.split("&");
            for (let i = 0; i < variables.length; i ++) {
                const pair: string[] = variables[i].split("=");
                if (pair.length === 2 && pair[0] === queryKey) {
                    return pair[1];
                }
            }
        }

        return null;
    }
}
