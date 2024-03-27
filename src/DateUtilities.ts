
export function dayOfYear(date: Date) {
    return Math.floor((date.getTime() - new Date(date.getFullYear(), 0, 0).getTime()) / 1000 / 3600 / 24);
}

export function addDays(date: Date, daysToAdd: number) {
    const millisecondsInDay = 86_400_000;
    return new Date(date.getTime() + millisecondsInDay * daysToAdd);
}

export class TimeSpan {
    protected static readonly millisecondsInSecond = 60_000;
    protected static readonly millisecondsInMinute = 60_000;
    protected static readonly millisecondsInHour = 3_600_000;
    protected static readonly millisecondsInDay = 86_400_000;

    static fromSeconds(seconds: number) {
        return new TimeSpan(seconds * this.millisecondsInSecond);
    }

    static fromMinutes(minutes: number) {
        return new TimeSpan(minutes * this.millisecondsInMinute);
    }

    static fromHours(hours: number) {
        return new TimeSpan(hours * this.millisecondsInHour);
    }

    static fromDays(days: number) {
        return new TimeSpan(days * this.millisecondsInDay);
    }

    constructor(private ms = 0) { }

    toMilliseconds() {
        return this.ms;
    }

    toSeconds() {
        return Math.floor(this.ms / TimeSpan.millisecondsInSecond);
    }

    toMinutes() {
        return Math.floor(this.ms / TimeSpan.millisecondsInMinute);
    }

    toHours() {
        return Math.floor(this.ms / TimeSpan.millisecondsInHour);
    }

    toDays() {
        return Math.floor(this.ms / TimeSpan.millisecondsInDay);
    }
}