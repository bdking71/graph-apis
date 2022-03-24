/* Define the interface to hold the events pulled for the Calendar used on this webpart */

export interface IACCalendarEvents {
    value: IACCalendarEvent[];
}

export interface IACCalendarEvent {
    id: any;
    color: string;
    from: Date;
    to: Date;
    title: string;
}