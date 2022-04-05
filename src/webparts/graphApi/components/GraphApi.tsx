//#region [header]
    //[header] @File Name:          GraphApi.tsx
    //[header] @Description:        Retrieves calendar data from msGraph and displays the data using 
    //[header]                      react-awesome-calendar [https://www.npmjs.com/package/react-awesome-calendar] and
    //[header]                      ReactWindow [https://www.npmjs.com/package/reactjs-windows]  
    //[header] @Author:             Bryan King
    //[header] @Date:               March 29, 2022
    //[header] @File Version:       20220328-1243  
//#endregion

//#region [Imports]

    import * as React from 'react';
    import { WebPartContext } from "@microsoft/sp-webpart-base";
    import styles from './GraphApi.module.scss';
    import {IGraphApiProps} from './IGraphApiProps';
    import {escape} from '@microsoft/sp-lodash-subset';
    import {MSGraphClient} from '@microsoft/sp-http';
    import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
    import Calendar from 'react-awesome-calendar';
    import {IACCalendarEvents, IACCalendarEvent} from './IACCalenderTypes';
    import axios from 'axios';
    import * as strings from 'GraphApiWebPartStrings';
    import classnames from 'classnames';
    import { spfi, SPFx } from "@pnp/sp";
    
    //import { graphfi } from "@pnp/graph";
    //import '@pnp/graph/calendars';
    //import '@pnp/graph/users';

    import "@pnp/sp/webs";
    import "@pnp/sp/lists";
    import "@pnp/sp/items";

//#endregion



//#region [Interfaces]

    export interface iState {   
        myURL: string;       
        calenderEvents: any[]; 
        outlookEvents: any[];
        sharePointEvents: any[];
    }

//#endregion

export default class GraphApi extends React.Component<IGraphApiProps, iState> {

    //#region [Variables]

        private calendar: any = null;  //* Reference to the calendar control on the page.  
        private sp;
        private graph;
        private sharepointEventsParser: any = require("sharepoint-events-parser");


    //#endregion

    //#region [ReactLifeCycleEvents]

        public constructor (props:IGraphApiProps) {
            super(props);     
    
            this.state = { 
                myURL: location.protocol + "//" + location.host + location.pathname,
                calenderEvents: [],   //* Storage of events for the Calendar
                outlookEvents: [],    //* Storage of events we've received from MsGraph
                sharePointEvents: []  //* Storage of events we've received from SharePoint Online. 
            };     
            //! I don't believe the "this.calendar=React.creatRef();" is needed, but I thought I would create 
            //! the ref in case there is a future need. 
            this.calendar = React.createRef();   
            
        }
        
        public componentDidMount() {   
            //* Verify there is data in the "CalendarCollection" prop.
            if (this.props.CalendarCollection) {
                //* Once the component completes it's initial mount; let's get outlook data.
                this.getandProcessOutlookEventData();
                this.getandProcessSharePointEventData();
            }                                    
        }

        public render(): React.ReactElement<IGraphApiProps> {   
            console.log("🚀 ~ file: GraphApi.tsx ~ line 72 ~ GraphApi ~ render ~ this.props", this.props);
            console.log("🚀 ~ file: GraphApi.tsx ~ line 73 ~ GraphApi ~ render ~ this.state", this.state);
                        
            let myEventElement: string = "graphEventContainer";
            let myEventDataElement: string = "myEventData";
            let myCalendarElement: string = "graphCalendarContainer";

            const handleButtonClick = () => {
                event.preventDefault();
                //* Let's hide the event and display the calendar.  
                document.getElementById(myCalendarElement).classList.remove(`${styles.hide}`);                
                document.getElementById(myCalendarElement).classList.add(`${styles.show}`);
                document.getElementById(myEventElement).classList.remove(`${styles.show}`);
                document.getElementById(myEventElement).classList.add(`${styles.hide}`);
            };

            const handleChange  = ()  => {
                //* The event.preventDefault() command below keeps the calendar from refreshing after
                //* the onChange event is fired. 
                event.preventDefault();
            };

            const handleClickEvent = (eventID:string) => {            
            
                console.log("🚀 ~ file: GraphApi.tsx ~ line 112 ~ GraphApi ~ handleClickEvent ~ eventID", eventID)
                let eSource:string = this.getEventSource(eventID);    
            

                //* Let's hide the calender and display a container for the event.  
                document.getElementById(myCalendarElement).classList.remove(`${styles.show}`);                
                document.getElementById(myCalendarElement).classList.add(`${styles.hide}`);
                document.getElementById(myEventElement).classList.remove(`${styles.hide}`);
                document.getElementById(myEventElement).classList.add(`${styles.show}`);

                //* Let's iterate through state.outlookEvents and see if we can find a matching eventID, and 
                //* use that row of data to display the Event to the user.   
                this.state.outlookEvents.map((item, index) => {                                                 
                    if (item.id == eventID) {
                        //* Really, can they make working with dates any more of a nightmare.  Ugg. 
                        let startDate:Date = this.SharePointDateMaker(item.start.dateTime);                      
                        let endDate:Date = this.SharePointDateMaker(item.end.dateTime);                                        
                        let myEventData: string = "";                                  
                        myEventData = `<h2>${item.subject}</h2>`;
                        myEventData += `<p><strong>Description:</strong><br><span>${item.body.content}</span></p>`;
                        if (item.location.displayName) {
                            myEventData += `<p><strong>Location: </strong><span>${item.location.displayName}</span></p>`;
                        }                    
                        myEventData += `<p><strong>Start Time: </strong><span>${startDate.toString()}</span></p>`;                        
                        myEventData += `<p><strong>End Time: </strong><span>${endDate.toString()}</span></p>`;                        
                        document.getElementById(myEventDataElement).innerHTML = myEventData;
                    }
                });     
            };
            //* Let's check to see if there are any events in the state to display.  If not, we
            //* we will show an empty calendar. 
            if (this.state.calenderEvents.length !== 0) {         
                return (
                    <div className={styles.graphApi}>                                      
                        <section id={myEventElement} className={classnames(styles.hide, styles.event)}>
                            <span id={myEventDataElement}></span>
                            <button className={styles.myButton} onClick={() => handleButtonClick()}>Back to calendar</button>    
                        </section>  
                        <section id={myCalendarElement} className={classnames(styles.show, styles.event)}>
                            <h2>{escape(this.props.GroupCalendarName)}</h2>
                            <Calendar  
                                ref={this.calendar} 
                                calendarType="US"
                                defaultView="month" 
                                events={this.state.calenderEvents}  
                                onChange={(event) => handleChange()}
                                onClickEvent={(event) => handleClickEvent(event)} />
                        </section>    
                    </div> );
            } else {                  
                return (<Calendar   /> );   
            }
        }

    //#endregion
  
    //#region [PrivateMethods]

        //* We need a method that will return the source of the event (eID) provided.  The return value 
        //* should be either "SPO" or "Outlook."  Later versions may expand by added more sources. 
        private getEventSource = (eID: string):string => {
            //* To be safe, we are going to verify we have events on the calendar.  I feel like I am
            //* checking both directions on an one direction road. 
            if (this.state.calenderEvents.length !== 0) {
                //* lets iterate the calendar collection until we find the row that matches the eID var. 
                this.state.calenderEvents.map((item, index) => {
                    if (item.id == eID) {
                        return item.origin.toString();
                    }                      
                 });
            }
            return null;
        };

        private SharePointDateMaker = (spDateString: string):Date => {
            //* The format that SP is giving us is "2022-04-01T20:30:00.0000000" without the UTC.  This bit of information is 
            //* one another line in the JSON.  Here we are trying to fix the date and time in UTC so we can convert it back
            //* to local time, IE: Tue Apr 05 2022 16:30:00 GMT-0400 (Eastern Daylight Time).  
            let retVal = new Date(spDateString.substring(0,10) + " " + spDateString.substring(11,22) + " UTC");                
            return retVal;
        }

        //* Really can JavaScript make date functions any worse? This function is a bit Kludgy, but it works! 
        //* Anyway, we need to build a date string for the calendar and force the React-Awesome-Calendar display 
        //* the event on the date time we provide.  We do so by creating the string this way: "YYYY-MM-DDTHH:MM:SS+00.00".  
        private RACDateMaker = (RACDateString: string): string => {            
            //BUGFIX: I had this working; however, when I checked moved the code base over to 
            //BUGFIX: an Ubuntu 21 machines, it started throwing invalid dates errors.  I am 
            //BUGFIX: working if the "Invalid Date" error is a issue with the Firefox 98.0.2? 
            let RACDate: Date = new Date (RACDateString.substring(0,10) + "T" + RACDateString.substring(11,16) + "Z");                    
            let myMonth: number = RACDate.getMonth() + 1;
            let retval: string = RACDate.getFullYear().toString() + "-" + 
                    ("0" + myMonth.toString()).slice(-2).toString() + "-" + 
                    ("0" + RACDate.getDate()).slice(-2).toString() + "T" +
                    ("0" + RACDate.getHours()).slice(-2).toString() + ":" +
                    ("0" + RACDate.getMinutes()).slice(-2).toString() + ":" +
                    ("0" + RACDate.getSeconds()).slice(-2).toString() + "+00:00";                                            
            return retval;
        }
            
        private getandProcessSharePointEventData = ():void => {  
            let items;             
            //* Let's get data from SharePoint using the SharePoint PnP modules. First, let's 
            //* verify we have connection data for any SharePoint Calendar.
            if (this.props.SharePointCalendarCollection.length !== 0) { 
                    //* let's make a copy of the current calender entries stored in the calendarEvents and the
                    //* sharePointEvents state so, we don't loose them when added new entries.  
                    let calenderEventsState: any[] = this.state.calenderEvents;
                    let sharePointEventsState: any[] = this.state.sharePointEvents;
                    
                //* Let's iterate through the array that contains the SharePoint connection info 
                for (let cnt:number = 0; cnt <= (this.props.SharePointCalendarCollection.length - 1); cnt ++) {
                    this.sp = spfi(this.props.SharePointCalendarCollection[cnt].SharePointCalendarSiteUrl).using(SPFx(this.props.Context));
                    this.sp.web.lists
                        .getByTitle(this.props.SharePointCalendarCollection[cnt].SharePointCalendarName)
                        .items()
                        .then((spEvents) => {
                            spEvents.map((spEvent) => {  
                                //* SharePoint stores recurring events as a single event in the calendar. It doesn't
                                //* look as if there is a good way to have SharePoint return the event expanded. So, we 
                                //* will need to do this the hard way.  First, we will check to see if the
                                //* current event in the loop is a recurring event (fRecurrence = true).  
                                if (spEvent.fRecurrence == true) {
                                    //* This event is part of the recurring event; The way to get the data from REST is: 
                                    //* [URL]/_api/web/lists/getByTitle('[Calendar Name]')/items([ID])/RecurrenceData. Since,
                                    //* I couldn't find a pnp plugin that would provide us with the recurrence data; therefore, 
                                    //* I will use Axios to preform a rest query.   
                                    let restQuery = `${this.props.SharePointCalendarCollection[cnt].SharePointCalendarSiteUrl}_api/web/lists/getByTitle('${this.props.SharePointCalendarCollection[cnt].SharePointCalendarName}')/items(${spEvent.ID})?$select=Duration,RecurrenceData,MasterSeriesItemID,EventType,*`;
                                    axios({method: 'get', url: restQuery, responseType: 'json'}).then(recurrenceInfo =>{                                        
                                        let parsedArray = this.sharepointEventsParser.parseEvent(recurrenceInfo.data);
                                        parsedArray.map((parsedEvent) => {
                                            let tmp = {
                                                id:parsedEvent.GUID,
                                                title:parsedEvent.Title,  
                                                from: parsedEvent.EventDate,
                                                to:  parsedEvent.EndDate,
                                                color: `${this.props.SharePointCalendarCollection[cnt].SharePointCalendarColor}`,
                                                origin: "SPO" //* This variable will tell the app where to look for the long Description of the Event.
                                            }; 
                                            //* After reach iteration,  we are going to push the data into the variable we stored
                                            //* the current state into.                                               
                                            calenderEventsState.push(tmp);    
                                            sharePointEventsState.push(parsedEvent);
                                        });
                                        
                                    });    
                                    // https://www.npmjs.com/package/sharepoint-events-parser
                                } else {
                                    let tmp = {
                                        id:  spEvent.Id,
                                        title: spEvent.Title,                                    
                                        from: spEvent.EventDate,
                                        to:  spEvent.EndDate,
                                        color: `${this.props.SharePointCalendarCollection[cnt].SharePointCalendarColor}`,
                                        origin: "SPO" //* This variable will tell the app where to look for the long Description of the Event.
                                    }; 
                                    //* After reach iteration,  we are going to push the data into the variable we stored
                                    //* the current state into.                                               
                                    calenderEventsState.push(tmp);    
                                    sharePointEventsState.push(spEvent);
                                }                            
                        });
                        //* Finally, we are going to set both the calendarEvents and the outlookEvents state with the 
                        //* events have in this iteration of the Calendar Collection.
                        this.setState({calenderEvents:calenderEventsState, sharePointEvents:sharePointEventsState});  
                    });
                }
            }
        }

        private getandProcessOutlookEventData = ():void => {          
            //* We will iterate through the Calendar Collection in order to pull data from MSGraph 
            //* for each calendar the user defined in the webpart props.            
            for (let cnt = 0; cnt <= (this.props.CalendarCollection.length - 1); cnt ++) {      
                //* let's makes a connection to the MSGraph...
                this.props.Context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {   
                    //* let's make a copy of the current calender entries stored in the calendarEvents and the 
                    //* outlookEvent state so, we don't loose them when added new entries.  
                    let calenderEventsState: any[] = this.state.calenderEvents;
                    let outlookEventsState: any[] = this.state.outlookEvents; 
                    //BUGFIX: [ID: [202204040921]Recurring Events stored in Outlook are not displaying all events.  Currently, it
                    //BUGFix: [ID: [202204040921]is showing only the first event in the series. 
                    client.api(`/groups/${this.props.CalendarCollection[cnt].CalendarGuid}/events`)            
                    .select('*')
                    .get((error, messages: any, rawResponse?: any) => {   
                        if (!messages) {
                            //* Let's log any errors out the console and throw an error back to the caller.
                            console.error(error);   
                            throw error; 
                        } else {   
                            let tmp; 
                            //* Graph returns the data we want in messages.value. We need iterate the array and store 
                            //* the data into a format that our calendar plug-in can understand.  Note: MSGraph returns
                            //* the to and from dates in a odd format. 
                            messages.value.map((eventItem) => {     
                                //* Let's check to see if this event is recurring.                                
                                if (eventItem.recurrence != null) {
                                    //* Let's make a second call the graph api, asking for all the instances of the
                                    //* current event.

                                    //https://docs.microsoft.com/en-us/graph/api/event-list-instances?view=graph-rest-1.0&tabs=http
                                    client
                                        .api(`/groups/${this.props.CalendarCollection[cnt].CalendarGuid}/calendar/events/${eventItem.id}/instances?startDateTime=2022-01-01&endDateTime=2023-01-01`)                                        
                                        .get((error0, messages0: any, rawResponse0?: any) => {
                                            messages0.value.map((OutLookRecurringEventItem) => {
                                                tmp = {
                                                    id:  OutLookRecurringEventItem.id,
                                                    title: OutLookRecurringEventItem.subject,                                    
                                                    from: this.RACDateMaker(OutLookRecurringEventItem.start.dateTime),
                                                    to:  this.RACDateMaker(OutLookRecurringEventItem.end.dateTime),
                                                    color: `${this.props.CalendarCollection[cnt].CalendarColor}`,
                                                    origin: "Outlook" //* This variable will tell the app where to look for the long Description of the Event.
                                                };                                                  
                                                calenderEventsState.push(tmp);    
                                                outlookEventsState.push(OutLookRecurringEventItem); 
                                            });                                            
                                    });                                                                       
                                } else { 
                                    tmp = {
                                        id:  eventItem.id,
                                        title: eventItem.subject,                                    
                                        from: this.RACDateMaker(eventItem.start.dateTime),
                                        to:  this.RACDateMaker(eventItem.end.dateTime),
                                        color: `${this.props.CalendarCollection[cnt].CalendarColor}`,
                                        origin: "Outlook" //* This variable will tell the app where to look for the long Description of the Event.
                                    };                                 
                                    //* After reach iteration,  we are going to push the data into the variable we stored
                                    //* the current state into.                                               
                                    calenderEventsState.push(tmp);    
                                    outlookEventsState.push(eventItem);
                                }
                            }); 
                        }                
                        //* Finally, we are going to set both the calendarEvents and the outlookEvents state with the 
                        //* events have in this iteration of the Calendar Collection.
                        this.setState({calenderEvents:calenderEventsState, outlookEvents:outlookEventsState});  
                    });                  
                }); 
            }     
        }

    //#endregion
}