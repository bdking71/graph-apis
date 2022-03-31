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
    import { removeOnThemeChangeCallback, ThemeSettingName } from 'office-ui-fabric-react';
    import * as strings from 'GraphApiWebPartStrings';
    import classnames from 'classnames';
    
//#endregion


//#region [Interfaces]

    export interface iState {   
        myURL: string;       
        calenderEvents: any[]; 
        outlookEvents: any[];
    }

//#endregion

export default class GraphApi extends React.Component<IGraphApiProps, iState> {

    //#region [Variables]

        private calendar: any = null;  //* Reference to the calendar control on the page.  

    //#endregion

    //#region [ReactLifeCycleEvents]

        public constructor (props:IGraphApiProps) {
            super(props);     
    
            this.state = { 
                myURL: location.protocol + "//" + location.host + location.pathname,
                calenderEvents: [],  //* Storage of events for the Calendar
                outlookEvents: []  //* Storage of events we've received from MsGraph
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
            }                                    
        }

        public render(): React.ReactElement<IGraphApiProps> {   
            console.log("🚀 ~ file: GraphApi.tsx ~ line 66 ~ GraphApi ~ render ~ this.props", this.props);
                                 
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
                //* Let's hide the calender and display a container for the event.  
                document.getElementById(myCalendarElement).classList.remove(`${styles.show}`);                
                document.getElementById(myCalendarElement).classList.add(`${styles.hide}`);
                document.getElementById(myEventElement).classList.remove(`${styles.hide}`);
                document.getElementById(myEventElement).classList.add(`${styles.show}`);

                //* Let's iterate through state.outlookEvents and see if we can find a matching eventID, and 
                //* use that row of data to display the Event to the user.   
                this.state.outlookEvents.map((item, index) => {                                                 
                    if (item.id == eventID) {
                        console.log("🚀 ~ file: GraphApi.tsx ~ line 90 ~ GraphApi ~ this.state.outlookEvents.map ~ item", item);                    
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
            let RACDate: Date = new Date (RACDateString.substring(0,10) + " " + RACDateString.substring(11,22) + " UTC");                    
            let myMonth: number = RACDate.getMonth() + 1;
            return RACDate.getFullYear().toString() + "-" + 
                    ("0" + myMonth.toString()).slice(-2).toString() + "-" + 
                    ("0" + RACDate.getDate()).slice(-2).toString() + "T" +
                    ("0" + RACDate.getHours()).slice(-2).toString() + ":" +
                    ("0" + RACDate.getMinutes()).slice(-2).toString() + ":" +
                    ("0" + RACDate.getSeconds()).slice(-2).toString() + "+00:00";                                            
        }
            
        private getandProcessSharePointEventData = ():void => {   
            //* Let's get data from SharePoint using the SharePoint PnP modules. 

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

                    client.api(`/groups/${this.props.CalendarCollection[cnt].CalendarGuid}/events`)            
                    .select('subject,body,bodyPreview,organizer,attendees,start,end,location')
                    .get((error, messages: any, rawResponse?: any) => {   
                        if (!messages) {
                            //* Let's log any errors out the console and throw an error back to the caller.
                            console.error(error);   
                            throw error; 
                        } else {   
                            //* Graph returns the data we want in messages.value. We need iterate the array and store 
                            //* the data into a format that our calendar plug-in can understand.  Note: MSGraph returns
                            //* the to and from dates in a odd format. 
                            //BUG: The times that are being pushed here are correct, but aren't being displayed correctly.    
                            messages.value.map((eventItem) => {  
                                let tmp = {
                                    id:  eventItem.id,
                                    title: eventItem.subject,                                    
                                    from: this.RACDateMaker(eventItem.start.dateTime),
                                    to:  this.RACDateMaker(eventItem.end.dateTime),
                                    color: `${this.props.CalendarCollection[cnt].CalendarColor}`
                                }; 
                                //* After reach iteration,  we are going to push the data into the variable we stored
                                //* the current state into.                                               
                                calenderEventsState.push(tmp);    
                                outlookEventsState.push(eventItem);
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