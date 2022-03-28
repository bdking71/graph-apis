//#region [header]
    //[header] @File Name:          GraphApi.tsx
    //[header] @Description:        Retreives calendar data from msGraph and displays the data using 
    //[header]                      react-awesome-calendar [https://www.npmjs.com/package/react-awesome-calendar] and
    //[header]                      ReactWindow [https://www.npmjs.com/package/reactjs-windows]  
    //[header] @Author:             Bryan King
    //[header] @Date:               March 29, 2022
    //[header] @File Version:       20220328-1243  
//#endregion

//#region [Imports]

    import * as React from 'react';
    import styles from './GraphApi.module.scss';
    import {IGraphApiProps} from './IGraphApiProps';
    import {escape} from '@microsoft/sp-lodash-subset';
    import {MSGraphClient} from '@microsoft/sp-http';
    import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
    import Calendar from 'react-awesome-calendar';
    import {IACCalendarEvents, IACCalendarEvent} from './IACCalenderTypes';
    import { removeOnThemeChangeCallback, ThemeSettingName } from 'office-ui-fabric-react';
import * as strings from 'GraphApiWebPartStrings';
    
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
            //* Once the component completes it's intital mount; let's get some data.      
            this.getandProcessEventData();        
        }

        public render(): React.ReactElement<IGraphApiProps> {            
            console.log("🚀 ~ file: GraphApi.tsx ~ line 65 ~ GraphApi ~ render ~ this.state", this.state);
            let myHiddenDiv: string = "graphEventContainer";
            const handleChange  = ()  => {
                //* The event.preventDefault() command below keeps the calendar from refreshing after
                //* the onChange event is fired. 
                event.preventDefault();
            };
            const handleClickEvent = (eventID:string) => {              
                //* Let's iterate through state.outlookEvents and see if we can find a matching eventID, and 
                //* use that row of data to display the Event to the user.   
                this.state.outlookEvents.map((item, index) => {   
                    //BUG: The item is comming equal this.state.outlookEvents instead of a single array from the 
                    //BUG: this.state.outlookEvents. For now, I will use item[index] to reference the row, until I 
                    //BUG: can debug why this is happening.                                      
                    if (item[index].id == eventID) {
                        let myEventData: string = "";                        
                        let myDiv: HTMLElement = document.getElementById(myHiddenDiv);
                        
                        myEventData = `<h1>${item[index].subject}</h1>`;
                        myEventData += `<span>${item[index].body.content}</span>`;
                        myEventData += `<span>${item[index].location.displayName}</span>`;

                        myDiv.innerHTML = myEventData;
                        console.log("🚀 ~ file: GraphApi.tsx ~ line 87 ~ GraphApi ~ this.state.outlookEvents.map ~ myEventData", myEventData)
                    }
                });     
            };

            //* Let's check to see if there are any events in the state to display.  If not, we
            //* we will show an empty calendar. 
            if (this.state.calenderEvents.length !== 0) {         
                return (
                    <section className={styles.graphApi}>
                        <div id={myHiddenDiv} className={styles.floated}>Blah, Blah, Blah!</div>
                        <h2>{escape(this.props.GroupCalendarName)}</h2>
                        <Calendar  
                            ref={this.calendar} 
                            calendarType="US"
                            defaultView="month" 
                            events={this.state.calenderEvents}  
                            onChange={(event) => handleChange()}
                            onClickEvent={(event) => handleClickEvent(event)} />
                    </section> );
            } else {                  
                return (<Calendar   /> );   
            }
        }

    //#endregion
  
    //#region [PrivateMethods]

        private getandProcessEventData = ():void => {        
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
                            //* the data into a format that our calendar plug-in can understand.  
                            messages.value.map((eventItem) => {  
                                let tmp = {
                                    id:  eventItem.id,
                                    title: eventItem.subject,
                                    to:  new Date(eventItem.end.dateTime).toLocaleDateString(),
                                    from: new Date(eventItem.start.dateTime).toLocaleDateString(),
                                    color: `${this.props.CalendarCollection[cnt].CalendarColor}`
                                }; 
                                //* After reach iteration,  we are going to push the data into the variable we stored
                                //* the current state into.                                               
                                calenderEventsState.push(tmp);    
                                outlookEventsState.push(messages.value);
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