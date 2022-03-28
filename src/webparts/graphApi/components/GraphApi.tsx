//#region [Imports]

    import * as React from 'react';
    import styles from './GraphApi.module.scss';
    import {IGraphApiProps} from './IGraphApiProps';
    import {escape} from '@microsoft/sp-lodash-subset';
    import {MSGraphClient} from '@microsoft/sp-http';
    import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
    import Calendar from 'react-awesome-calendar';
    import {IACCalendarEvents, IACCalendarEvent} from './IACCalenderTypes';

//#endregion

//#region [Interfaces]

    export interface iState {   
        myURL: string;       
        calenderEvents: any[]; 
    }

//#endregion

export default class GraphApi extends React.Component<IGraphApiProps, iState> {

private calendar: any = null;  

    //#region [ReactLifeCycleEvents]

        public constructor (props:IGraphApiProps) {
            super(props);         
            this.state = { 
                myURL: location.protocol + "//" + location.host + location.pathname,
                calenderEvents: []
            };     
            //! I don't believe this is needed, but I thought I would create the ref in case there is a 
            //! Future need. 
            this.calendar = React.createRef();                       
        }
        
        public componentDidMount() {   
            //*Once the component completes it's intital mount; let's get some data.      
            this.getandProcessEventData();        
        }

        public render(): React.ReactElement<IGraphApiProps> {
            const handleChange  = (event)  => {
                //* The event.preventDefault() command below keeps the calendar from refreshing after
                //* the onChange event is fired.  Also, I have added a work-a-round for the event method 
                //* being depecated, by adding {(event) => handleChange(event)}, and capturing the event in
                //* this method.   
                event.preventDefault();
            };
            const handleClickEvent = () => {
                //todo: Finish.  
                alert("CLICK!"); 
            };
            //* Let's check to see if there are any events in the state to display.  If not, we
            //* we will show an empty calendar. 
            if (this.state.calenderEvents.length !== 0) {         
                return (
                    <section className={styles.graphApi}>
                        <h2>{escape(this.props.GroupCalendarName)}</h2>
                        <Calendar  
                            ref={this.calendar} 
                            calendarType="US"
                            defaultView="month" 
                            events={this.state.calenderEvents}  
                            onChange={(event) => handleChange(event)}
                            onClickEvent={() => handleClickEvent()} />
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
                    //* Here we are duplicating the current calender entries stored in the state value, 
                    //* so we don't loose them when added new entries.     
                    let calenderEventsState: any[] = this.state.calenderEvents;                   
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
                            }); 
                        }                
                        //* Finally, we are going to set the state with the events have in this iteration of the Calendar Collection 
                        this.setState({calenderEvents:calenderEventsState});  
                    });                  
                }); 
            }     
        }

    //#endregion

}