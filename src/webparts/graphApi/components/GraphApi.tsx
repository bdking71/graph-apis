import * as React from 'react';
import styles from './GraphApi.module.scss';
import {IGraphApiProps} from './IGraphApiProps';
import {escape} from '@microsoft/sp-lodash-subset';
import {MSGraphClient} from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import Calendar from 'react-awesome-calendar';
import {IACCalendarEvents, IACCalendarEvent} from './IACCalenderTypes';

export interface iState {   
    myURL: string;       
    calenderEvents: any[]; 
}

export default class GraphApi extends React.Component<IGraphApiProps, iState> {
    public constructor (props:IGraphApiProps) {
        super(props);         
        this.state = { 
            myURL: location.protocol + "//" + location.host + location.pathname,
            calenderEvents: []
        };                            
    }
    
    public componentDidMount() {        
        this.getandProcessEventData();        
    }

    public render(): React.ReactElement<IGraphApiProps> {
        const handleChange  = ()  => {
            event.preventDefault();
        };
        const handleClickEvent = () => {
            alert("CLICK!"); 
        };
 
        if (this.state.calenderEvents.length !== 0) {         
            return (
                <section className={styles.graphApi}>
                    <h2>{escape(this.props.GroupCalendarName)}</h2>
                    <Calendar   
                        calendarType="US"
                        defaultView="month" 
                        events={this.state.calenderEvents}  
                        onChange={() => handleChange()}
                        onClickEvent={() => handleClickEvent()} />
                </section> );
        } else {                  
            return (<Calendar   /> );   
        }
    }

   
                    
    /* 
        Using the Graph API,  We are going to pull the all events 
        from the calendars that the current users has access.  We 
        will need to filter out the calendar we want to display in 
        our webpart. 
    */
    private getandProcessEventData = ():void => {               
        for (let cnt = 0; cnt <= (this.props.CalendarCollection.length - 1); cnt ++) {            
            this.props.Context.msGraphClientFactory.getClient()
            .then((client: MSGraphClient): void => {                   
                let calenderEventsState: any[] = this.state.calenderEvents;                   
                client.api(`/groups/${this.props.CalendarCollection[cnt].CalendarGuid}/events`)            
                .select('subject,body,bodyPreview,organizer,attendees,start,end,location')
                .get((error, messages: any, rawResponse?: any) => {   
                    if (!messages) {
                        console.error(error);   
                        throw error; 
                    } else {   
                        /* 
                            Using the data we received, we are going to transform the data to match what is expected
                            by the React-Awesome-Calendar and store the data in the Application State.  
                        */    
                        messages.value.map((eventItem) => {  
                            let tmp = {
                                id:  eventItem.id,
                                title: eventItem.subject,
                                to:  new Date(eventItem.end.dateTime).toLocaleDateString(),
                                from: new Date(eventItem.start.dateTime).toLocaleDateString(),
                                color: `${this.props.CalendarCollection[cnt].CalendarColor}`
                            };                                              
                            calenderEventsState.push(tmp);                       
                        }); 
                    }                
                    this.setState({calenderEvents:calenderEventsState});  
                });                  
            }); 
        }        
    }

}