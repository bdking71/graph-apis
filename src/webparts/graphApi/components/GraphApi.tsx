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
    calenderEvents: IACCalendarEvent[]; 
}

export default class GraphApi extends React.Component<IGraphApiProps, iState> {

    public constructor (props:IGraphApiProps) {
        super(props); 
        this.state = { 
            myURL: location.protocol + "//" + location.host + location.pathname,
            calenderEvents: []
        };                            
    }
    
    public componentWillMount() {
        this.getandProcessEventData();        
    }

    public render(): React.ReactElement<IGraphApiProps> {
        const handleChange  = ()  => {
            event.preventDefault();
        };
        if (this.state.calenderEvents) {
            return (
                <section className={styles.graphApi}>
                    <h2>{escape(this.props.GroupCalendarName)}</h2>
                    <Calendar   
                        calendarType="US"
                        defaultView="month" 
                        events={this.state.calenderEvents}  
                        onChange={() => handleChange()} />
                </section>
                );
        } else {
            return (<React.Fragment></React.Fragment>);
        }
    }


    /* 
        Using the Graph API,  We are going to pull the all events 
        from the calendars that the current users has access.  We 
        will need to filter out the calendar we want to display in 
        our webpart. 
    */
    private getandProcessEventData = ():void => {                        
        this.props.context.msGraphClientFactory.getClient()
        .then((client: MSGraphClient): void => {          
            client.api(`/groups/${this.props.GroupCalendarGUID}/events`)
            .select('subject,body,bodyPreview,organizer,attendees,start,end,location')
            .get((error, messages: any, rawResponse?: any) => {   
                console.log(messages); 
                if (!messages) {
                    console.error(error);   
                    throw error; 
                } else {   
                    /* 
                        Using the data we received, we are going to transform the data to match what is expected
                        by the React-Awesome-Calendar and store the data in the Application State.  
                    */
                    let retval = [];         
                    messages.value.map((eventItem) => {  

                        console.log(eventItem);

                        let tmp = {
                            id:  eventItem.id,
                            title: eventItem.subject,
                            to:  new Date(eventItem.end.dateTime).toLocaleDateString(),
                            from: new Date(eventItem.start.dateTime).toLocaleDateString(),
                            color: "#0F04F7"
                         };
                        retval.push(tmp); 
                    });
                    /*
                        We probally don't need to save the data we received from MSGraph, but just in case 
                    */
                    this.setState({calenderEvents:retval});  
                }  
            });
        });
    }

}