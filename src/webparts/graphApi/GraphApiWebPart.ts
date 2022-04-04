
//#region [header]
    //[header] @File Name:          GraphApiWebPart.ts
    //[header] @Description:        Retrieves calendar data from msGraph and displays the data using 
    //[header]                      react-awesome-calendar [https://www.npmjs.com/package/react-awesome-calendar] and
    //[header]                      ReactWindow [https://www.npmjs.com/package/reactjs-windows]  
    //[header] @Author:             Bryan King
    //[header] @Date:               March 29, 2022
    //[header] @File Version:       20220328-1243  
//#endregion


//#region [Imports]

  import * as React from 'react';
  import * as ReactDom from 'react-dom';
  import { Version } from '@microsoft/sp-core-library';
  import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
  import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
  import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
  import { IReadonlyTheme } from '@microsoft/sp-component-base';
  import * as strings from 'GraphApiWebPartStrings';
  import GraphApi from './components/GraphApi';
  import { IGraphApiProps } from './components/IGraphApiProps';
  import PnPTelemetry from "@pnp/telemetry-js";

//#endregion

//#region [Interfaces]

  export interface IGraphApiWebPartProps {
    GroupCalendarName: string;
    CalendarCollection: any[];
    SharePointCalendarCollection: any[];
  }

//#endregion

const telemetry = PnPTelemetry.getInstance();
export default class GraphApiWebPart extends BaseClientSideWebPart<IGraphApiWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGraphApiProps> = React.createElement(
      GraphApi,
      {
        GroupCalendarName: this.properties.GroupCalendarName,  
        CalendarCollection: this.properties.CalendarCollection,
        SharePointCalendarCollection: this.properties.SharePointCalendarCollection,        
        Context: this.context
      }
    );
    ReactDom.render(element, this.domElement);
    telemetry.optOut();    
    
    console.log("🚀 ~ file: GraphApiWebPart.ts ~ line 55 ~ GraphApiWebPart ~ render ~ SharePointCalendarCollection", this.properties.SharePointCalendarCollection);
  }

  //#region [ProtectedMethods]

    protected onDispose(): void {
      ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
      return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
      return {
        pages: [
          {
            header: {
              description: strings.PropertyPaneDescription
            },
            groups: [
              {
                groupName: strings.BasicGroupName,
                groupFields: [ 
                  PropertyPaneTextField('GroupCalendarName', {
                      label: strings.GroupCalendarNameFieldLabel
                  }),
                  //* The CalendarCollection property is designed to keep 
                  //* information on Outlook Group Calendars.  
                  PropertyFieldCollectionData("CalendarCollection", {
                      key: "collectionData",
                      label: "Exchange Calendars",
                      panelHeader: "Calendars to Display",
                      manageBtnLabel: "Configure Calendars",
                      value: this.properties.CalendarCollection,
                      fields: [
                        {
                          id: "CalendarTitle",
                          title: "Calendar Title",
                          type: CustomCollectionFieldType.string,
                          required: true
                        },
                        {
                          id: "CalendarGuid",
                          title: "Calendar GUID",
                          type: CustomCollectionFieldType.string,
                          required: true
                        },
                        {
                          id: "CalendarColor",
                          title: "Calendar Color",
                          type: CustomCollectionFieldType.color,
                          required: true
                        },
                        
                      ],
                      disabled: false
                    }),
                    //* The SharePointCalendarCollection property is designed to keep 
                    //* information on Calendars that exists in SharePoint.  
                    PropertyFieldCollectionData("SharePointCalendarCollection", {
                      key: "SharePointCalendarCollection",
                      label: "SharePoint Calendars",
                      panelHeader: "Configure SharePoint Calendars to Display",
                      manageBtnLabel: "Configure SharePoint Calendars",
                      value: this.properties.SharePointCalendarCollection,
                      fields: [
                        {
                          id: "SharePointCalendarTitle",
                          title: "Calendar Title",
                          type: CustomCollectionFieldType.string,
                          required: true
                        },
                        {
                          id: "SharePointCalendarSiteUrl",
                          title: "SharePoint Site",
                          type: CustomCollectionFieldType.string,
                          required: true
                        },
                        {
                          id: "SharePointCalendarName",
                          title: "Calendar Name",
                          type: CustomCollectionFieldType.string,
                          required: true
                        },
                        {
                          id: "SharePointCalendarColor",
                          title: "Calendar Color",
                          type: CustomCollectionFieldType.color,
                          required: true
                        },
                        
                      ],
                      disabled: false
                    })

                ]
              }
            ]
          }
        ]
      };
    }

  //#endregion

}