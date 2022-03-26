/* https://pnp.github.io/sp-dev-fx-property-controls*/

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

export interface IGraphApiWebPartProps {
  GroupCalendarName: string;
  CalendarCollection: any[];
}

const telemetry = PnPTelemetry.getInstance();

export default class GraphApiWebPart extends BaseClientSideWebPart<IGraphApiWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGraphApiProps> = React.createElement(
      GraphApi,
      {
        GroupCalendarName: this.properties.GroupCalendarName,  
        CalendarCollection: this.properties.CalendarCollection,
        Context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
    telemetry.optOut();
  }


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
                PropertyFieldCollectionData("CalendarCollection", {
                    key: "collectionData",
                    label: "Exchange Calendars",
                    panelHeader: "Calendars to Display in this Webpart.",
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
                  })

              ]
            }
          ]
        }
      ]
    };
  }
}
