import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'GraphApiWebPartStrings';
import GraphApi from './components/GraphApi';
import { IGraphApiProps } from './components/IGraphApiProps';

export interface IGraphApiWebPartProps {
  GroupCalendarName: string;
  GroupCalendarGUID: string; 
}

export default class GraphApiWebPart extends BaseClientSideWebPart<IGraphApiWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGraphApiProps> = React.createElement(
      GraphApi,
      {
        GroupCalendarName: this.properties.GroupCalendarName,
        GroupCalendarGUID: this.properties.GroupCalendarGUID,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);

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
                PropertyPaneTextField('GroupCalendarGUID', {
                    label: strings.GroupCalendarGUIDFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
