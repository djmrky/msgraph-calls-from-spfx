import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'MsGraphCallsWebPartStrings';
import MsGraphCalls from './components/MsGraphCalls';
import { IMsGraphCallsProps } from './components/IMsGraphCallsProps';


import { MSGraphClient } from '@microsoft/sp-http';




export interface IMsGraphCallsWebPartProps {
  description: string;
}

export default class MsGraphCallsWebPart extends BaseClientSideWebPart<IMsGraphCallsWebPartProps> {

  public render(): void {

    const element: React.ReactElement<IMsGraphCallsProps> = React.createElement(
      MsGraphCalls,
      {
        //description: this.properties.description,
        spfxContext: this.context
        //getUsersFromAzureAD: this.getUsersFromAzureAD
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
