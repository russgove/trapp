import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'trTimeCardStrings';
import TrTimeCard from './components/TrTimeCard';
import { ITrTimeCardProps } from './components/ITrTimeCardProps';
import { ITrTimeCardWebPartProps } from './ITrTimeCardWebPartProps';

export default class TrTimeCardWebPart extends BaseClientSideWebPart<ITrTimeCardWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITrTimeCardProps > = React.createElement(
      TrTimeCard,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
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
