import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ListChartWebPartStrings';
import ListChart from './components/ListCharts';
import { IListChartProps } from './components/IListChartsProps';

export interface IListChartWebPartProps {
  description: string;
}

export default class ListChartWebPart extends BaseClientSideWebPart<IListChartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IListChartProps > = React.createElement(
      ListChart, null
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('0.0.1.0');
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
