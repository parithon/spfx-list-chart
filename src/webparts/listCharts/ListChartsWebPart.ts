import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, ServiceScope } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ListChartsWebPartStrings';
import ListCharts from './components/ListCharts';
import { IListChartsProps } from './components/IListChartsProps';
import { IListChartsWebPartProps } from './IListChartsWebPart';
import ChartOptions from './ChartOptions';
import ServiceFactory from '../../services/ServiceFactory';
import { ISPService } from '../../services/ISPService';

export default class ListChartsWebPart extends BaseClientSideWebPart<IListChartsWebPartProps> {

  private _service: ISPService;

  public render(): void {
    const element: React.ReactElement<IListChartsProps> = React.createElement(
      ListCharts,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    console.debug('onPropertyPaneConfigurationStart called');
    if (!this.properties.chartConfig) {
      this.properties.chartConfig = [ChartOptions.DefaultChartConfiguration('Demo Chart, Edit Web Part to Customize')];
    }
    this.getPropertyPaneData();
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

  private getPropertyPaneData() {
    if (this._service === undefined) {
      this._service = ServiceFactory.createService(this.context);
    }
    this._service.getRootWeb()
      .then(rootWebUrl => {
        console.debug('rootWebUrl', rootWebUrl);
        this._service.getSites(rootWebUrl)
          .then(siteCollection => {
            console.debug('siteCollection', siteCollection);
          })
          .catch(err => console.error('An error occured while retrieving the SharePoint sites.', err));
      })
      .catch(err => console.error("An error occured while retrieving the root web.", err));
  }
}
