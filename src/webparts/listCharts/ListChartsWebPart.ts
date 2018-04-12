import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, ServiceScope, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdownOptionType,
  PropertyPaneDropdown,
  IPropertyPaneGroup,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneButton,
  PropertyPaneButtonType
} from '@microsoft/sp-webpart-base';

import * as strings from 'ListChartsWebPartStrings';
import ListCharts from './components/ListCharts';
import { IListChartsProps } from './components/IListChartsProps';
import { IListChartsWebPartProps, IChartConfiguration } from './IListChartsWebPart';
import ChartOptions from './ChartOptions';
import ServiceFactory from '../../services/ServiceFactory';
import { ISPService, ListQueryOrderBy } from '../../services/ISPService';

export const AVAILALBE_SITES_KEY: string = 'available';
export const OTHER_SITE_KEY: string = 'other';
export const NO_OPTIONS_AVAILABLE: string = 'no-options';

export default class ListChartsWebPart extends BaseClientSideWebPart<IListChartsWebPartProps> {

  private _service: ISPService;
  private _sites: IPropertyPaneDropdownOption[] = [];
  private _listOptions: Array<IPropertyPaneDropdownOption[]> = [];

  public render(): void {
    const element: React.ReactElement<IListChartsProps> = React.createElement(
      ListCharts,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);

    if (DEBUG && !this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.open();
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    console.debug('onPropertyPaneConfigurationStart called');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    console.debug('onPropertyPaneFieldChanged called', propertyPath);
    let pPath = propertyPath;
    const pPathIdx = propertyPath[13];
    const chartConfig: IChartConfiguration = this.properties.chartConfig[pPathIdx];

    if (pPath === 'numCharts' && oldValue !== newValue) {
      if (this.properties.chartConfig.length < newValue) {
        while (this.properties.chartConfig.length < newValue) {
          const idx = this.properties.chartConfig.push(ChartOptions.DefaultChartConfiguration(this.context.pageContext.web.absoluteUrl)) - 1;
          this.getAvailableListsForSite(this.context.pageContext.web.absoluteUrl, idx);
        }
      } else if (this.properties.chartConfig.length > newValue) {
        while (this.properties.chartConfig.length > newValue) {
          this.properties.chartConfig.pop();
        }
      }
    }

    if (propertyPath.indexOf('[') !== -1) {
      pPath = propertyPath.substring(18).replace('\']', '');
    }

    this.context.propertyPane.refresh();
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    console.debug('getPropertyPaneConfiguration called');
    if (!this.properties.chartConfig) {
      this.properties.chartConfig = [ChartOptions.DefaultChartConfiguration(this.context.pageContext.web.absoluteUrl)];
    }
    const columnPropertyOptions: IPropertyPaneGroup[] = [];
    columnPropertyOptions.push({
      groupName: strings.BasicGroupName,
      groupFields: [
        PropertyPaneTextField('description', {
          label: strings.DescriptionFieldLabel
        }),
        PropertyPaneSlider('numCharts', {
          label: 'Number of Charts',
          min: 1,
          max: 10
        }),
        PropertyPaneSlider('maxResults', {
          label: 'Maximum Data Results',
          min: 1,
          max: 5000
        })
      ]
    });
    for (var i = 0; i < this.properties.numCharts; i++) {
      const chartConfig: IChartConfiguration = this.properties.chartConfig[i];
      const cc: string = `chartConfig['${i}']`;
      columnPropertyOptions.push({
        groupName: strings.ChartGroupName.replace(/\{0\}/g, `${(i + 1)}`),
        groupFields: [
          PropertyPaneTextField(`${cc}['title']`, {
            label: 'Chart Title'
          }),
          PropertyPaneTextField(`${cc}["description"]`, {
            label: 'Description'
          }),
          PropertyPaneChoiceGroup(`${cc}["type"]`, {
            label: 'Chart Type',
            options: ChartOptions.ChartTypeOptions
          }),
          PropertyPaneDropdown(`${cc}['size'`, {
            label: 'Chart Size',
            options: ChartOptions.ChartSizeOptions
          }),
          PropertyPaneButton(`${cc}["theme"]`, {
            buttonType: PropertyPaneButtonType.Normal,
            text: 'Generate Theme',
            icon: 'Color',
            onClick: (val => {
              return new Date().valueOf();
            })
          }),
          PropertyPaneDropdown(`${cc}['siteUrl']`, {
            label: 'SharePoint Site',
            options: this._sites,
            selectedKey: this.context.pageContext.web.absoluteUrl
          }),
          PropertyPaneTextField(`${cc}['otherUrl']`, {
            label: 'Other Site Url',
            placeholder: 'Url (e.g. https://sp.gfins.org/sites/mysite)',
            disabled: this.properties.chartConfig[i].siteUrl !== OTHER_SITE_KEY
          }),
          PropertyPaneDropdown(`${cc}['listUrl']`, {
            label: 'List Data Source',
            options: this._listOptions[i],
            disabled: (this.properties.chartConfig[i].dataDisabled)
          })
        ]
      });
    }
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: columnPropertyOptions
        }
      ]
    };
  }

  private getPropertyPaneData(): void {
    throw new Error('Not Implemented Yet');
  }

  private getAvailableListsForSite(siteUrl: string, idx): void {
    throw new Error('Not Implemented Yet');
  }

  private isUrl = (value: string): boolean => {
    const regx = new RegExp("(https\:\/+)([^\/\s]*)([a-z0-9\-@\^=%&;\/~\+]*)[\?]?([^ \#]*)#?([^ \#]*)");
    const result = regx.test(value);
    return result;
  }
}
