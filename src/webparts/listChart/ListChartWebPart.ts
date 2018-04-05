import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  IPropertyPaneGroup,
  PropertyPaneChoiceGroup,
  IPropertyPaneChoiceGroupOption,
  PropertyPaneDropdown,
  PropertyPaneButton,
  PropertyPaneButtonType,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import {
  PropertyFieldListPicker
} from '@pnp/spfx-property-controls';

import * as strings from 'ListChartWebPartStrings';
import ListChart from './components/ListChart';
import { IListChartProps } from './components/IListChartProps';

import { IDataHelper } from '../data-helpers/DataHelperBase';
import { DataHelpersFactory } from '../data-helpers/DataHelpersFactory';
import { ISPField } from '../common/SPEntities';

export interface IChartConfigOption {
  title: string;
  description: string;
  type: string;
  size: number;
  theme: string;
  list: string;
  columns: Array<IPropertyPaneDropdownOption>;
  colsDisabled: boolean;
  col1: string;
  col2: string;
  unique: string;
  act: string;
}

export interface IListChartWebPartProps {
  description: string;
  numCharts: number;
  maxResults: number;
  firstLoad: boolean;
  chartConfig: Array<IChartConfigOption>;
}

export default class ListChartWebPart extends BaseClientSideWebPart<IListChartWebPartProps> {

  private colsDisabled: boolean = true;

  private _chartTypeOptions: IPropertyPaneChoiceGroupOption[] = [
    { key: 'bar', text: 'Bar', iconProps: { officeFabricIconFontName: 'BarChartVertical' }},
    { key: 'hbar', text: 'Horizontal Bar', iconProps: { officeFabricIconFontName: 'BarChartHorizontal' }},
    { key: 'doughnut', text: 'Doughnut', iconProps: { officeFabricIconFontName: 'DonutChart' }},
    { key: 'line', text: 'Line', iconProps: { officeFabricIconFontName: 'LineChart' }},
    { key: 'pie', text: 'Pie', iconProps: { officeFabricIconFontName: 'PieSingle' }}
  ];

  private _chartSizeOptions: IPropertyPaneChoiceGroupOption[] = [
    { key: 3, text:'Small' },
    { key: 6, text:'Medium' },
    { key: 9, text: 'Medium-Large' },
    { key: 12, text: 'Large' }
  ];

  private _chartColActions: IPropertyPaneDropdownOption[] = [
    { key: 'average', text: 'Average' },
    { key: 'count', text: 'Count' },
    { key: 'sum', text: 'Sum' }
  ];

  private _defaultChartConfig(chartDesc: string): IChartConfigOption {
    var defConf: IChartConfigOption = {
      title: 'Chart Title',
      description: chartDesc,
      type: 'doughnut',
      size: 12,
      theme: 'Random',
      list: '',
      columns: [],
      colsDisabled: true,
      col1: '',
      col2: '',
      unique: '',
      act: ''
    };
    return defConf;
  }

  private _updateListColumns(newValue: string, chartConfig: IChartConfigOption): void {
    const respFields: IPropertyPaneDropdownOption[] = [];
    // Clear out old values
    chartConfig.columns = [];
    chartConfig.col1 = '';
    chartConfig.col2 = '';
    chartConfig.unique = '';
    chartConfig.act = '';

    // If newValue is empty, return.
    if (newValue === '') {
      chartConfig.colsDisabled = true;
      this.context.propertyPane.refresh();
      return;
    }

    // Fetch the list fields and fill the columns options
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'Loading List Data...');
    const dataHelper: IDataHelper = DataHelpersFactory.createDataHelper(this.context);
    dataHelper.getFields(newValue)
      .then((response: ISPField[]) => {        
        response.forEach(field => {
          respFields.push({key: field.Id, text: field.Title});
        });
        chartConfig.columns = respFields;
        chartConfig.colsDisabled= respFields.length === 0;
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.context.propertyPane.refresh();
        this.render();
      })
      .catch((err) => {
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.context.statusRenderer.renderError(this.domElement, `There was an error loading the list fields.`);
        console.error(err);
      });
  }

  public constructor() {
    super();
    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
  }

  public render(): void {
    if (this.properties.firstLoad) {
      this.properties.firstLoad = false;
      this.properties.chartConfig = [
        this._defaultChartConfig('')
      ];
    }

    const element: React.ReactElement<IListChartProps > = React.createElement(
      ListChart,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    var pPath = propertyPath;
    var pPathInd = propertyPath[12];
    if (pPath === 'numCharts' && oldValue != newValue) {
      if (this.properties.chartConfig.length < newValue) {
        while (this.properties.chartConfig.length < newValue) {
          this.properties.chartConfig.push(this._defaultChartConfig('Chart Description'));
        }
      } else if (this.properties.chartConfig.length > newValue) {
        while (newValue < this.properties.chartConfig.length) {
          this.properties.chartConfig.pop();
        }
      }
    }

    if (propertyPath.indexOf('[') != -1) {
      pPath = propertyPath.substring(16).replace('\"]','');
    }
    if (pPath === 'list' && (oldValue != newValue)) {
      this.properties.chartConfig[pPathInd].list = newValue;
      this._updateListColumns(newValue,this.properties.chartConfig[pPathInd]);
    }
    this.context.propertyPane.refresh();
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let generalOptionsGroup: IPropertyPaneGroup = {
      groupName: 'General Options',
      groupFields: [
        PropertyPaneTextField('description', {
          label: 'Web Part Title'
        }),
        PropertyPaneSlider('numCharts', {
          label: 'Number of Charts',
          value: this.properties.numCharts,
          min: 1,
          max: 10,
          step: 1,
          showValue: true
        }),
        PropertyPaneSlider('maxResults', {
          label: 'Max # of list items',
          min: 1,
          max: 5000,
          step: 100,
          showValue: true,
          value: this.properties.maxResults
        })
      ]
    };
    let chartsConfigurationGroups: IPropertyPaneGroup[] = [
      generalOptionsGroup
    ];
    for (let i = 0; i < this.properties.numCharts; i++) {
      const chartConfig: string = `chartConfig[${i}]`;
      chartsConfigurationGroups.push({
        groupName: `Chart ${i + 1} Configuration`,
        groupFields: [
          PropertyPaneTextField(`${chartConfig}["title"]`, {
            label: 'Chart Title'
          }),
          PropertyPaneTextField(`${chartConfig}["description"]`, {
            label: 'Description'
          }),
          PropertyPaneChoiceGroup(`${chartConfig}["type"]`, {
            label: 'Type',
            options: this._chartTypeOptions
          }),
          PropertyPaneDropdown(`${chartConfig}["size"]`, {
            label: 'Size',
            options: this._chartSizeOptions
          }),
          PropertyPaneButton(`${chartConfig}["theme"]`, {
            buttonType: PropertyPaneButtonType.Normal,
            text: 'Generate Theme',
            icon: 'Color',
            onClick: ((val) => {
              return new Date().valueOf();
            })
          }),
          PropertyFieldListPicker(`${chartConfig}["list"]`, {
            label: 'List',
            selectedList: this.properties.chartConfig[i].list,
            context: this.context,
            onPropertyChange: this.onPropertyPaneFieldChanged,
            properties: this.properties,
            key: 'listId'
          }),
          PropertyPaneDropdown(`${chartConfig}["col1"]`, {
            label: 'Label Column',
            selectedKey: this.properties.chartConfig[i].col1,
            options: this.properties.chartConfig[i].columns,
            disabled: this.properties.chartConfig[i].colsDisabled
          }),
          PropertyPaneDropdown(`${chartConfig}["col2"]`, {
            label: 'Data Column',
            selectedKey: this.properties.chartConfig[i].col2,
            options: this.properties.chartConfig[i].columns,
            disabled: this.properties.chartConfig[i].colsDisabled
          }),
          PropertyPaneDropdown(`${chartConfig}["unique"]`, {
            label: 'Unique Identifier',
            selectedKey: this.properties.chartConfig[i].unique,
            options: this.properties.chartConfig[i].columns,
            disabled: this.properties.chartConfig[i].colsDisabled
          }),
          PropertyPaneDropdown(`${chartConfig}["act"]`, {
            label: 'Operation',
            selectedKey: this.properties.chartConfig[i].act,
            options: this._chartColActions,
            disabled: this.properties.chartConfig[i].colsDisabled
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
          groups: chartsConfigurationGroups
        }
      ]
    };
  }
}
