export interface IListChartsWebPartProps {
  description: string;
  numCharts: number;
  maxResults: number;
  chartConfig: Array<IChartConfiguration>;
  siteOption: string;
  otherUrl: string;
}

export interface IChartConfiguration {
  title: string;
  description: string;
  type: string;
  size: number;
  listId?: string;
  theme: string;
  bgColors: Array<string>;
  hoverColors: Array<string>;
  labelListFieldId?: string;
  dataListFieldId?: string;
  uniqueListFieldId?: string;
  action?: string;
}
