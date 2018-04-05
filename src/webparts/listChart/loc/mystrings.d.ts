declare interface IGeneralOptionsStrings {
  Name: string;
  Description: string;
  NumCharts: string;
  MaxResults: string;
}
declare interface IChartConfigurationStrings {
  Name: string;
  Title: string;
  Description: string;
  Type: string;
  Size: string;
  Theme: string;
  List: string;
  Col1: string;
  Col2: string;
  Unique: string;
  Act: string;
}
declare interface IListChartWebPartStrings {
  PropertyPaneDescription: string;
  GeneralOptions: IGeneralOptionsStrings;
  ChartConfigurationOptions: IChartConfigurationStrings;
}
declare module 'ListChartWebPartStrings' {
  const strings: IListChartWebPartStrings;
  export = strings;
}
