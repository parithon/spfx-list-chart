import { IPropertyPaneDropdownOption, IPropertyPaneChoiceGroupOption } from "@microsoft/sp-webpart-base/lib";
import { IChartConfiguration } from "./IListChartsWebPart";
import * as cs from 'color-scheme';

export interface IChartColors {
  bgColors: Array<string>;
  hoverColors: Array<string>;
}

export default class ChartOptions {
  public static ChartSizeOptions: IPropertyPaneDropdownOption[] = [
    { key: 3, text: 'Small' },
    { key: 6, text: 'Medium' },
    { key: 9, text: 'Medium-Large' },
    { key: 12, text: 'Large' }
  ];
  public static ChartTypeOptions: IPropertyPaneChoiceGroupOption[] = [
    { key: 'bar', text: 'Bar', iconProps: { officeFabricIconFontName: 'BarChartVertical' } },
    { key: 'hBar', text: 'Horizontal Bar', iconProps: { officeFabricIconFontName: 'BarChartHorizontal' }},
    { key: 'donut', text: 'Doughnut', iconProps: { officeFabricIconFontName: 'DonutChart' }},
    { key: 'line', text: 'Line', iconProps: { officeFabricIconFontName: 'LineChart' }},
    { key: 'pie', text: 'Pie', iconProps: { officeFabricIconFontName: 'PieSingle' }}
  ];
  public static ChartActionOptions: IPropertyPaneDropdownOption[] = [
    { key: 'avg', text: 'Average' },
    { key: 'count', text: 'Count' },
    { key: 'sum', text: 'Sum' }
  ];
  public static DefaultOptions: object = {
    legend: {
      display: false,
      layout: {
        padding: 10
      },
      position: 'bottom',
      labels: {
        fontColor: 'rgba(100,100,100,1.0)'
      }
    }
  };
  public static DefaultChartConfiguration = (chartDesc: string): IChartConfiguration => {
    var colors: IChartColors = ChartOptions.RandomColors();
    var defConfig: IChartConfiguration = {
      title: 'Chart Title',
      description: chartDesc,
      type: 'donut',
      size: 12,
      theme: 'Random',
      bgColors: colors.bgColors,
      hoverColors: colors.hoverColors
    };
    return defConfig;
  }
  private static RandomColors(): IChartColors {
    var colors = {bgColors: [], hoverColors: []};
    var colorTheme = new cs;
    var colorHue = Math.floor(Math.random()*360);
    var colorPalette = colorTheme.from_hue(colorHue).scheme('analogic').variation('default');
    colors.bgColors = ChartOptions.ShuffleArray(colorPalette.add_complement(true).colors());
    colors.hoverColors = ChartOptions.ShuffleArray(colorPalette.add_complement(true).colors()).splice(6,6);
    colors.bgColors.forEach((hex,idx) => { colors.bgColors[idx] = '#' + hex; });
    colors.hoverColors.forEach((hex,idx) => { colors.hoverColors[idx] = '#' + hex; });
    return colors;
  }
  private static ShuffleArray(array) {
    var currentIndex = array.length, temporaryValue, randomIndex;
    while (currentIndex !== 0) {
      randomIndex = Math.floor(Math.random() * currentIndex);
      currentIndex--;
      temporaryValue = array[currentIndex];
      array[currentIndex] = array[randomIndex];
      array[randomIndex] = temporaryValue;
    }
    return array;
  }
}
