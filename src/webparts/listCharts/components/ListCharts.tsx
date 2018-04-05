import * as React from 'react';
import styles from './ListCharts.module.scss';
import { IListChartProps } from './IListChartsProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class ListChart extends React.Component<IListChartProps, {}> {
  public render(): React.ReactElement<IListChartProps> {
    return (
      <div className={ styles.listChart }>
      </div>
    );
  }
}
