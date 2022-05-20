import * as React from 'react';
import styles from './ChartExample.module.scss';
import { IChartExampleProps } from './IChartExampleProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
} from 'chart.js';
import { Bar } from 'react-chartjs-2';


ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend
);

export const options = {
  responsive: true,
  plugins: {
    legend: {
      position: 'top' as const,
    },
    title: {
      display: true,
      text: 'Bar Chart',
    },
  },
};



export default class ChartExample extends React.Component<IChartExampleProps, {}> {
  public render(): React.ReactElement<IChartExampleProps> {
    const {
      chartData,      
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <Bar options={options} data={this.props.chartData} />
    );
  }
}
