export interface IChartExampleProps {  
  chartData: IChartData;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

export interface IChartData {
  labels: string [];
  datasets: any [];
}
