import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ChartExampleWebPartStrings';
import ChartExample from './components/ChartExample';
import { IChartExampleProps, IChartData } from './components/IChartExampleProps';
import { faker } from '@faker-js/faker';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IWorkStatusItem } from '../../models/IWorkStatusItem';

export interface IChartExampleWebPartProps {
  description: string;
}


export default class ChartExampleWebPart extends BaseClientSideWebPart<IChartExampleWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    
    this._getListItems()
      .then(response => {
      
        const data: IChartData = {
          labels: [],
          datasets: []
        };
        
        data.datasets.push({			
          label: "Estado de trabajos",
          data: response.map(item => item.PercentComplete),
          backgroundColor: response.map(item => `#${Math.floor(Math.random()*16777215).toString(16)}`),
        }); //'rgba(255, 99, 132, 0.5)'
		    data.labels = response.map(item => item.Title);


        const element: React.ReactElement<IChartExampleProps> = React.createElement(
          ChartExample,
          {
            chartData: data,
            isDarkTheme: this._isDarkTheme,
            environmentMessage: this._environmentMessage,
            hasTeamsContext: !!this.context.sdks.microsoftTeams,
            userDisplayName: this.context.pageContext.user.displayName
          }
        );

        ReactDom.render(element, this.domElement);
        
      });

  }

  private _getListItems(): Promise<IWorkStatusItem[]> { 
    const endpoint: string = this.context.pageContext.web.absoluteUrl
    + `/_api/web/lists/getbytitle('Work Status')/items?$select=Id,Title, PercentComplete`;
    
    return this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {                
        return jsonResponse.value;
      }) as Promise<IWorkStatusItem[]>;
  }


  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
