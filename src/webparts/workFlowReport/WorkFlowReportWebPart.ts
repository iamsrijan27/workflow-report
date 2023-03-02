import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WorkFlowReportWebPartStrings';
import WorkFlowReport from './components/WorkFlowReport';
import { IWorkFlowReportProps } from './components/IWorkFlowReportProps';
import IService from '../../common/Services/IServices';
import Service from '../../common/Services/Services';
import { sp } from '@pnp/sp';
export interface IWorkFlowReportWebPartProps {
  webPartTitle: string;
  listName: string;
  BouncelistName: string;
  listURL: string;
  pageSize: string;
  webPartPageURL: string;

}

export default class WorkFlowReportWebPart extends BaseClientSideWebPart<IWorkFlowReportWebPartProps> {
  private helperService: IService;
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IWorkFlowReportProps> = React.createElement(
      WorkFlowReport,
      {
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        context: this.context,
        webPartTitle: this.properties.webPartTitle,
        listName: this.properties.listName,
        listURL:this.properties.listURL,
        BouncelistName: this.properties.BouncelistName,
        helperService: this.helperService,
        pageSize: this.properties.pageSize != "" ? Number(this.properties.pageSize) : 10,
        webPartPageURL: this.properties.webPartPageURL,
        
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.helperService = new Service();

    return super.onInit().then(_ => {
      // other init code may be present
      sp.setup({
        spfxContext: {pageContext: this.context.pageContext}      
      });
    });
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

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
                PropertyPaneTextField('webPartTitle', {
                  label: "Title",
                  description: "Web part title"
                }),
                PropertyPaneTextField('listName', {
                  label: "List Name",
                  description: "List Name"
                }),
                PropertyPaneTextField('BouncelistName', {
                  label: "Bounce Back List Name",
                  description: "Bounce Back List Name"
                }),
                PropertyPaneTextField('listURL', {
                  label: "Category Contacts List URL",
                  description: "Categroy Contacts List URL"
                }),
                PropertyPaneTextField('pageSize', {
                  label: "Page Size",
                  description: "Page Size"
                }),
                
                PropertyPaneTextField('webPartPageURL', {
                  label: "Web Part Page URL",
                  description: "Please enter web part page URL"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
