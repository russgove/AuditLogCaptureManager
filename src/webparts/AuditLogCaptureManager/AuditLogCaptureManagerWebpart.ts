import { Version } from '@microsoft/sp-core-library';
import { Log } from '@microsoft/sp-core-library';
import { AadHttpClient } from '@microsoft/sp-http';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'AuditLogCaptureManagerWebPartStrings';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import AuditLogCaptureManager from './components/AuditLogCaptureManager';
import { IAuditLogCaptureManagerProps } from './components/IAuditLogCaptureManagerProps';
import { IAuditLogCaptureManagerState } from './components/IAuditLogCaptureManagerState';

const LOG_SOURCE: string = 'AuditLogCaptureMananger';
export interface IAuditLogCaptureManagerWebPartProps {
  managementApiUrl: string;
}

export default class AuditLogCaptureManagerWebPart extends BaseClientSideWebPart<IAuditLogCaptureManagerWebPartProps> {

  private aadHttpClient: AadHttpClient;
  public onInit(): Promise<void> {
    debugger;
    Log.info(LOG_SOURCE, 'Initialized TrondocsCommandsCommandSet');
    //sessionStorage.setItem("spfx-debug", ""); ////   REMOVE THIS
    return super.onInit().then(_ => {

      return this.context.aadHttpClientFactory
        .getClient(this.properties.managementApiUrl)
        .then((client: AadHttpClient): void => {
          this.aadHttpClient = client;
        });
    });
  }
  public render(): void {
    const element: React.ReactElement<IAuditLogCaptureManagerProps> = React.createElement(
      AuditLogCaptureManager,
      {
        managementApiUrl: this.properties.managementApiUrl,
        aadHttpClient: this.aadHttpClient
      }
    );

    ReactDom.render(element, this.domElement);
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
