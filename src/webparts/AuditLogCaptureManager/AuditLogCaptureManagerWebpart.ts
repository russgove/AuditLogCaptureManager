import { Version } from '@microsoft/sp-core-library';
import { Log } from '@microsoft/sp-core-library';
import { AadHttpClient } from '@microsoft/sp-http';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp";
import * as strings from 'AuditLogCaptureManagerWebPartStrings';
import { initializeIcons } from 'office-ui-fabric-react';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { QueryClient, QueryClientProvider, useQuery } from 'react-query';

import AuditLogCaptureManager from './components/AuditLogCaptureManager';
import { IAuditLogCaptureManagerProps } from './components/IAuditLogCaptureManagerProps';

const LOG_SOURCE: string = 'AuditLogCaptureMananger';
export interface IAuditLogCaptureManagerWebPartProps {
  managementApiUrl: string;
}

export default class AuditLogCaptureManagerWebPart extends BaseClientSideWebPart<IAuditLogCaptureManagerWebPartProps> {

  private aadHttpClient: AadHttpClient;
  public onInit(): Promise<void> {

    Log.info(LOG_SOURCE, 'Initialized TrondocsCommandsCommandSet');

    return super.onInit().then(_ => {
      initializeIcons();
      sp.setup({
        spfxContext: this.context,
        // defaultCachingStore: "session", // or "local"
        // defaultCachingTimeoutSeconds: 30,
        // globalCacheDisable: true // or true to disable caching in case of debugging/testing
      });
      if (!this.properties.managementApiUrl) {
        return;
      }
      return this.context.aadHttpClientFactory
        .getClient(this.properties.managementApiUrl)
        .then((client: AadHttpClient): void => {
          this.aadHttpClient = client;
        })
        .catch(err => {
          alert(err);
        });
    });
  }
  public render(): void {
    debugger;
    const d = sp.web.contentTypes.getById("0x01").get().then((x) => {
      debugger;
    });
    const queryClient = new QueryClient();
    const element: React.ReactElement<IAuditLogCaptureManagerProps> = React.createElement(
      AuditLogCaptureManager,
      {
        managementApiUrl: this.properties.managementApiUrl,
        aadHttpClient: this.aadHttpClient,
        queryClient: queryClient
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
                PropertyPaneTextField('managementApiUrl', {
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
