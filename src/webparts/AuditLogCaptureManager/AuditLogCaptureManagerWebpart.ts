import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'AuditLogCaptureManagerWebPartStrings';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import AuditLogCaptureManager from './components/AuditLogCaptureManager';
import { IAuditLogCaptureManagerProps } from './components/IAuditLogCaptureManagerProps';
import { IAuditLogCaptureManagerState } from './components/IAuditLogCaptureManagerState';
export interface IAuditLogCaptureManagerWebPartProps {
  description: string;
}

export default class AuditLogCaptureManagerWebPart extends BaseClientSideWebPart<IAuditLogCaptureManagerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAuditLogCaptureManagerProps> = React.createElement(
      AuditLogCaptureManager,
      {
        description: this.properties.description
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
