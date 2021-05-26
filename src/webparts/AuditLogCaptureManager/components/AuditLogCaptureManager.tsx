import { Toolbar } from '@pnp/spfx-controls-react/lib/controls/toolbar';
import * as React from 'react';
import { QueryClientProvider } from 'react-query';
import { ReactQueryDevtools } from 'react-query/devtools';

import styles from './AuditLogCaptureManager.module.scss';
import { AvailableContent } from './AvalaibleContent/AvailableContent';
import { Captures } from './Captures/Captures';
import { CrawledContent } from './CrawledContent/CrawledContent';
import { DateFormatPicker } from './DateFormatPicker';
import { IAuditLogCaptureManagerProps } from './IAuditLogCaptureManagerProps';
import { IAuditLogCaptureManagerState } from './IAuditLogCaptureManagerState';
import { Notifications } from './Notifications/Notifications';
import { Subscriptions } from "./Subscriptions/Subscriptions";

export const CutomPropertyContext: any = React.createContext<IAuditLogCaptureManagerProps>(undefined);
export default class AuditLogCaptureManager extends React.Component<IAuditLogCaptureManagerProps, IAuditLogCaptureManagerState> {
  public constructor(props: IAuditLogCaptureManagerProps) {
    super(props);

    this.state = {
      currentAction: "Captures", selectedDateFormat: "Local", aadHttpClient: this.props.aadHttpClient,
      managementApiUrl: this.props.managementApiUrl, queryClient: this.props.queryClient, auditItemContentTypeId: this.props.auditItemContentTypeId,
      webPartContext: this.props.webPartContext
    };
  }
  public render(): React.ReactElement<IAuditLogCaptureManagerProps> {
    var content;
    switch (this.state.currentAction) {
      case "Captures":
        content = <Captures></Captures>;
        break;
      case "Subscriptions":
        content = <Subscriptions></Subscriptions>;
        break;
      case "Subscriptions":
        content = <Subscriptions></Subscriptions>;
        break;
      case "AvailableContent":
        content = <AvailableContent></AvailableContent>;
        break;
      case "CrawledContent":
        content = <CrawledContent></CrawledContent>;
        break;
      case "Notifications":
        content = <Notifications></Notifications>;
        break;
      default:
        content = <div>no action selected</div>;
    }

    return (
      <div className={styles.AuditLogCaptureManager}>
        <QueryClientProvider client={this.props.queryClient}>
          <ReactQueryDevtools initialIsOpen={true} position='bottom-right' />
          <CutomPropertyContext.Provider value={this.state}>
            <Toolbar
              actionGroups={{
                'group1': {
                  'Captures': {
                    title: '  Capture Points',
                    iconName: 'Edit',
                    onClick: () => {
                      this.setState((current) => ({ ...current, currentAction: "Captures" }));
                    }
                  },
                  'Subscriptions': {
                    title: '  Subscriptions',
                    iconName: 'Add',
                    onClick: () => {
                      this.setState((current) => ({ ...current, currentAction: "Subscriptions" }));
                    }
                  },
                  'AvailableContent': {
                    title: '  Available Content',
                    iconName: 'AddReaction',
                    onClick: () => {
                      this.setState((current) => ({ ...current, currentAction: "AvailableContent" }));
                    }
                  },
                  'CrawledContent': {
                    title: '  Crawled Content',
                    iconName: 'AddReaction',
                    onClick: () => {
                      this.setState((current) => ({ ...current, currentAction: "CrawledContent" }));
                    }
                  },
                  'Notifications': {
                    title: '  Notifications',
                    iconName: 'AddReaction',
                    onClick: () => {
                      this.setState((current) => ({ ...current, currentAction: "Notifications" }));
                    }
                  }
                }
              }} />
            <DateFormatPicker selectedDateFormat={this.state.selectedDateFormat}
              onFormatChange={(e: string) => {
                this.setState((current) => ({ ...current, selectedDateFormat: e }));
              }}></DateFormatPicker>
            {content}
          </CutomPropertyContext.Provider>
        </QueryClientProvider>
      </div>
    );
  }
}
