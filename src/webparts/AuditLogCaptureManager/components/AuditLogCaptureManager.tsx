
import { Toolbar } from '@pnp/spfx-controls-react/lib/controls/toolbar';
import * as React from 'react';
import { QueryClient, QueryClientProvider, useQuery } from 'react-query';
import styles from './AuditLogCaptureManager.module.scss';
import { AvailableContent } from './AvalaibleContent/AvailableContent';
import { Captures } from './Captures/Captures';
import { IAuditLogCaptureManagerProps } from './IAuditLogCaptureManagerProps';
import { IAuditLogCaptureManagerState } from './IAuditLogCaptureManagerState';
import { Subscriptions } from "./Subscriptions/Subscriptions";

export const CutomPropertyContext: any = React.createContext<IAuditLogCaptureManagerProps>(undefined);
export default class AuditLogCaptureManager extends React.Component<IAuditLogCaptureManagerProps, IAuditLogCaptureManagerState> {


  public constructor(props: IAuditLogCaptureManagerProps) {
    super(props);
    this.state = { currentAction: "Captures" };
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
      default:
        content = <div>no action selected</div>;
    }

    return (
      <div className={styles.AuditLogCaptureManager}>
        <QueryClientProvider client={this.props.queryClient}>
          <CutomPropertyContext.Provider value={this.props}>
            <Toolbar

              actionGroups={{
                'group1': {
                  'Captures': {
                    title: 'Captures',

                    iconName: 'Edit',
                    onClick: () => {
                      this.setState((current) => ({ ...current, currentAction: "Captures" }));
                    }
                  },
                  'Subscriptions': {
                    title: 'Subscriptions',
                    iconName: 'Add',
                    onClick: () => {
                      this.setState((current) => ({ ...current, currentAction: "Subscriptions" }));
                    }
                  },
                  'AvailableContent': {
                    title: 'AvailableContent',
                    iconName: 'AddReaction',
                    onClick: () => {
                      this.setState((current) => ({ ...current, currentAction: "AvailableContent" }));
                    }
                  }
                }
              }} />

            {content}
          </CutomPropertyContext.Provider>
        </QueryClientProvider>
      </div>
    );
  }
}
