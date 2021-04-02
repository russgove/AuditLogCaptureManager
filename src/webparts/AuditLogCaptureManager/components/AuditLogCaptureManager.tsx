import styles from './AuditLogCaptureManager.module.scss';
import { Subscriptions, ISubscriptionsProps } from "./Subscriptions/Subscriptions";
import { Captures, ICapturesProps } from './Captures/Captures';
import { IAuditLogCaptureManagerProps } from './IAuditLogCaptureManagerProps';
import { IAuditLogCaptureManagerState } from './IAuditLogCaptureManagerState';
import { escape } from '@microsoft/sp-lodash-subset';
import { IToolbarProps, TActionGroups, Toolbar } from '@pnp/spfx-controls-react/lib/controls/toolbar';
import * as React from 'react';

export const CutomPropertyContext: any = React.createContext(undefined);
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
      default:
        content = <div>no action selected</div>;
    }

    return (
      <div className={styles.AuditLogCaptureManager}>
        <CutomPropertyContext.Provider value={this.props}>
          <Toolbar

            actionGroups={{
              'group1': {
                'action1': {
                  title: 'Captures',

                  iconName: 'Edit',
                  onClick: () => {
                    this.setState((current) => ({ ...current, currentAction: "Captures" }));
                  }
                },
                'action2': {
                  title: 'Subscriptions',
                  iconName: 'Add',
                  onClick: () => {
                    this.setState((current) => ({ ...current, currentAction: "Subscriptions" }));
                  }

                }
              }
            }} />

          {content}
        </CutomPropertyContext.Provider>
      </div>
    );
  }
}
