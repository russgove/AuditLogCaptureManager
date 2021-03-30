import { escape } from '@microsoft/sp-lodash-subset';
import { IToolbarProps, TActionGroups, Toolbar } from '@pnp/spfx-controls-react/lib/controls/toolbar'
import * as React from 'react';

import styles from './AuditLogCaptureManager.module.scss';
import { IAuditLogCaptureManagerProps } from './IAuditLogCaptureManagerProps';
import { IAuditLogCaptureManagerState } from './IAuditLogCaptureManagerState';
import { Callbacks, ICallbacksProps } from "./Callbacks/Callbacks";
import { Captures, ICapturesProps } from './Captures/Captures';

export default class AuditLogCaptureManager extends React.Component<IAuditLogCaptureManagerProps, IAuditLogCaptureManagerState> {


  public constructor(props: IAuditLogCaptureManagerProps) {
    super(props);
    this.state = { currentAction: "IIS" };
  }
  public render(): React.ReactElement<IAuditLogCaptureManagerProps> {
    var content;
    switch (this.state.currentAction) {
      case "Captures":
        content = <Captures description="SS"></Captures>;
        break;
      case "Callbacks":
        content = <Callbacks description="SS"></Callbacks>
        break;
      default:
        content = <div>no action selected</div>
    }

    return (
      <div className={styles.AuditLogCaptureManager}>
        <Toolbar actionGroups={{
          'group1': {
            'action1': {
              title: 'Captures',

              iconName: 'Edit',
              onClick: () => { this.setState((current) => ({ ...current, currentAction: "Captures" })) }
            },
            'action2': {
              title: 'Callbacks',
              iconName: 'Add',
              onClick: () => { this.setState((current) => ({ ...current, currentAction: "Callbacks" })) }

            }
          }
        }} />

        {content}

      </div>
    );
  }
}
