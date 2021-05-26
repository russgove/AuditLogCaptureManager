import { AuditItem, CallbackItem, Subscription } from '../../model/Model';
import { callManagementApi } from '../../utilities/callManagementApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import { IButtonProps, IconButton, Layer } from 'office-ui-fabric-react';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { ContextualMenuItemType } from 'office-ui-fabric-react/lib/ContextualMenu';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import * as React from 'react';
import { useState } from 'react';
import { useQuery } from 'react-query';

const parentContext: any = React.useContext<any>(CutomPropertyContext);
export interface CallbackItemECBProps {
  callbackItem: CallbackItem;
  viewCallback: (x: CallbackItem) => void;
}
export const CallbackItemECB: React.FunctionComponent<CallbackItemECBProps> = (props) => {
  const replay = async (source: string, event) => {

    const url = `${parentContext.managementApiUrl}/api/EnqueueCallbackItems`;
    await callManagementApi(parentContext.aadHttpClient, url, "POST", JSON.stringify([this]));//make it an array
    alert(`1 files where queued`);
  };

  const [panelOpen, setPanelOpen] = useState<boolean>(false);
  return (
    <div >
      <IconButton id='ContextualMenuButton1'

        text=''
        width='30'
        split={false}
        iconProps={{ iconName: 'MoreVertical' }}
        menuIconProps={{ iconName: '' }}
        menuProps={{
          shouldFocusOnMount: true,
          items: [
            {

              key: 'action1',
              name: 'Replay',
              onClick: replay.bind(this, props.callbackItem)
            },
            {
              key: 'divider_1',
              itemType: ContextualMenuItemType.Divider
            },
            {
              key: 'action2',
              name: 'view',
              onClick: props.viewCallback(props.callbackItem)
            },
            {
              key: 'disabled',
              name: 'Disabled action',
              disabled: true,
              onClick: () => console.error('Disabled action should not be clickable.')
            }
          ]
        }} />
    </div>
  );
};
