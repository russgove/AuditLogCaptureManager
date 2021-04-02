import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { Panel, PanelType, IPanelProps } from 'office-ui-fabric-react/lib/Panel';
import { PrimaryButton, ButtonType, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { getIconClassName } from '@uifabric/styling';
import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';
import { SubscriptionForm, ISubscriptionFormProps } from './SubscriptionForm';
import { Subscription, Webhook } from '../../model/Model';
import { fetchAZFunc } from '../../utilities/fetchApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';
import { divProperties } from 'office-ui-fabric-react/lib/Utilities';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ISubscriptionsProps {
  description: string;
}
export const Subscriptions: React.FunctionComponent<ISubscriptionsProps> = (props) => {
  debugger;
  const parentContext: any = React.useContext<any>(CutomPropertyContext)
  const viewFields: IViewField[] = [
    { name: 'contentType', minWidth: 250, maxWidth: 90, displayName: 'Content Type', sorting: true, isResizable: true },
    { name: 'status', minWidth: 136, maxWidth: 90, displayName: 'Status', sorting: true, isResizable: true },
    { name: 'webhook.address', minWidth: 136, maxWidth: 300, displayName: 'Callback Address', sorting: true, isResizable: true },
    { name: 'webhook.authId', minWidth: 136, maxWidth: 300, displayName: 'Auth Id', sorting: true, isResizable: true },
    { name: 'webhook.expiration', minWidth: 136, maxWidth: 300, displayName: 'Expriration', sorting: true, isResizable: true },
    { name: 'webhook.status', minWidth: 136, maxWidth: 300, displayName: 'Status', sorting: true, isResizable: true },
    {
      name: 'actions', displayName: 'Actions', render: (item?: any, index?: number) => {
        return <div>
          <i className={getIconClassName('Edit')} onClick={(e) => {
            debugger;
            setMode("Edit");
            setSelectedItem(item);
          }}></i>
        </div>
      }
    }

  ];
  const [items, setItems] = useState<Array<Subscription>>();
  const [mode, setMode] = useState<string>("display");
  // useEffect(() => {
  //   setMode(mode);
  // }, [mode]);
  const [selectedItem, setSelectedItem] = useState<Subscription>(null);
  const fetchMyAPI = useCallback(async () => {
    const url = parentContext.managementApiUrl + "/api/ListSubscriptions";
    let response = await fetchAZFunc(parentContext.aadHttpClient, url);
    debugger;
    setItems(response);
  }, []);

  useEffect(() => {
    debugger;
    fetchMyAPI()
  }, [fetchMyAPI])

  return (
    <div>
      Subscriptions {mode}
      <ListView items={items} viewFields={viewFields}></ListView>

      <Panel type={PanelType.smallFixedFar} headerText="Edit Subscription" isOpen={mode === "Edit"} onDismiss={(e) => {
        setMode("Display")
      }} >
        <SubscriptionForm subscription={selectedItem}
          cancel={(e) => {
            setMode("Display")
          }}
          save={(subscription) => {
            setMode("Display")
          }}

        ></SubscriptionForm>
      </Panel>
    </div>
  );
};
