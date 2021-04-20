import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import * as React from 'react';
import { useState } from 'react';
import { useQuery } from 'react-query';

import { Subscription } from '../../model/Model';
import { fetchAZFunc } from '../../utilities/fetchApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';
import { SubscriptionForm } from './SubscriptionForm';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export const Subscriptions: React.FunctionComponent = () => {
  const parentContext: any = React.useContext<any>(CutomPropertyContext);
  const [mode, setMode] = useState<string>("display");
  const [selectedItem, setSelectedItem] = useState<Subscription>(null);
  const subscriptions = useQuery<Array<Subscription>>('subscriptions', () => {
    const url = parentContext.managementApiUrl + "/api/ListSubscriptions";
    return fetchAZFunc(parentContext.aadHttpClient, url, "GET");
  });

  const viewFields: IViewField[] = [
    {
      name: 'actions', displayName: 'Actions', render: (item?: any, index?: number) => {
        return <div>
          <i className={getIconClassName('Edit')} onClick={(e) => {

            setMode("Edit");
            setSelectedItem(item);
          }}></i>
        </div>;
      }
    },
    { name: 'contentType', minWidth: 250, maxWidth: 90, displayName: 'Content Type', sorting: true, isResizable: true },
    { name: 'status', minWidth: 136, maxWidth: 90, displayName: 'Status', sorting: true, isResizable: true },
    { name: 'webhook.address', minWidth: 136, maxWidth: 300, displayName: 'Callback Address', sorting: true, isResizable: true },
    { name: 'webhook.authId', minWidth: 136, maxWidth: 300, displayName: 'Auth Id', sorting: true, isResizable: true },
    { name: 'webhook.expiration', minWidth: 136, maxWidth: 300, displayName: 'Expiration', sorting: true, isResizable: true },
    { name: 'webhook.status', minWidth: 136, maxWidth: 300, displayName: 'Status', sorting: true, isResizable: true },
  ];
  return (
    <div>
      Subscriptions
      <ListView items={subscriptions.data} viewFields={viewFields}></ListView>
      <Panel type={PanelType.smallFixedFar} headerText="Edit Subscription" isOpen={mode === "Edit"} onDismiss={(e) => {
        setMode("Display");
      }} >
        <SubscriptionForm subscription={selectedItem}
          cancel={(e) => {
            subscriptions.refetch();
            setMode("Display");
          }}
        ></SubscriptionForm>
      </Panel>
    </div>
  );
};
