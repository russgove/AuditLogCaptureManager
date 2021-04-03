import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';

import { Subscription } from '../../model/Model';
import { fetchAZFunc } from '../../utilities/fetchApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface IAvailableContentProps {

}
export const AvailableContent: React.FunctionComponent<IAvailableContentProps> = (props) => {

  const parentContext: any = React.useContext<any>(CutomPropertyContext);
  const [items, setItems] = useState<Array<Subscription>>();
  const [mode, setMode] = useState<string>("display");
  const [selectedItem, setSelectedItem] = useState<Subscription>(null);
  const fetchMyAPI = useCallback(async () => {
    const url = parentContext.managementApiUrl + "/api/ListAvailableContent";
    let response = await fetchAZFunc(parentContext.aadHttpClient, url, "GET");
    debugger;
    setItems(response);
  }, []);

  useEffect(() => {

    fetchMyAPI();
  }, [fetchMyAPI]);
  const viewFields: IViewField[] = [
    {
      name: 'actions', displayName: 'Actions', render: (item?: any, index?: number) => {
        return <div>
          <i className={getIconClassName('Replay')} onClick={(e) => {

            setMode("Edit");
            setSelectedItem(item);
          }}></i>
        </div>;
      }
    },
    { name: 'contentType', minWidth: 200, maxWidth: 300, displayName: 'Content Type', sorting: true, isResizable: true },

    { name: 'contentCreated', minWidth: 100, maxWidth: 200, displayName: 'Content Created', sorting: true, isResizable: true },
    { name: 'contentExpiration', minWidth: 100, maxWidth: 200, displayName: 'Expires', sorting: true, isResizable: true },
    { name: 'contentUri', minWidth: 240, maxWidth: 390, displayName: 'Content Uri', sorting: true, isResizable: true },

    { name: 'contentId', minWidth: 136, maxWidth: 200, displayName: 'ID', sorting: true, isResizable: true },


  ];


  return (
    <div>
      AvailableContent {mode}
      <ListView items={items} viewFields={viewFields}></ListView>

      <Panel type={PanelType.smallFixedFar} headerText="Edit Subscription" isOpen={mode === "Edit"} onDismiss={(e) => {
        setMode("Display");
      }} >

      </Panel>
    </div>
  );
};
