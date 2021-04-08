import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';

import { Subscription } from '../../model/Model';
import { fetchAZFunc } from '../../utilities/fetchApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';
import { useQuery } from 'react-query';
export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface IAvailableContentProps {

}
export const AvailableContent: React.FunctionComponent<IAvailableContentProps> = (props) => {
  const [selectedDate, setSelectedDate] = useState<Date>(new Date());
  const { isLoading, error, data } = useQuery<Subscription>('repoData', () => {
    var date = new Date();
    const url = `${parentContext.managementApiUrl}/api/ListAvailableContent/${selectedDate.getFullYear()}-${selectedDate.getMonth() + 1}-${selectedDate.getDate()}`;
    return fetchAZFunc(parentContext.aadHttpClient, url, "GET");
  });

  const parentContext: any = React.useContext<any>(CutomPropertyContext);
  const [items, setItems] = useState<Array<Subscription>>();

  const [mode, setMode] = useState<string>("display");
  const [selectedItem, setSelectedItem] = useState<Subscription>(null);
  const fetchMyAPI = useCallback(async (date: Date) => {
    const url = `${parentContext.managementApiUrl}/api/ListAvailableContent/${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
    let response = await fetchAZFunc(parentContext.aadHttpClient, url, "GET");
    debugger;
    setItems(response);
  }, []);

  useEffect(() => {

    fetchMyAPI(new Date());
  }, [fetchMyAPI]);
  const viewFields: IViewField[] = [
    {
      name: 'actions', displayName: 'Actions', render: (item?: any, index?: number) => {
        return <div>
          <i className={getIconClassName('Redo')} onClick={async (e) => {
            debugger;
            const url = `${parentContext.managementApiUrl}/api/EnqueueCallbackItems`;
            const selected = [item];
            var response = await fetchAZFunc(parentContext.aadHttpClient, url, "POST", JSON.stringify(selected));

            alert(`${selected.length} files where queued`);
          }}></i>
          &nbsp;&nbsp;    &nbsp;&nbsp;    &nbsp;&nbsp;
          <i className={getIconClassName('View')} onClick={async (e) => {
            debugger;
            const url = `${parentContext.managementApiUrl}/api/EnqueueCallbackItems`;
            const selected = [item];
            var response = await fetchAZFunc(parentContext.aadHttpClient, url, "POST", JSON.stringify(selected));
            return response;
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
      <DatePicker onSelectDate={(date) => {
        setSelectedDate(date);
      }}></DatePicker>
      <PrimaryButton disabled={!selectedDate} onClick={async (e) => {
        setItems([]);
        fetchMyAPI(selectedDate);

      }}>Get Available Content</PrimaryButton>

      <ListView items={items} viewFields={viewFields}></ListView>

      <Panel type={PanelType.smallFixedFar} headerText="Edit Subscription" isOpen={mode === "Edit"} onDismiss={(e) => {
        setMode("Display");
      }} >

      </Panel>
    </div>
  );
};
