import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import * as React from 'react';
import { useState } from 'react';
import { useQuery } from 'react-query';

import { Notification } from '../../model/Model';
import { fetchAZFunc } from '../../utilities/fetchApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export const Notifications: React.FunctionComponent = () => {
  const [selectedDate, setSelectedDate] = useState<Date>(new Date());
  const crawledCallbackItems = useQuery<Notification[]>('notifications', () => {
    const url = `${parentContext.managementApiUrl}/api/ListNotifications/Audit.SharePoint/${selectedDate.getFullYear()}-${selectedDate.getMonth() + 1}-${selectedDate.getDate()}`;
    return fetchAZFunc(parentContext.aadHttpClient, url, "GET");
  },
    { refetchOnWindowFocus: false, enabled: false }
  );
  const parentContext: any = React.useContext<any>(CutomPropertyContext);


  const viewFieldsNotifications: IViewField[] = [
    { name: 'notificationSent', minWidth: 100, maxWidth: 200, displayName: 'Notification Sent', sorting: true, isResizable: true },
    { name: 'notificationStatus', minWidth: 100, maxWidth: 200, displayName: 'Status', sorting: true, isResizable: true },
    { name: 'contentType', minWidth: 100, maxWidth: 200, displayName: 'Content Type', sorting: true, isResizable: true },
    { name: 'contentCreated', minWidth: 80, maxWidth: 120, displayName: 'Content Created', sorting: true, isResizable: true },
    { name: 'contentExpiration', minWidth: 80, maxWidth: 120, displayName: 'Expires', sorting: true, isResizable: true },
    { name: 'contentUri', minWidth: 40, maxWidth: 500, displayName: 'Content Uri', sorting: true, isResizable: true },
    { name: 'contentId', minWidth: 40, maxWidth: 300, displayName: 'ID', sorting: true, isResizable: true },


  ];

  return (
    <div>
      Crawled Content
      <DatePicker value={selectedDate}
        onSelectDate={(date) => {
          setSelectedDate(date);
        }}></DatePicker>
      <PrimaryButton disabled={!selectedDate || crawledCallbackItems.isFetching}
        onClick={async (e) => {
          debugger;
          crawledCallbackItems.refetch();

        }}>Get Notifications</PrimaryButton>

      <ListView items={crawledCallbackItems.data} viewFields={viewFieldsNotifications}></ListView>

    </div>
  );
};
