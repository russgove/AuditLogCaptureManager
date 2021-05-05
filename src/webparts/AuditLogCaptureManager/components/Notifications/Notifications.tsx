import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import * as React from 'react';
import { useState, useRef } from 'react';
import { useQuery } from 'react-query';
import { renderDate } from '../../utilities/renderDate';
import { Notification } from '../../model/Model';
import { callManagementApi } from '../../utilities/callManagementApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';
import { DateFormatPicker } from '../DateFormatPicker';
import { IAuditLogCaptureManagerState } from '../IAuditLogCaptureManagerState';
export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export const Notifications: React.FunctionComponent = () => {

  const [selectedDate, setSelectedDate] = useState<Date>(new Date());

  const crawledCallbackItems = useQuery<Notification[]>('notifications', () => {
    var now = new Date();
    const url = `${parentContext.managementApiUrl}/api/ListNotifications/Audit.SharePoint/${selectedDate.getFullYear()}-${selectedDate.getMonth() + 1}-${selectedDate.getDate()}T${now.getHours()}:${now.getMinutes()}:${now.getSeconds()}`;
    return callManagementApi(parentContext.aadHttpClient, url, "GET");
  },
    { refetchOnWindowFocus: false, enabled: false }
  );
  const parentContext: IAuditLogCaptureManagerState = React.useContext<IAuditLogCaptureManagerState>(CutomPropertyContext);


  const viewFieldsNotifications: IViewField[] = [
    {
      name: 'notificationSent', minWidth: 100, maxWidth: 200, displayName: 'Notification Sent', sorting: true,
      render: renderDate(parentContext.selectedDateFormat), isResizable: true
    },
    { name: 'notificationStatus', minWidth: 50, maxWidth: 100, displayName: 'Status', sorting: true, isResizable: true },
    { name: 'contentType', minWidth: 75, maxWidth: 200, displayName: 'Content Type', sorting: true, isResizable: true },
    {
      name: 'contentCreated', minWidth: 100, maxWidth: 200, displayName: 'Content Created', sorting: true,
      render: renderDate(parentContext.selectedDateFormat), isResizable: true
    },
    {
      name: 'contentExpiration', minWidth: 100, maxWidth: 200, displayName: 'Expires', sorting: true,
      render: renderDate(parentContext.selectedDateFormat), isResizable: true
    },
    { name: 'contentId', minWidth: 40, maxWidth: 300, displayName: 'ID', sorting: true, isResizable: true },
    { name: 'contentUri', minWidth: 40, maxWidth: 500, displayName: 'Content Uri', sorting: true, isResizable: true },


  ];

  return (
    <div>
      Notifications
      <DatePicker value={selectedDate}
        onSelectDate={(date) => {
          setSelectedDate(date);
        }}></DatePicker>
      <PrimaryButton disabled={!selectedDate || crawledCallbackItems.isFetching}
        onClick={async (e) => {

          crawledCallbackItems.refetch();

        }}>Get Notifications</PrimaryButton>

      <ListView items={crawledCallbackItems.data} viewFields={viewFieldsNotifications}></ListView>

    </div>
  );
};
