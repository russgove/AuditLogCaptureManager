import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { ComboBox, IComboBoxOption } from 'office-ui-fabric-react/lib/ComboBox';
import { ContextualMenuItemType } from 'office-ui-fabric-react/lib/ContextualMenu';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import * as React from 'react';
import { useState } from 'react';
import { useQuery } from 'react-query';

import { AuditItem, CallbackItem } from '../../model/Model';
import { ContentTypes, IContentType, Subscription } from '../../model/Model';
import { callManagementApi } from '../../utilities/callManagementApi';
import { renderDate } from '../../utilities/renderDate';
import { CutomPropertyContext } from '../AuditLogCaptureManager';
import { IAuditLogCaptureManagerState } from '../IAuditLogCaptureManagerState';
import { CallbackItemECB, CallbackItemECBProps } from './CallbackItemECB';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export const AvailableContent: React.FunctionComponent = () => {
  const [selectedDate, setSelectedDate] = useState<Date>(new Date());
  const [contentType, setContentType] = useState<string>("Audit.SharePoint");
  const callbackItems = useQuery<CallbackItem[]>('callbackitems', () => {

    var now = new Date();
    const url = `${parentContext.managementApiUrl}/api/ListAvailableContent/${contentType}/${selectedDate.getFullYear()}-${selectedDate.getMonth() + 1}-${selectedDate.getDate()}T${now.getHours()}:${now.getMinutes()}:${now.getSeconds()}`;
    return callManagementApi(parentContext.aadHttpClient, url, "GET");
  },
    { refetchOnWindowFocus: false, enabled: false }
  );

  const [selectedCallbackItem, setSelectedCallbackItem] = useState<CallbackItem>(null);
  const auditItems = useQuery<AuditItem[]>(['audititems', selectedCallbackItem], () => {


    const url = `${parentContext.managementApiUrl}/api/FetchAvailableContentItem?contentUri=${encodeURIComponent(selectedCallbackItem.contentUri)}`;
    console.log(url);
    return callManagementApi(parentContext.aadHttpClient, url, "GET");
  },
    { refetchOnWindowFocus: false, enabled: true });

  const parentContext: IAuditLogCaptureManagerState = React.useContext<any>(CutomPropertyContext);
  const [mode, setMode] = useState<string>("display");
  const options: Array<IComboBoxOption> = ContentTypes.map((ct) => {
    return { key: ct.ContentType, text: ct.ContentType, disabled: !ct.Enabled };
  });
  const viewFieldsCallbackItems: IViewField[] = [
    {
      name: 'actions', minWidth: 50, maxWidth: 50, displayName: 'Actions', render: (item?: any, index?: number) => {
        return <div>
          <i className={getIconClassName('Redo')}
            onClick={async (e) => {

              const url = `${parentContext.managementApiUrl}/api/EnqueueCallbackItems`;
              const selected = [item];
              await callManagementApi(parentContext.aadHttpClient, url, "POST", JSON.stringify(selected));
              alert(`${selected.length} files where queued`);
            }}></i>
          &nbsp;&nbsp;    &nbsp;&nbsp;    &nbsp;&nbsp;
          <i className={getIconClassName('View')} onClick={(e) => {

            setSelectedCallbackItem(item);
            setMode("showselected");

          }}></i>
        </div>;
      }
    },
    { name: 'contentType', minWidth: 75, maxWidth: 200, displayName: 'Content Type', sorting: true, isResizable: true },
    {
      name: 'contentCreated', minWidth: 100, maxWidth: 200, displayName: 'Content Created', sorting: true,
      render: renderDate(parentContext.selectedDateFormat), isResizable: true
    },
    {
      name: 'contentExpiration', minWidth: 100, maxWidth: 200, displayName: 'Expires', sorting: true,
      render: renderDate(parentContext.selectedDateFormat), isResizable: true
    },
    { name: 'contentUri', minWidth: 40, maxWidth: 500, displayName: 'Content Uri', sorting: true, isResizable: true },
    { name: 'contentId', minWidth: 40, maxWidth: 300, displayName: 'ID', sorting: true, isResizable: true },


  ];
  const viewFieldsAuditItems: IViewField[] = [
    {
      name: 'actions', minWidth: 50, maxWidth: 50, displayName: 'Actions', render: (item?: any, index?: number) => {
        return <div>
          <i className={getIconClassName('Redo')} onClick={async (e) => {
debugger;
            const url = `${parentContext.managementApiUrl}/api/EnqueueAuditItems`;
            const selected = [item];
            await callManagementApi(parentContext.aadHttpClient, url, "POST", JSON.stringify(selected));
            alert(`${selected.length} items where queued`);
          }}></i>
          &nbsp;&nbsp;    &nbsp;&nbsp;    &nbsp;&nbsp;
          <i className={getIconClassName('View')} onClick={async (e) => {
            // setSelectedCallbackItem(item);
            // setMode("showselected");
            // auditItems.refetch();
          }}></i>
        </div>;
      }
    },
    {
      name: 'CreationTime', minWidth: 150, maxWidth: 300, displayName: 'CreationTime ', sorting: true,
      render: renderDate(parentContext.selectedDateFormat), isResizable: true
    },
    { name: 'UserId', minWidth: 300, maxWidth: 300, displayName: 'UserId ', sorting: true, isResizable: true },
    { name: 'Operation', minWidth: 100, maxWidth: 100, displayName: 'Operation ', sorting: true, isResizable: true },
    { name: 'ObjectId', minWidth: 400, maxWidth: 900, displayName: 'ObjectId ', sorting: true, isResizable: true },
    { name: 'ClientIP', minWidth: 100, maxWidth: 200, displayName: 'ClientIP ', sorting: true, isResizable: true },
    { name: 'ItemType', minWidth: 100, maxWidth: 100, displayName: 'ItemType ', sorting: true, isResizable: true },
    { name: 'SiteUrl', minWidth: 200, maxWidth: 300, displayName: 'SiteUrl  ', sorting: true, isResizable: true },
    { name: 'SourceFileName', minWidth: 200, maxWidth: 300, displayName: 'SourceFileName  ', sorting: true, isResizable: true },
    { name: 'SourceRelativeUrl', minWidth: 200, maxWidth: 300, displayName: 'SourceRelativeUrl  ', sorting: true, isResizable: true },
    { name: 'FromApp', minWidth: 100, maxWidth: 100, displayName: 'FromApp  ', sorting: true, isResizable: true },
    { name: 'UserType', minWidth: 200, maxWidth: 400, displayName: 'UserType ', sorting: true, isResizable: true },
    { name: 'UserKey', minWidth: 200, maxWidth: 400, displayName: 'UserKey ', sorting: true, isResizable: true },
    { name: 'UserAgent', minWidth: 400, maxWidth: 600, displayName: 'UserAgent ', sorting: true, isResizable: true },


    { name: 'Id', minWidth: 100, maxWidth: 200, displayName: 'Id ', sorting: true, isResizable: true },
    { name: 'OrganizationId', minWidth: 100, maxWidth: 200, displayName: 'OrganizationId ', sorting: true, isResizable: true },
    { name: 'RecordType', minWidth: 100, maxWidth: 200, displayName: 'RecordType ', sorting: true, isResizable: true },
    { name: 'Version', minWidth: 100, maxWidth: 300, displayName: 'Version', sorting: true, isResizable: true },
    { name: 'Workload', minWidth: 100, maxWidth: 300, displayName: 'Workload ', sorting: true, isResizable: true },
    { name: 'CorrelationId', minWidth: 100, maxWidth: 300, displayName: 'CorrelationId ', sorting: true, isResizable: true },
    { name: 'CustomUniqueId', minWidth: 100, maxWidth: 300, displayName: 'CustomUniqueId ', sorting: true, isResizable: true },
    { name: 'EventSource', minWidth: 100, maxWidth: 300, displayName: 'EventSource ', sorting: true, isResizable: true },
    { name: 'ListId', minWidth: 100, maxWidth: 300, displayName: 'ListId ', sorting: true, isResizable: true },
    { name: 'ListItemUniqueId', minWidth: 100, maxWidth: 300, displayName: ' ', sorting: true, isResizable: true },
    { name: 'Site', minWidth: 100, maxWidth: 300, displayName: 'Site ', sorting: true, isResizable: true },
    { name: 'WebId', minWidth: 100, maxWidth: 300, displayName: 'WebId  ', sorting: true, isResizable: true },
    { name: 'SourceFileExtension', minWidth: 100, maxWidth: 300, displayName: 'SourceFileExtension  ', sorting: true, isResizable: true },
    { name: 'HighPriorityMediaProcessing', minWidth: 100, maxWidth: 300, displayName: 'HighPriorityMediaProcessing  ', sorting: true, isResizable: true },
    { name: 'DoNotDistributeEvent', minWidth: 100, maxWidth: 300, displayName: 'DoNotDistributeEvent  ', sorting: true, isResizable: true },
    { name: 'IsDocLib', minWidth: 200, maxWidth: 100, displayName: 'IsDocLib  ', sorting: true, isResizable: true },


  ];




  return (

    <div>
      AvailableContent
      <ComboBox label="Content Type" options={options}

        text={contentType}
        // onRenderOption={(option): JSX.Element => {
        //   return (
        //     <div>
        //       <b>{option.key}</b>--{option.text}
        //     </div>
        //   );
        // }}
        selectedKey={contentType}
        //dropdownWidth={}
        onChange={(e, newValue) => {
          setContentType(newValue.text);


        }}
        onResolveOptions={(e) => {
          return e;
        }}
      />
      <DatePicker value={selectedDate}
       label="Date"
        onSelectDate={(date) => {
          setSelectedDate(date);
        }}></DatePicker>
      <PrimaryButton disabled={!selectedDate || callbackItems.isFetching}
        onClick={async (e) => {

          callbackItems.refetch();

        }}>Get Available Content</PrimaryButton>



      <ListView items={callbackItems.data} viewFields={viewFieldsCallbackItems}></ListView>

      <Panel type={PanelType.extraLarge}
        headerText="Audit Items"
        isOpen={mode === "showselected"}
        onDismiss={(e) => {
          setMode("Display");
        }} >
        <ListView
          items={auditItems.data}
          viewFields={viewFieldsAuditItems}
        //  stickyHeader={true}
        ></ListView>
      </Panel>
    </div>
  );
};
