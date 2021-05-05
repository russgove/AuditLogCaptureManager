import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import * as React from 'react';
import { useState } from 'react';
import { useMutation, useQuery } from 'react-query';

import { SiteToCapture } from '../../model/Model';
import { callManagementApi } from '../../utilities/callManagementApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';
import { IAuditLogCaptureManagerState } from '../IAuditLogCaptureManagerState';
import { CaptureForm } from './CaptureForm';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);

export const Captures: React.FunctionComponent = () => {
    const sitesToCapture = useQuery<Array<SiteToCapture>>('sitestocapture', () => {
        const url = parentContext.managementApiUrl + "/api/ListSitesToCapture";
        return callManagementApi(parentContext.aadHttpClient, url, "GET");
    });
    const deleteSiteToCapture = useMutation((siteToCapture: SiteToCapture) => {
        const url = `${parentContext.managementApiUrl}/api/DeleteSiteToCapture?siteId=${siteToCapture.siteId}`;
        return callManagementApi(parentContext.aadHttpClient, url, "Get");
    }, {
        onSuccess: () => {
            parentContext.queryClient.invalidateQueries('sitestocapture');
        }
    });
    const parentContext: IAuditLogCaptureManagerState = React.useContext<IAuditLogCaptureManagerState>(CutomPropertyContext);
    const [mode, setMode] = useState<string>("");
    const [selectedItem, setSelectedItem] = useState<SiteToCapture>(null);
    const viewFields: IViewField[] = [
        {
            name: 'actions', displayName: 'Actions', minWidth: 50, maxWidth: 50, isResizable: true, render: (item?: any, index?: number) => {
                return <div>
                    <i className={getIconClassName('Edit')} onClick={(e) => {
                        setMode("Edit");
                        setSelectedItem(item);
                    }}></i>
              &nbsp;&nbsp;    &nbsp;&nbsp;    &nbsp;&nbsp;

                    <i className={getIconClassName('Delete')} onClick={async (e) => {
                        if (confirm("Are You Sure you wanna?")) {
                            deleteSiteToCapture.mutateAsync(item)
                                .catch((err) => {
                                    alert(err);
                                });
                        }

                    }}></i>
                </div>;
            }
        },
        {
            name: 'siteUrl', minWidth: 300, maxWidth: 700, linkPropertyName: 'siteUrl', displayName: 'Site Url', sorting: true, isResizable: true, render: (item?: any, index?: number) => {
                return decodeURIComponent(item.siteUrl);
            }
        },
        { name: 'eventsToCapture', minWidth: 200, maxWidth: 800, displayName: 'Events to Capture', sorting: true, isResizable: true },

        { name: 'siteId', minWidth: 50, maxWidth: 250, displayName: 'Site Id', sorting: true, isResizable: true },
        { name: 'captureToListId', minWidth: 50, maxWidth: 200, displayName: 'Capture To List Id', sorting: true, isResizable: true },
        { name: 'captureToSiteId', minWidth: 50, maxWidth: 200, displayName: 'Capture To Site Id', sorting: true, isResizable: true },
    ];
    return (
        <div>
            Events being Captured
            <br />
            <PrimaryButton onClick={async (e) => {
                setMode("Edit");
                setSelectedItem(new SiteToCapture());

            }}>Add Site</PrimaryButton>
            <ListView items={sitesToCapture.data} viewFields={viewFields}></ListView>

            <Panel type={PanelType.largeFixed} headerText="Configure Site to Capture" isOpen={mode === "Edit"} onDismiss={(e) => {
                setMode("Display");
            }} >
                <CaptureForm siteToCapture={selectedItem}
                    cancel={(e) => {
                        //  fetchMyAPI();
                        setMode("Display");
                    }}
                ></CaptureForm>
            </Panel>

            <PrimaryButton onClick={async (e) => {
                setMode("Edit");
                setSelectedItem(new SiteToCapture());

            }}>Add Site</PrimaryButton>

        </div>

    );
};
