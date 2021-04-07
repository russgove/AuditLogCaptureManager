import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';

import { SiteToCapture } from '../../model/Model';
import { fetchAZFunc } from '../../utilities/fetchApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';
import { CaptureForm } from './CaptureForm';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ICapturesProps {

}
export const Captures: React.FunctionComponent<ICapturesProps> = (props) => {
    const parentContext: any = React.useContext<any>(CutomPropertyContext);
    const [captures, setCaptures] = useState<Array<SiteToCapture>>();
    const [mode, setMode] = useState<string>("");
    const [selectedItem, setSelectedItem] = useState<SiteToCapture>(null);
    const fetchMyAPI = useCallback(async () => {
        const url = parentContext.managementApiUrl + "/api/ListSitesToCapture";
        let response = await fetchAZFunc(parentContext.aadHttpClient, url, "GET");
        setCaptures(response);
    }, []);
    useEffect(() => {
        fetchMyAPI();
    }, [fetchMyAPI]);

    const viewFields: IViewField[] = [
        {
            name: 'actions', displayName: 'Actions', render: (item?: any, index?: number) => {
                return <div>
                    <i className={getIconClassName('Edit')} onClick={(e) => {
                        setMode("Edit");
                        setSelectedItem(item);
                    }}></i>
                    {" "}

                    <i className={getIconClassName('Delete')} onClick={async (e) => {
                        if (confirm("Are You Sure you wanna?")) {
                            const url = `${parentContext.managementApiUrl}/api/DeleteSiteToCapture?siteId=${item.siteId}`;
                            let response = await fetchAZFunc(parentContext.aadHttpClient, url, "Get");

                            fetchMyAPI();
                        }

                    }}></i>

                </div>;
            }
        },
        {
            name: 'siteUrl', minWidth: 250, maxWidth: 90, displayName: 'Site Url', sorting: true, isResizable: true, render: (item?: any, index?: number) => {
                return decodeURIComponent(item.siteUrl);
            }
        },
        { name: 'siteId', minWidth: 136, maxWidth: 90, displayName: 'Site Id', sorting: true, isResizable: true },
        { name: 'eventsToCapture', minWidth: 200, maxWidth: 90, displayName: 'Events to Capture', sorting: true, isResizable: true },
        { name: 'captureToListId', minWidth: 136, maxWidth: 90, displayName: 'Capture To List Id', sorting: true, isResizable: true },
        { name: 'captureToSiteId', minWidth: 136, maxWidth: 90, displayName: 'Capture To Site Id', sorting: true, isResizable: true },


    ];


    return (
        <div>
            Events being Captured
            <br />
            <PrimaryButton onClick={async (e) => {
                setMode("Edit");
                setSelectedItem(new SiteToCapture());

            }}>Add Site</PrimaryButton>
            <ListView items={captures} viewFields={viewFields}></ListView>

            <Panel type={PanelType.largeFixed} headerText="Configure Site to Capture" isOpen={mode === "Edit"} onDismiss={(e) => {
                setMode("Display");
            }} >
                <CaptureForm siteToCapture={selectedItem}
                    cancel={(e) => {
                        fetchMyAPI();
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
