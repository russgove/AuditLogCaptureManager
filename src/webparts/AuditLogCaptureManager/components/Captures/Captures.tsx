import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';

import { fetchAZFunc } from '../../utilities/fetchApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ICapturesProps {
    description: string;
}
export const Captures: React.FunctionComponent<ICapturesProps> = (props) => {
    const parentContext: any = React.useContext<any>(CutomPropertyContext);
    const [captures, setCaptures] = useState<Array<any>>();
    const [mode, setMode] = useState<string>("");
    const [selectedItem, setSelectedItem] = useState<string>(""); const viewFields: IViewField[] = [
        { name: 'siteUrl', minWidth: 250, maxWidth: 90, displayName: 'Site Url', sorting: true, isResizable: true },
        { name: 'siteId', minWidth: 136, maxWidth: 90, displayName: 'Site Id', sorting: true, isResizable: true },
        { name: 'eventsToCapture', minWidth: 200, maxWidth: 90, displayName: 'Events to Capture', sorting: true, isResizable: true },
        { name: 'captureToListId', minWidth: 136, maxWidth: 90, displayName: 'Capture To List Id', sorting: true, isResizable: true },
        { name: 'captureToSiteId', minWidth: 136, maxWidth: 90, displayName: 'Capture To Site Id', sorting: true, isResizable: true },
        {
            name: 'actions', displayName: 'Actions', render: (item?: any, index?: number) => {
                return <div>
                    <i className="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
                    <i className={getIconClassName('Edit')} onClick={(e) => {
                        setMode("Edit");
                    }}></i>

                </div>;
            }
        }

    ];

    const fetchMyAPI = useCallback(async () => {
        const url = parentContext.managementApiUrl + "/api/ListSitesToCapture";
        let response = await fetchAZFunc(parentContext.aadHttpClient, url, "GET");
        setCaptures(response);
    }, []);
    useEffect(() => {
        fetchMyAPI();
    }, [fetchMyAPI]);

    return (
        <div>
            Events being Captured
            <ListView items={captures} viewFields={viewFields}></ListView>
        </div>
    );
};
