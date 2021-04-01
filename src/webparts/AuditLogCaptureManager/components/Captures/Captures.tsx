import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView'
import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';

import { fetchAZFunc } from '../../utilities/fetchApi'
import { CutomPropertyContext } from '../AuditLogCaptureManager'

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ICapturesProps {
    description: string;
}
export const Captures: React.FunctionComponent<ICapturesProps> = (props) => {
    const parentContext: any = React.useContext<any>(CutomPropertyContext)
    const viewFields: IViewField[] = [
        { name: 'siteUrl', minWidth: 250, maxWidth: 90, displayName: 'Site Url', sorting: true, isResizable: true },
        { name: 'siteId', minWidth: 136, maxWidth: 90, displayName: 'Site Id', sorting: true, isResizable: true },
        { name: 'eventsToCapture', minWidth: 30, maxWidth: 90, displayName: 'Events to Capture', sorting: true, isResizable: true },
        { name: 'captureToListId', minWidth: 30, maxWidth: 90, displayName: 'Capture To List Id', sorting: true, isResizable: true },
        { name: 'captureToSiteId', minWidth: 30, maxWidth: 90, displayName: 'Capture To Site Id', sorting: true, isResizable: true },
        {
            name: 'ddadss', displayName: 'Actions', render: (item?: any, index?: number) => {
                return <div>

                </div>
            }
        }

    ];
    const [captures, setCaptures] = useState<Array<any>>();

    const fetchMyAPI = useCallback(async () => {
        const url = parentContext.managementApiUrl + "/api/ListSitesToCapture";
        let response = await fetchAZFunc(parentContext.aadHttpClient, url);
        debugger;
        setCaptures(response);
    }, [])

    useEffect(() => {
        debugger;
        fetchMyAPI()
    }, [fetchMyAPI])
    // React.useEffect(() => {      
    //     async function fetchMyAPI() {
    //         const url = parentContext.managementApiUrl + "/api/ListSitesToCapture"
    //         return await fetchAZFunc(parentContext.aadHttpClient, url);
    //       }
    //    var results=   fetchMyAPI()
    //     debugger;
    //     setCaptures(results)
    // }, []);// empty array only runs when component is mounted
    return (
        <div>
            Events being Captured
            <ListView items={captures} viewFields={viewFields}></ListView>
        </div>
    );
};
