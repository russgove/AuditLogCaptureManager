import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView'
import { Button } from 'office-ui-fabric-react/lib/Button';
import { Icon } from 'office-ui-fabric-react/lib/icon';
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
        { name: 'eventsToCapture', minWidth: 200, maxWidth: 90, displayName: 'Events to Capture', sorting: true, isResizable: true },
        { name: 'captureToListId', minWidth: 136, maxWidth: 90, displayName: 'Capture To List Id', sorting: true, isResizable: true },
        { name: 'captureToSiteId', minWidth: 136, maxWidth: 90, displayName: 'Capture To Site Id', sorting: true, isResizable: true },
        {
            name: 'ddadss', displayName: 'Actions', render: (item?: any, index?: number) => {
                return <div>
                    <i className="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
                    <Icon iconName="Edit" style={{ color: "Red", width: "130px", height: "330px" }} onClick={(e) => {
                        debugger;
                    }}></Icon>
                    <Button iconName="Edit" style={{ color: "Red", width: "130px", height: "330px" }} onClick={(e) => {
                        debugger;
                    }}>Edit</Button>
                </div>
            }
        }

    ];
    const [captures, setCaptures] = useState<Array<any>>();
    const [mode, setMode] = useState<string>("");
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
