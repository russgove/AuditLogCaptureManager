import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import { Site } from 'microsoft-graph';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { ITextFieldProps, TextField } from 'office-ui-fabric-react/lib/TextField';
import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';

import { SiteToCapture } from '../../model/Model';
import { fetchAZFunc } from '../../utilities/fetchApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ICaptureFormProps {
    siteToCapture: SiteToCapture;
    cancel: (e: any) => void;
}
export const CaptureForm: React.FunctionComponent<ICaptureFormProps> = (props) => {
    const parentContext: any = React.useContext<any>(CutomPropertyContext);
    const save = async (siteToCapture: SiteToCapture) => {
        debugger;
        const url = `${parentContext.managementApiUrl}/api/AddSiteToCapture?siteUrl=${siteToCapture.siteUrl}&siteId=${siteToCapture.siteId}&eventsToCapture=${siteToCapture.eventsToCapture}&captureToListId=${siteToCapture.captureToListId}&captureToSiteId=${siteToCapture.captureToSiteId}`;
        var response = await fetchAZFunc(parentContext.aadHttpClient, url, "POST", JSON.stringify(siteToCapture));
        return response;
    };
    const [item, setItem] = useState<SiteToCapture>(props.siteToCapture);
    const [siteName, setSiteName] = useState<string>(item.siteUrl ? new URL(decodeURIComponent(item.siteUrl)).pathname.split[2] : "");
    const [errorMessage, setErrorMessage] = useState<string>("");
    return (
        <div>
            <TextField label="Site Name" value={siteName}
                onChange={(e, newValue) => {
                    setSiteName(newValue);
                }}
                onBlur={async () => {
                    setItem((temp) => ({ ...temp, siteUrl: "" }));
                    const url = `${parentContext.managementApiUrl}/api/GetSPSiteByName/${siteName}`;
                    var response: Site = await fetchAZFunc(parentContext.aadHttpClient, url, "GET");
                    debugger;
                    if (response) {
                        setItem((temp) => ({ ...temp, siteUrl: response.webUrl, siteId: response.id.split(',')[1] }));
                    }

                }}


            ></TextField>
            <TextField label="Site Url" value={item.siteUrl} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, siteUrl: newValue }));
            }}></TextField>
            <TextField label="Site ID" value={item.siteId} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, siteId: newValue }));
            }}></TextField>
            <TextField label="Events To Capture" value={item.eventsToCapture} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, eventsToCapture: newValue }));
            }}></TextField>
            <TextField label="Capture To List Id" value={item.captureToListId} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, captureToListId: newValue }));
            }}></TextField>
            <TextField label="Capture To Site Id" value={item.captureToSiteId} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, captureToSiteId: newValue }));
            }}></TextField>

            {errorMessage}
            <div>
                <PrimaryButton disabled={!item.siteId || !item.siteUrl || !item.eventsToCapture || !item.captureToListId || !item.captureToSiteId} onClick={async (e) => {
                    debugger;
                    const resp = await save(item);
                    if (resp.error) {
                        setErrorMessage(resp.error.message);

                    } else {
                        setErrorMessage("");
                        props.cancel(e);
                    }

                }}>Save</PrimaryButton>
                <DefaultButton onClick={(e) => {
                    props.cancel(e);
                }}>Cancel</DefaultButton>
            </div>

        </div>
    );
};
