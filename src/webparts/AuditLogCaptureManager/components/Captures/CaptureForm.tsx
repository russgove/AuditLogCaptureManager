import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
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
        let response = await fetchAZFunc(parentContext.aadHttpClient, url, "POST", JSON.stringify(siteToCapture));
        return response;
    };
    const [item, setItem] = useState<SiteToCapture>(props.siteToCapture);
    const [errorMessage, setErrorMessage] = useState<string>("");


    return (
        <div>

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
                <PrimaryButton onClick={async (e) => {
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
            This operation starts a subscription to the specified content type. If a subscription to the specified content type already exists, this operation is used to:

Update the properties of an active webhook.

Enable a webhook that was disabled because of excessive failed notifications.

Re-enable an expired webhook by specifying a later or null expiration date.

Remove a webhook.
See https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference
        </div>
    );
};
