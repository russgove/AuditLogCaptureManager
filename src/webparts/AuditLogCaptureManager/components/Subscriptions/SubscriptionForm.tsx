import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { ITextFieldProps, TextField } from 'office-ui-fabric-react/lib/TextField';
import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';

import { Subscription } from '../../model/Model';
import { fetchAZFunc } from '../../utilities/fetchApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ISubscriptionFormProps {
    subscription: Subscription;
    cancel: (e: any) => void;
}
export const SubscriptionForm: React.FunctionComponent<ISubscriptionFormProps> = (props) => {
    debugger;
    const parentContext: any = React.useContext<any>(CutomPropertyContext);
    const save = async (subscription: Subscription) => {
        debugger;
        console.log(subscription.contentType);
        const url = `${parentContext.managementApiUrl}/api/StartsUBSCRIPTION?ContentType=${subscription.contentType}&address=${subscription["webhook.address"]}&authId=${subscription["webhook.authId"]}&expiration=${subscription["webhook.expiration"]}`;
        let response = await fetchAZFunc(parentContext.aadHttpClient, url, "POST", JSON.stringify(subscription));
        return response;
    };
    const [item, setItem] = useState<Subscription>(props.subscription);
    const [errorMessage, setErrorMessage] = useState<string>("");


    return (
        <div>

            <TextField label="Content Type" value={item.contentType} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, contentType: newValue }));
            }}></TextField>
            <TextField label="Status" value={item.status} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, status: newValue }));
            }}></TextField>
            <TextField label="Address" value={item["webhook.address"]} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, "webhook.address": newValue }));
            }}></TextField>
            <TextField label="AuthId" value={item["webhook.authId"]} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, "webhook.authId": newValue }));
            }}></TextField>
            <TextField label="Expiration" value={item["webhook.expiration"]} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, "webhook.expiration": newValue }));
            }}></TextField>
            <TextField label="Status" value={item["webhook.status"]} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, "webhook.status": newValue }));
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
