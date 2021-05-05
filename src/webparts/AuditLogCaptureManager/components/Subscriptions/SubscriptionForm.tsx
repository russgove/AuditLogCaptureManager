import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import * as React from 'react';
import { useState } from 'react';
import { useMutation } from 'react-query';

import { Subscription } from '../../model/Model';
import { callManagementApi } from '../../utilities/callManagementApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ISubscriptionFormProps {
    subscription: Subscription;
    cancel: (e: any) => void;
}
export const SubscriptionForm: React.FunctionComponent<ISubscriptionFormProps> = (props) => {

    const parentContext: any = React.useContext<any>(CutomPropertyContext);

    const saveSubscription = useMutation((subscription: Subscription) => {
        const url = `${parentContext.managementApiUrl}/api/StartsUBSCRIPTION?ContentType=${subscription.contentType}&address=${subscription["webhook.address"]}&authId=${subscription["webhook.authId"]}&expiration=${subscription["webhook.expiration"]}`;
        return callManagementApi(parentContext.aadHttpClient, url, "POST", JSON.stringify(subscription));
    }, {
        onSuccess: () => {
            parentContext.queryClient.invalidateQueries('subscriptions');
        }
    });
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

                    try {
                        debugger;
                        saveSubscription.mutateAsync(item)
                            .then(() => {
                                setErrorMessage("");
                                props.cancel(e);
                            })
                            .catch((err) => {
                                setErrorMessage(err.message);
                            });
                    }
                    catch (err) {
                        setErrorMessage(err.message);
                    }
                }}
                >Save</PrimaryButton>
                <PrimaryButton onClick={(e) => {
                    props.cancel(e);
                }}>Cancel</PrimaryButton>
            </div>
            This operation starts a subscription to the specified content type.If a subscription to the specified content type already exists, this operation is used to:

    Update the properties of an active webhook.

    Enable a webhook that was disabled because of excessive failed notifications.

        Re - enable an expired webhook by specifying a later or null expiration date.

    Remove a webhook.
    See https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference
        </div >
    );
};
