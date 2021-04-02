import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView';
import { getIconClassName } from '@uifabric/styling';
import { TextField, ITextFieldProps } from 'office-ui-fabric-react/lib/TextField';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';

import { Subscription } from '../../model/Model';
import { fetchAZFunc } from '../../utilities/fetchApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';
export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ISubscriptionFormProps {
    subscription: Subscription;
    cancel: (e: any) => void;
    save: (subs: Subscription) => void;
}
export const SubscriptionForm: React.FunctionComponent<ISubscriptionFormProps> = (props) => {
    debugger;
    const [item, setItem] = useState<Subscription>(props.subscription);
    const parentContext: any = React.useContext<any>(CutomPropertyContext)
    return (
        <div>

            <TextField label="Content Type" value={item.contentType} onChange={(e, newValue) => {
                setItem((item) => ({ ...item, contentType: newValue }));
            }}></TextField>
            <TextField label="Status" value={item.status} onChange={(e, newValue) => {
                setItem((item) => ({ ...item, status: newValue }));
            }}></TextField>
            <TextField label="Address" value={item["webhook.address"]} onChange={(e, newValue) => {
                setItem((item) => ({ ...item, "webhook.address": newValue }));
            }}></TextField>
            <TextField label="AuthId" value={item["webhook.authId"]} onChange={(e, newValue) => {
                setItem((item) => ({ ...item, "webhook.authId": newValue }));
            }}></TextField>
            <TextField label="Expiration" value={item["webhook.expriration"]} onChange={(e, newValue) => {
                setItem((item) => ({ ...item, "webhook.expriration": newValue }));
            }}></TextField>
            <TextField label="Status" value={item["webhook.status"]} onChange={(e, newValue) => {
                setItem((item) => ({ ...item, "webhook.status": newValue }));
            }}></TextField>
            <div>
                <PrimaryButton onClick={(e) => {
                    debugger; props.save(item);
                }}>Save</PrimaryButton>
                <DefaultButton onClick={(e) => {
                    props.cancel(e);
                }}>Cancel</DefaultButton>
            </div>

            {item.contentType}
        </div>
    );
};
