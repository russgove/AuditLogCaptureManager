import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Site } from 'microsoft-graph';
import { DefaultButton, IconButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { ComboBox, IComboBoxOption, IComboBoxProps } from 'office-ui-fabric-react/lib/ComboBox';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import * as React from 'react';
import { useState } from 'react';

import { SharePointAuditOperations, SiteToCapture } from '../../model/Model';
import { createCaptureList } from '../../utilities/createCaptureList';
import { fetchAZFunc } from '../../utilities/fetchApi';
import { CutomPropertyContext } from '../AuditLogCaptureManager';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ICaptureFormProps {
    siteToCapture: SiteToCapture;
    cancel: (e: any) => void;
}
export const CaptureForm: React.FunctionComponent<ICaptureFormProps> = (props) => {
    const options: Array<IComboBoxOption> = SharePointAuditOperations.map((sao) => {
        return { key: sao.Operation, text: sao.Description };
    });
    const parentContext: any = React.useContext<any>(CutomPropertyContext);
    const save = async (siteToCapture: SiteToCapture) => {

        const url = `${parentContext.managementApiUrl}/api/AddSiteToCapture?siteUrl=${siteToCapture.siteUrl}&siteId=${siteToCapture.siteId}&eventsToCapture=${siteToCapture.eventsToCapture}&captureToListId=${siteToCapture.captureToListId}&captureToSiteId=${siteToCapture.captureToSiteId}`;
        var response = await fetchAZFunc(parentContext.aadHttpClient, url, "POST", JSON.stringify(siteToCapture));
        return response;
    };
    const [newListName, setnewListName] = useState<string>("");
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

                    if (siteName) {
                        setItem((temp) => ({ ...temp, siteUrl: "", siteId: "" }));
                        setErrorMessage("");
                        const url = `${parentContext.managementApiUrl}/api/GetSPSiteByName/${siteName}`;
                        var response: Site;
                        try {
                            response = await fetchAZFunc(parentContext.aadHttpClient, url, "GET");
                            if (response) {
                                setItem((temp) => ({ ...temp, siteUrl: response.webUrl, siteId: response.id.split(',')[1] }));
                            }
                        }
                        catch (e) {

                            setErrorMessage(e);
                        }

                    }
                }}
            ></TextField>
            <TextField label="Site Url" value={item.siteUrl}
                onChange={(e, newValue) => {
                    setItem((temp) => ({ ...temp, siteUrl: newValue }));
                }}
                onBlur={async () => {
                    if (item.siteUrl) {
                        setItem((temp) => ({ ...temp, siteId: "" }));
                        setErrorMessage("");
                        const url = `${parentContext.managementApiUrl}/api/GetSPSiteByName/${item.siteUrl}`;
                        try {
                            var response: Site = await fetchAZFunc(parentContext.aadHttpClient, url, "GET");
                            if (response) {
                                setItem((temp) => ({ ...temp, siteUrl: response.webUrl, siteId: response.id.split(',')[1], captureToSiteId: response.id.split(',')[1] }));
                            }
                        }
                        catch (e) {

                            setErrorMessage(e);
                        }

                    }
                }
                }
            ></TextField>
            <TextField label="Site ID" value={item.siteId}
                onChange={(e, newValue) => {
                    setItem((temp) => ({ ...temp, siteId: newValue }));
                }}
            ></TextField>
            <TextField label="Events To Capture" value={item.eventsToCapture}
                onChange={(e, newValue) => {
                    setItem((temp) => ({ ...temp, eventsToCapture: newValue }));
                }}
            ></TextField>
            <ComboBox label="Events To Capture" options={options} multiSelect={true}
                text={item.eventsToCapture}
                dropdownWidth={800}
                onChange={(e, newValue) => {
                    var events = item.eventsToCapture ? item.eventsToCapture.split(";") : [];
                    events.push(newValue.key as string);
                    setItem((temp) => ({ ...temp, eventsToCapture: events.join(";") }));

                }}
                onResolveOptions={(e) => {

                    return e;
                }}
            >

            </ComboBox>
            <TextField label="Capture To Site Id" value={item.captureToSiteId} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, captureToSiteId: newValue }));
            }}></TextField>

            <TextField label="Capture To List Id" value={item.captureToListId} onChange={(e, newValue) => {
                setItem((temp) => ({ ...temp, captureToListId: newValue }));
            }}></TextField>


            <TextField label="Capture To List" value={newListName} onChange={(e, newValue) => {
                setnewListName(newValue);
            }}></TextField>
            <IconButton iconProps={{ iconName: "NewFolder" }} onClick={(async (e) => {
                debugger;
                var listId = await createCaptureList(parentContext.aadHttpClient, item.siteUrl, newListName, parentContext.managementApiUrl);
                console.log(listId);
                setItem((temp) => ({ ...temp, captureToListId: listId }));
                debugger;
            })}>Create</IconButton>
            <Label style={{ color: "red" }}>
                {errorMessage}
            </Label>


            <div>
                <PrimaryButton disabled={!item.siteId || !item.siteUrl || !item.eventsToCapture || !item.captureToListId || !item.captureToSiteId} onClick={async (e) => {

                    const resp = await save(item);
                    if (resp.error) {
                        setErrorMessage(resp.error.message);
                    }
                    else {
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
