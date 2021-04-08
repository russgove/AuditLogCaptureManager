import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DefaultButton, IconButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { ComboBox, IComboBoxOption } from 'office-ui-fabric-react/lib/ComboBox';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import * as React from 'react';
import { useState } from 'react';
import { useMutation, useQuery } from 'react-query';

import { SharePointAuditOperations, SiteToCapture } from '../../model/Model';
import { createCaptureList } from '../../utilities/createCaptureList';
import { fetchAZFunc } from '../../utilities/fetchApi';
import { getList } from '../../utilities/getList';
import { getSite } from '../../utilities/getSite';
import { CutomPropertyContext } from '../AuditLogCaptureManager';
import { IAuditLogCaptureManagerProps } from '../IAuditLogCaptureManagerProps';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ICaptureFormProps {
    siteToCapture: SiteToCapture;
    cancel: (e: any) => void;
}
export const CaptureForm: React.FunctionComponent<ICaptureFormProps> = (props) => {
    const options: Array<IComboBoxOption> = SharePointAuditOperations.map((sao) => {
        return { key: sao.Operation, text: sao.Description };
    });
    const parentContext: IAuditLogCaptureManagerProps = React.useContext<IAuditLogCaptureManagerProps>(CutomPropertyContext);
    const saveSiteToCapture = useMutation((siteToCapture: SiteToCapture) => {
        const url = `${parentContext.managementApiUrl}/api/AddSiteToCapture?siteUrl=${siteToCapture.siteUrl}&siteId=${siteToCapture.siteId}&eventsToCapture=${siteToCapture.eventsToCapture}&captureToListId=${siteToCapture.captureToListId}&captureToSiteId=${siteToCapture.captureToSiteId}`;
        return fetchAZFunc(parentContext.aadHttpClient, url, "POST", JSON.stringify(siteToCapture));
    }, {
        onSuccess: () => {
            parentContext.queryClient.invalidateQueries('sitestocapture');
        }
    });
    const [item, setItem] = useState<SiteToCapture>(props.siteToCapture);
    const siteLookup = useQuery<any>(['siteLookup', item.siteUrl], (x) => {
        debugger;
        setErrorMessage("");
        return getSite(item.siteUrl);
    }, {
        refetchOnWindowFocus: false,
        enabled: false, // turned off by default, manual refetch is needed
        onSuccess: (response) => {
            debugger;
            setItem((temp) => ({
                ...temp,
                siteId: response.Id,
                captureToSiteId: response.Id
            }));
        }
        , onError: ((err: string) => {
            setErrorMessage(err);
        })
    });
    const listLookup = useQuery<any>(['listLookup', item.captureToListId], (x) => {
        debugger;
        setErrorMessage("");
        return getList(item.siteUrl, item.captureToListId);
    }, {
        onSuccess: (response) => {
            debugger;
            setList(response);
        }
        , onError: ((err: any) => {
            setErrorMessage(err.Message);
        })
    });
    const [newListName, setnewListName] = useState<string>("");
    const [list, setList] = useState<any>();
    const [errorMessage, setErrorMessage] = useState<string>("");
    return (
        <div>
            <TextField label="Site Url" value={decodeURIComponent(item.siteUrl)}
                onChange={(e, newValue) => {
                    setItem((temp) => ({ ...temp, siteUrl: newValue }));
                }}
                onBlur={async () => {
                    debugger;
                    siteLookup.refetch();

                }}
            ></TextField>
            <TextField label="Site ID" value={item.siteId}
                onChange={(e, newValue) => {
                    setItem((temp) => ({ ...temp, siteId: newValue }));
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

            <TextField label="Capture To List" value={list ? list.Title : ""} ></TextField>
            <TextField label="New Capture To List" value={newListName} onChange={(e, newValue) => {
                setnewListName(newValue);
            }}></TextField>
            <TextField label="Capture To List" value={newListName} onChange={(e, newValue) => {
                setnewListName(newValue);
            }}></TextField>
            <IconButton iconProps={{ iconName: "NewFolder" }} text="Crate" label="DDDD" onClick={(async (e) => {
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
                    try {
                        debugger;
                        saveSiteToCapture.mutateAsync(item)
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
                }}>Save</PrimaryButton>
                <DefaultButton onClick={(e) => {
                    props.cancel(e);
                }}>Cancel</DefaultButton>
            </div>

        </div>
    );
};
