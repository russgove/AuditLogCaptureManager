import { ISharePointAuditOperation, SharePointAuditOperations, SiteToCapture } from '../../model/Model';
import { callManagementApi } from '../../utilities/callManagementApi';
import { createCaptureList } from '../../utilities/createCaptureList';
import { getContentType } from '../../utilities/getContentType';
import { getLists } from '../../utilities/getLists';
import { getSite } from '../../utilities/getSite';
import { CutomPropertyContext } from '../AuditLogCaptureManager';
import { IAuditLogCaptureManagerProps } from '../IAuditLogCaptureManagerProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SitePicker } from "@pnp/spfx-controls-react/lib/SitePicker";
import { IconButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { ComboBox, IComboBoxOption } from 'office-ui-fabric-react/lib/ComboBox';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import * as React from 'react';
import { useState } from 'react';
import { useMutation, useQuery } from 'react-query';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ICaptureFormProps {
    siteToCapture: SiteToCapture;
    cancel: (e: any) => void;
}
export const CaptureForm: React.FunctionComponent<ICaptureFormProps> = (props) => {

    // make below a useref
    const options: Array<IComboBoxOption> = SharePointAuditOperations.map((sao) => {
        return { key: sao.Operation, text: sao.Description };
    });
    const parentContext: IAuditLogCaptureManagerProps = React.useContext<IAuditLogCaptureManagerProps>(CutomPropertyContext);

    const saveSiteToCapture = useMutation((siteToCapture: SiteToCapture) => {
        //const url = `${parentContext.managementApiUrl}/api/AddSiteToCapture?siteUrl=${siteToCapture.siteUrl}&siteId=${siteToCapture.siteId}&eventsToCapture=${siteToCapture.eventsToCapture}&captureToListId=${siteToCapture.captureToListId}&captureToSiteId=${siteToCapture.captureToSiteId}`;
        const url = `${parentContext.managementApiUrl}/api/AddSiteToCapture`;
        return callManagementApi(parentContext.aadHttpClient, url, "POST", JSON.stringify(siteToCapture));
    }, {
        onSuccess: () => {
            parentContext.queryClient.invalidateQueries('sitestocapture');
        }
    });
    const [item, setItem] = useState<SiteToCapture>(props.siteToCapture);

    const selectedOptions: Array<JSX.Element> = item.eventsToCapture
        ?
        SharePointAuditOperations.filter((sao) => {

            return item.eventsToCapture.indexOf(sao.Operation) != -1;
        })
            .map((sao) => {
                return (<div>
                    <b>{sao.Operation}</b>-- {sao.Description}
                </div>);
            })
        :
        [];
    // to do, swithc to capture to site 
    const contentTypeLookup = useQuery<any>(['contentTypeLookup', item.siteUrl], (x) => {
        debugger;
        setErrorMessage("");
        if (item.siteUrl) {
            return getContentType(parentContext, item.siteUrl);
        } else {
            return Promise.resolve();
        }
    }, {
        refetchOnWindowFocus: false,
        enabled: true, // turned off by default, manual refetch is needed
        onSuccess: (response) => {

            // setItem((temp) => ({
            //     ...temp,
            //     siteId: response.Id,
            //     captureToSiteId: response.Id
            // }));
        }
        //  onError: ((err: string) => {

        //     setErrorMessage(err);
        // })
    });
    const addContentTypeToSite = useMutation((siteToCapture: SiteToCapture) => {
        const url = `${parentContext.managementApiUrl}/api/AddContentTypeToSite?siteUrl=${siteToCapture.siteUrl}&siteId=${siteToCapture.siteId}&eventsToCapture=${siteToCapture.eventsToCapture}&captureToListId=${siteToCapture.captureToListId}&captureToSiteId=${siteToCapture.captureToSiteId}`;
        return callManagementApi(parentContext.aadHttpClient, url, "POST", JSON.stringify(siteToCapture));
    }, {
        onSuccess: () => {
            debugger;
            contentTypeLookup.refetch();
            //parentContext.queryClient.invalidateQueries('contentTypeLookup');
        }
    });
    const listLookup = useQuery<any>(['listLookup', item.siteUrl], (x) => {
        setErrorMessage("");
        if (item.siteUrl) {
            debugger;
            return getLists(item.siteUrl);
        } else {
            return Promise.resolve;
        }
    }, {
        refetchOnWindowFocus: false,
        enabled: false, // turned off by default, manual refetch is needed;
    });
    const [newListName, setnewListName] = useState<string>("");

    const [errorMessage, setErrorMessage] = useState<string>("");

    return (
        <div>
            <SitePicker
                initialSites={[{ id: item.siteId, url: decodeURIComponent(item.siteUrl), title: "sd" }]}
                context={parentContext.webPartContext}
                label={'Select site to capture'}
                mode={'site'}

                multiSelect={false}

                onChange={(sites) => {
                    setItem((temp) => ({ ...temp, siteUrl: sites[0].url, siteId: sites[0].id, siteTitle: sites[0].title }));
                    console.log(sites);
                }}
                placeholder={'Select sites'}
                searchPlaceholder={'Filter sites'} />
            <TextField label="Site Id" value={item.siteId} readOnly={true} borderless={true}
                onChange={(e, newValue) => {
                    setItem((temp) => ({ ...temp, siteId: newValue }));
                }}
            ></TextField>

            <ComboBox label="Events To Capture" options={options} multiSelect
                disabled={contentTypeLookup.isError}
                text={item.eventsToCapture}
                onRenderOption={(option): JSX.Element => {
                    return (
                        <div>
                            <b>{option.key}</b>--{option.text}
                        </div>
                    );
                }}
                selectedKey={item.eventsToCapture ? item.eventsToCapture.split(";") : []}
                dropdownWidth={800}
                onChange={(e, newValue) => {
                    console.log("onChange");
                    console.log(e, newValue);
                    var events = item.eventsToCapture ? item.eventsToCapture.split(";") : [];
                    if (newValue.selected) {
                        events.push(newValue.key as string);
                    } else {
                        events = events.filter((event) => { return event !== newValue.key; });
                    }
                    console.log(`events are now  ${events.join(";")}`);
                    setItem((temp) => ({ ...temp, eventsToCapture: events.join(";") }));


                }}
                onResolveOptions={(e) => {
                    return e;
                }}
            >
            </ComboBox>
            {selectedOptions}

            <SitePicker
                initialSites={[{ id: item.siteId, url: decodeURIComponent(item.siteUrl), title: "sd" }]}
                context={parentContext.webPartContext}
                label={'Select site to save captured Audit Items to'}
                mode={'site'}

                multiSelect={false}

                onChange={(sites) => {
                    setItem((temp) => ({ ...temp, captureToSiteId: sites[0].id, captureToSiteUrl: sites[0].url, captureToSiteTitle: sites[0].title }));
                    console.log(sites);
                }}
                placeholder={'Select sites'}
                searchPlaceholder={'Filter sites'} />
            <TextField label="Capture To Site Id" value={item.captureToSiteId} readOnly={true} borderless={true}
            // onChange={(e, newValue) => {
            //     setItem((temp) => ({ ...temp, captureToSiteId: newValue }));
            // }}
            ></TextField>
            <div style={{ display: contentTypeLookup.isError ? 'block' : 'none' }}>
                <span style={{ color: "Red" }} >  The Audit Item content type ({parentContext.auditItemContentTypeId}) does not exist on this site.</span>
                <PrimaryButton
                    onClick={async (e) => {
                        debugger;
                        try {
                            addContentTypeToSite.mutateAsync(item)
                                .catch((err) => {
                                    debugger;
                                    setErrorMessage(err.message);
                                });
                        }
                        catch (err) {
                            debugger;
                            setErrorMessage(err.message);
                        }
                    }}
                >Create the Audit Item content type</PrimaryButton>
            </div>
            <div style={{ display: addContentTypeToSite.isLoading ? 'block' : 'none' }}>
                <span style={{ color: "Green" }} >  The Audit Item content type ({parentContext.auditItemContentTypeId}) is being created on this site.</span>
            </div>

            <TextField label="Capture To List Id" value={item.captureToListId} readOnly={true} borderless={true}
            // onChange={(e, newValue) => {
            //     setItem((temp) => ({ ...temp, captureToListId: newValue }));
            // }}
            ></TextField>

            <TextField label="Capture To List" value={""} readOnly={true} borderless={true}></TextField>
            <TextField label="New Capture To List" value={newListName} onChange={(e, newValue) => {
                setnewListName(newValue);
            }}></TextField>

            <IconButton iconProps={{ iconName: "NewFolder" }} text="Crate" label="DDDD" onClick={(async (e) => {
                debugger;
                var listId = await createCaptureList(parentContext, item.siteUrl, newListName, parentContext.managementApiUrl);
                console.log(listId);
                setItem((temp) => ({ ...temp, captureToListId: listId }));
                debugger;
            })}>Create</IconButton>
            <Label style={{ color: "red" }}>
                {errorMessage}
            </Label>


            <div>
                <PrimaryButton disabled={!item.siteId || !item.siteUrl || !item.eventsToCapture || !item.captureToListId || !item.captureToSiteId}
                    onClick={async (e) => {
                        try {

                            saveSiteToCapture.mutateAsync(item)
                                .then(() => {
                                    setErrorMessage("");
                                    props.cancel(e);
                                })
                                .catch((err) => {
                                    debugger;
                                    setErrorMessage(err.message);
                                });
                        }
                        catch (err) {
                            debugger;
                            setErrorMessage(err.message);
                        }
                    }}
                >Save</PrimaryButton>
                <PrimaryButton onClick={(e) => {
                    props.cancel(e);
                }}>Cancel</PrimaryButton>
            </div>

        </div>
    );
};
