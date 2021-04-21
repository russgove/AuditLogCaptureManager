import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { sp } from "@pnp/sp";
import { ContentTypes, IContentType, IContentTypeInfo, IContentTypes } from "@pnp/sp/content-types";
import { Fields, FieldTypes } from "@pnp/sp/fields";
import { ILists, Lists } from "@pnp/sp/lists";
import { IContextInfo, ISite, Site } from "@pnp/sp/sites";
import { IWebs, Web, Webs } from "@pnp/sp/webs";
import { find } from 'lodash';

import { createContentType } from './createContentType';

import "@pnp/sp/presets/all";
import "@pnp/sp/sites";

export async function createCaptureList(client: AadHttpClient, siteUrl: string, listName: string, managementApiUrl: string): Promise<any> {

    try {
        var url: string = decodeURIComponent(siteUrl);
        var rootweb = Web(url);
        //const ct: IContentType = await rootweb.contentTypes.getById("0x0100002cf808dcf34fdfbaf1378b8bcaa777").get();
        var ctId: string = await rootweb.contentTypes.get().
            then((cts) => {
                return find(cts, (ct) => { return ct.Name === "Audit Item" }).Id.StringValue;
            });
        debugger;
        if (!ctId) {

            try {
                ctId = await createContentType(siteUrl);
            }
            catch (err) {
                console.log(err);
                debugger;
            }
            //Common  schema }
        }
        // await rootweb.lists.filter(`Title eq '${listName}'`).get()
        //     .then((results) => {
        //         debugger;
        //         if (results.length > 0) {
        //             throw new Error("List already exists");
        //         }
        //     })


        debugger;
        const newList = await rootweb.lists.add(listName, "Audit Data", 100, true);
        const addresult = await newList.list.contentTypes.addAvailableContentType(ctId);

        debugger;

        const list = await newList.list.get();

        return list.Id;
    }
    catch (ee) {
        debugger;
    }
}