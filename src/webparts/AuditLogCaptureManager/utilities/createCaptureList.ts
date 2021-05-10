
import { Web } from "@pnp/sp/webs";


import { IAuditLogCaptureManagerProps } from '../components/IAuditLogCaptureManagerProps';
import { createContentType } from './createContentTypeViaApi';

import "@pnp/sp/presets/all";
import "@pnp/sp/sites";

export async function createCaptureList(parentContext: IAuditLogCaptureManagerProps, siteUrl: string, listName: string, managementApiUrl: string): Promise<any> {
    debugger;
    const ctId = "0x0100002CF808DCF34FDFBAF1378B8bCAA777";
    try {
        var url: string = decodeURIComponent(siteUrl);
        var rootweb = Web(url);
        debugger;
        const found = await (await rootweb.contentTypes.getById(ctId)()).Id;
        debugger;
        // var ctId: string = await rootweb.contentTypes.get().
        //     then((cts) => {
        //         return find(cts, (ct) => { return ct.Name === "Audit Item" }).Id.StringValue;
        //     });

        debugger;
        if (!found) {

            try {
                debugger;
                var ctAddResult = await createContentType(siteUrl, parentContext);
                debugger;
                const ct = await rootweb.contentTypes.getById(ctId)();

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
        debugger;
        const addresult = await newList.list.contentTypes.addAvailableContentType(ctId);

        debugger;

        const list = await newList.list.get();
        debugger;
        return list.Id;
    }
    catch (ee) {
        console.log(ee);
        debugger;
    }
}