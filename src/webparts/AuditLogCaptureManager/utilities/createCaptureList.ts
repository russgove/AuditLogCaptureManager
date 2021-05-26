import { Web } from "@pnp/sp/webs";

import { IAuditLogCaptureManagerProps } from '../components/IAuditLogCaptureManagerProps';
import { createContentType } from './createContentTypeViaApi';

import "@pnp/sp/presets/all";
import "@pnp/sp/sites";

export async function createCaptureList(parentContext: IAuditLogCaptureManagerProps, siteUrl: string, listName: string, managementApiUrl: string): Promise<any> {


    try {
        var url: string = decodeURIComponent(siteUrl);
        var rootweb = Web(url);

        const found = await (await rootweb.contentTypes.getById(parentContext.auditItemContentTypeId)()).Id;

        if (!found) {
            alert(`The Audit Item content type (${parentContext.auditItemContentTypeId}) as not found on this site.`);
            return;
        }
        const newList = await rootweb.lists.add(listName, "Audit Data", 100, true);
        const addresult = await newList.list.contentTypes.addAvailableContentType(parentContext.auditItemContentTypeId);
        const list = await newList.list.get();
        return list.Id;
    }
    catch (ee) {
        console.log(ee);
        debugger;
    }
}