import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { sp } from "@pnp/sp";
import { ContentTypes, IContentTypes } from "@pnp/sp/content-types";
import { Fields, FieldTypes } from "@pnp/sp/fields";
import { ILists, Lists } from "@pnp/sp/lists";
import { IContextInfo, ISite, Site } from "@pnp/sp/sites";
import { IWebs, Web, Webs } from "@pnp/sp/webs";

import "@pnp/sp/sites";

export async function enqueueCallbackItems(client: AadHttpClient, siteUrl: string, listName: string, managementApiUrl: string): Promise<any> {

    var url: string = decodeURIComponent(siteUrl);
    var rootweb = Web(url);
    try {
        rootweb.lists.getByTitle(listName);
        throw new Error("List already exists");
    }
    catch (e) {
    }
    const newList = await rootweb.lists.add(listName, "Audit Data", 100, true);

    const ct = rootweb.contentTypes.getById("0x0100002cf808dcf34fdfbaf1378b8bcaa777");
    if (!ct) {
        var batch = sp.createBatch();
        //SharePoint Base schema
        newList.list.fields.inBatch(batch).add("Site", "SP.FieldText", { Description: "The GUID of the site where the file or folder accessed by the user is located.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("ItemType", "SP.FieldText", { Description: "The type of object that was accessed or modified. See the ItemType table for details on the types of objects.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("EventSource", "SP.FieldText", { Description: "Identifies that an event occurred in SharePoint. Possible values are SharePoint or ObjectModel.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("SourceName", "SP.FieldText", { Description: "The entity that triggered the audited operation. Possible values are SharePoint or ObjectModel.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("UserAgent", "SP.FieldText", { Description: "Information about the user's client or browser. This information is provided by the client or browser.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("MachineDomainInfo", "SP.FieldText", { Description: "Information about device sync operations. This information is reported only if it's present in the request.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("MachineId", "SP.FieldText", { Description: "Information about device sync operations. This information is reported only if it's present in the request.", FieldTypeKind: 3, Group: "Audit Capture" });

        // "File and folder activities" 
        newList.list.fields.inBatch(batch).add("SiteUrl", "SP.FieldText", { Description: "The URL of the site where the file or folder accessed by the user is located.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("SourceRelativeUrl", "SP.FieldText", { Description: "	The URL of the folder that contains the file accessed by the user. The combination of the values for the SiteURL, SourceRelativeURL, and SourceFileName parameters is the same as the value for the ObjectID property, which is the full path name for the file accessed by the user.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("SourceFileName", "SP.FieldText", { Description: "The name of the file or folder accessed by the user.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("SourceFileExtension", "SP.FieldText", { Description: "The file extension of the file that was accessed by the user. This property is blank if the object that was accessed is a folder.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("DestinationRelativeUrl", "SP.FieldText", { Description: "The URL of the destination folder where a file is copied or moved. The combination of the values for SiteURL, DestinationRelativeURL, and DestinationFileName parameters is the same as the value for the ObjectID property, which is the full path name for the file that was copied. This property is displayed only for FileCopied and FileMoved events.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("DestinationFileName", "SP.FieldText", { Description: "The name of the file that is copied or moved. This property is displayed only for FileCopied and FileMoved events.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("DestinationFileExtension", "SP.FieldText", { Description: "The file extension of a file that is copied or moved. This property is displayed only for FileCopied and FileMoved events.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("UserSharedWith", "SP.FieldText", { Description: "The user that a resource was shared with.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("SharingType", "SP.FieldText", { Description: "The type of sharing permissions that were assigned to the user that the resource was shared with. This user is identified by the UserSharedWith parameter.", FieldTypeKind: 3, Group: "Audit Capture" });

        //SharePoint Sharing schema
        newList.list.fields.inBatch(batch).add("TargetUserOrGroupName", "SP.FieldText", { Description: "	Stores the UPN or name of the target user or group that a resource was shared with.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("TargetUserOrGroupType	", "SP.FieldText", { Description: "Identifies whether the target user or group is a Member, Guest, Group, or Partner.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("EventData", "SP.FieldText", { Description: "Conveys follow-up information about the sharing action that has occurred, such as adding a user to a group or granting edit permissions.", FieldTypeKind: 3, Group: "Audit Capture" });

        //    SharePoint schema
        newList.list.fields.inBatch(batch).add("CustomEvent", "SP.FieldText", { Description: "Optional string for custom events.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("EventData", "SP.FieldText", { Description: "", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("ModifiedProperties", "SP.FieldText", { Description: "Optional payload for custom events.", FieldTypeKind: 3, Group: "Audit Capture" });
        newList.list.fields.inBatch(batch).add("SiteUrl", "SP.FieldText", { Description: "The property is included for admin events, such as adding a user as a member of a site or a site collection admin group. The property includes the name of the property that was modified (for example, the Site Admin group), the new value of the modified property (such the user who was added as a site admin), and the previous value of the modified object.", FieldTypeKind: 3, Group: "Audit Capture" });
    }
    return await (await newList.list.get()).Id;
}