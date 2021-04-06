import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { sp } from "@pnp/sp";
import { ContentTypes, IContentTypes } from "@pnp/sp/content-types";
import { Fields, FieldTypes } from "@pnp/sp/fields";
import { ILists, Lists } from "@pnp/sp/lists";
import { IContextInfo, ISite, Site } from "@pnp/sp/sites";
import { IWebs, Web, Webs } from "@pnp/sp/webs";

import "@pnp/sp/presets/all";
import "@pnp/sp/sites";

export async function createCaptureList(client: AadHttpClient, siteUrl: string, listName: string, managementApiUrl: string): Promise<any> {
    debugger;
    try {
        var url: string = decodeURIComponent(siteUrl);
        var rootweb = Web(url);

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
        // const ct = await rootweb.contentTypes.getById("0x0100002cf808dcf34fdfbaf1378b8bcaa777").get();
        // debugger;
        // if (!ct) {
        //var batch = sp.createBatch();
        //Common  schema 
        await newList.list.fields.add("AuditItemId", "SP.FieldText", { Description: "Unique identifier of an audit record.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("RecordType", "SP.FieldText", { Description: "The type of operation indicated by the record. See the AuditLogRecordType table for details on the types of audit log records.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("Operation", "SP.FieldText", { Description: "The name of the user or admin activity. For a description of the most common operations/activities, see Search the audit log in the Office 365 Protection Center. For Exchange admin activity, this property identifies the name of the cmdlet that was run. For Dlp events, this can be 'DlpRuleMatch', 'DlpRuleUndo' or 'DlpInfo', which are described under 'DLP schema' below.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("OrganizationId", "SP.FieldText", { Description: "The GUID for your organization's Office 365 tenant. This value will always be the same for your organization, regardless of the Office 365 service in which it occurs.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("UserType", "SP.FieldText", { Description: "The type of user that performed the operation. See the UserType table for details on the types of users.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("UserKey", "SP.FieldText", { Description: "An alternative ID for the user identified in the UserId property. For example, this property is populated with the passport unique ID (PUID) for events performed by users in SharePoint, OneDrive for Business, and Exchange. This property may also specify the same value as the UserID property for events occurring in other services and events performed by system accounts.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("Workload", "SP.FieldText", { Description: "The Office 365 service where the activity occurred.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("ResultStatus", "SP.FieldText", { Description: "Indicates whether the action (specified in the Operation property) was successful or not. Possible values are Succeeded, PartiallySucceeded, or Failed. For Exchange admin activity, the value is either True or False. Important: Different workloads may overwrite the value of the ResultStatus property. For example, for Azure Active Directory STS logon events, a value of Succeeded for ResultStatus indicates only that the HTTP operation was successful; it doesn't mean the logon was successful. To determine if the actual logon was successful or not, see the LogonError property in the Azure Active Directory STS Logon schema. If the logon failed, the value of this property will contain the reason for the failed logon attempt.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("ObjectId", "SP.FieldText", { Description: "For SharePoint and OneDrive for Business activity, the full path name of the file or folder accessed by the user. For Exchange admin audit logging, the name of the object that was modified by the cmdlet.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("UserId", "SP.FieldText", { Description: "The UPN (User Principal Name) of the user who performed the action (specified in the Operation property) that resulted in the record being logged; for example, my_name@my_domain_name. Note that records for activity performed by system accounts (such as SHAREPOINT\system or NT AUTHORITY\SYSTEM) are also included. In SharePoint, another value display in the UserId property is app@sharepoint. This indicates that the 'user' who performed the activity was an application that has the necessary permissions in SharePoint to perform organization-wide actions (such as search a SharePoint site or OneDrive account) on behalf of a user, admin, or service. For more information, see The app@sharepoint user in audit records.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("ClientIP", "SP.FieldText", { Description: "The IP address of the device that was used when the activity was logged. The IP address is displayed in either an IPv4 or IPv6 address format.For some services, the value displayed in this property might be the IP address for a trusted application (for example, Office on the web apps) calling into the service on behalf of a user and not the IP address of the device used by person who performed the activity. Also, for Azure Active Directory-related events, the IP address isn't logged and the value for the ClientIP property is null.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("Scope", "SP.FieldText", { Description: "Was this event created by a hosted O365 service or an on-premises server? Possible values are online and onprem. Note that SharePoint is the only workload currently sending events from on-premises to O365", FieldTypeKind: 3, Group: "Audit Capture" });




        //SharePoint Base schema
        await newList.list.fields.add("Site", "SP.FieldText", { Description: "The GUID of the site where the file or folder accessed by the user is located.", FieldTypeKind: 3, Group: "Audit Capture" });
        debugger;
        await newList.list.fields.add("ItemType", "SP.FieldText", { Description: "The type of object that was accessed or modified. See the ItemType table for details on the types of objects.", FieldTypeKind: 3, Group: "Audit Capture" }); debugger; debugger;
        await newList.list.fields.add("EventSource", "SP.FieldText", { Description: "Identifies that an event occurred in SharePoint. Possible values are SharePoint or ObjectModel.", FieldTypeKind: 3, Group: "Audit Capture" }); debugger;
        await newList.list.fields.add("SourceName", "SP.FieldText", { Description: "The entity that triggered the audited operation. Possible values are SharePoint or ObjectModel.", FieldTypeKind: 3, Group: "Audit Capture" }); debugger;
        await newList.list.fields.add("UserAgent", "SP.FieldText", { Description: "Information about the user's client or browser. This information is provided by the client or browser.", FieldTypeKind: 3, Group: "Audit Capture" }); debugger;
        await newList.list.fields.add("MachineDomainInfo", "SP.FieldText", { Description: "Information about device sync operations. This information is reported only if it's present in the request.", FieldTypeKind: 3, Group: "Audit Capture" }); debugger;
        await newList.list.fields.add("MachineId", "SP.FieldText", { Description: "Information about device sync operations. This information is reported only if it's present in the request.", FieldTypeKind: 3, Group: "Audit Capture" }); debugger;
        await newList.list.fields.add("CreationTime", "SP.FieldText", { Description: "The date and time in Coordinated Universal Time (UTC) when the user performed the activity.", FieldTypeKind: 3, Group: "Audit Capture" }); debugger;



        // "File and folder activities" 
        await newList.list.fields.add("SiteUrl", "SP.FieldText", { Description: "The URL of the site where the file or folder accessed by the user is located.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("SourceRelativeUrl", "SP.FieldText", { Description: "	The URL of the folder that contains the file accessed by the user. The combination of the values for the SiteURL, SourceRelativeURL, and SourceFileName parameters is the same as the value for the ObjectID property, which is the full path name for the file accessed by the user.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("SourceFileName", "SP.FieldText", { Description: "The name of the file or folder accessed by the user.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("SourceFileExtension", "SP.FieldText", { Description: "The file extension of the file that was accessed by the user. This property is blank if the object that was accessed is a folder.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("DestinationRelativeUrl", "SP.FieldText", { Description: "The URL of the destination folder where a file is copied or moved. The combination of the values for SiteURL, DestinationRelativeURL, and DestinationFileName parameters is the same as the value for the ObjectID property, which is the full path name for the file that was copied. This property is displayed only for FileCopied and FileMoved events.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("DestinationFileName", "SP.FieldText", { Description: "The name of the file that is copied or moved. This property is displayed only for FileCopied and FileMoved events.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("DestinationFileExtension", "SP.FieldText", { Description: "The file extension of a file that is copied or moved. This property is displayed only for FileCopied and FileMoved events.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("UserSharedWith", "SP.FieldText", { Description: "The user that a resource was shared with.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("SharingType", "SP.FieldText", { Description: "The type of sharing permissions that were assigned to the user that the resource was shared with. This user is identified by the UserSharedWith parameter.", FieldTypeKind: 3, Group: "Audit Capture" });

        //SharePoint Sharing schema
        await newList.list.fields.add("TargetUserOrGroupName", "SP.FieldText", { Description: "	Stores the UPN or name of the target user or group that a resource was shared with.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("TargetUserOrGroupType	", "SP.FieldText", { Description: "Identifies whether the target user or group is a Member, Guest, Group, or Partner.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("EventData", "SP.FieldText", { Description: "Conveys follow-up information about the sharing action that has occurred, such as adding a user to a group or granting edit permissions.", FieldTypeKind: 3, Group: "Audit Capture" });

        //    SharePoint schema
        await newList.list.fields.add("CustomEvent", "SP.FieldText", { Description: "Optional string for custom events.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("EventData", "SP.FieldText", { Description: "", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("ModifiedProperties", "SP.FieldText", { Description: "Optional payload for custom events.", FieldTypeKind: 3, Group: "Audit Capture" });
        await newList.list.fields.add("SiteUrl", "SP.FieldText", { Description: "The property is included for admin events, such as adding a user as a member of a site or a site collection admin group. The property includes the name of the property that was modified (for example, the Site Admin group), the new value of the modified property (such the user who was added as a site admin), and the previous value of the modified object.", FieldTypeKind: 3, Group: "Audit Capture" });
        // }
        debugger;
        //  const xx = await batch.execute();
        debugger;
        const list = await newList.list.get();
        debugger;
        return list.Id;
    }
    catch (ee) {
        debugger;
    }
}