// ** CSOM STUFF
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
export async function createContentType(siteUrl: string): Promise<string> {

  const context: SP.ClientContext = new SP.ClientContext(decodeURIComponent(siteUrl));
  var itemContentType = context.get_site().get_rootWeb().get_contentTypes().getById("0x01");
  context.load(itemContentType);
  await executeQuery(context)
    .catch((err) => {
      console.log(err);
      debugger;
    });


  var contentTypeCreationInformation = new SP.ContentTypeCreationInformation();
  contentTypeCreationInformation.set_name("Audit Item");
  contentTypeCreationInformation.set_description("Microsoft 365 SharePoint Audit Capture detail record");
  contentTypeCreationInformation.set_parentContentType(itemContentType);

  var newContentType: SP.ContentType = context.get_site().get_rootWeb().get_contentTypes()
    .add(contentTypeCreationInformation);
  await addFields(context, newContentType);
  await executeQuery(context)
    .catch((err) => {
      console.log(err);
    });
  return newContentType.get_stringId();
}
async function addFields(context: SP.ClientContext, newContentType: SP.ContentType) {
  await addTextField(context, newContentType, "CreationTime", "Creation Time", "The date and time in Coordinated Universal Time (UTC) when the user performed the activity.");
  await addTextField(context, newContentType, "UserId", "User Id", "The UPN (User Principal Name) of the user who performed the action (specified in the Operation property) that resulted in the record being logged; for example, my_name@my_domain_name. Note that records for activity performed by system accounts (such as SHAREPOINT\system or NT AUTHORITY\SYSTEM) are also included. In SharePoint, another value display in the UserId property is app@sharepoint. This indicates that the 'user' who performed the activity was an application that has the necessary permissions in SharePoint to perform organization-wide actions (such as search a SharePoint site or OneDrive account) on behalf of a user, admin, or service. For more information, see The app@sharepoint user in audit records.");
  await addTextField(context, newContentType, "Operation", "Operation", "The name of the user or admin activity. For a description of the most common operations/activities, see Search the audit log in the Office 365 Protection Center. For Exchange admin activity, this property identifies the name of the cmdlet that was run. For Dlp events, this can be 'DlpRuleMatch', 'DlpRuleUndo' or 'DlpInfo', which are described under 'DLP schema' below.");
  await addTextField(context, newContentType, "ClientIP", "Client IP", "The IP address of the device that was used when the activity was logged. The IP address is displayed in either an IPv4 or IPv6 address format.For some services, the value displayed in this property might be the IP address for a trusted application (for example, Office on the web apps) calling into the service on behalf of a user and not the IP address of the device used by person who performed the activity. Also, for Azure Active Directory-related events, the IP address isn't logged and the value for the ClientIP property is null.");
  await addTextField(context, newContentType, "ItemType", "Item Type", "The type of object that was accessed or modified. See the ItemType table for details on the types of objects.");
  await addTextField(context, newContentType, "SiteUrl", "Site Url", "The URL of the site where the file or folder accessed by the user is located.");

  await addTextField(context, newContentType, "SourceRelativeUrl", "Source Relative Url", "	The URL of the folder that contains the file accessed by the user. The combination of the values for the SiteURL, SourceRelativeURL, and SourceFileName parameters is the same as the value for the ObjectID property, which is the full path name for the file accessed by the user.");

  await addTextField(context, newContentType, "SourceFileName", "Source File Name", "The name of the file or folder accessed by the user.");


  await addTextField(context, newContentType, "SourceFileExtension", "Source File Extension", "The file extension of the file that was accessed by the user. This property is blank if the object that was accessed is a folder.");


  await addTextField(context, newContentType, "DestinationRelativeUrl", "Destination Relative Url", "The URL of the destination folder where a file is copied or moved. The combination of the values for SiteURL, DestinationRelativeURL, and DestinationFileName parameters is the same as the value for the ObjectID property, which is the full path name for the file that was copied. This property is displayed only for FileCopied and FileMoved events.");


  await addTextField(context, newContentType, "DestinationFileName", "Destination File Name", "The name of the file that is copied or moved. This property is displayed only for FileCopied and FileMoved events.");


  await addTextField(context, newContentType, "DestinationFileExtension", "Destination File Extension", "The file extension of a file that is copied or moved. This property is displayed only for FileCopied and FileMoved events.");


  await addTextField(context, newContentType, "UserSharedWith", "User Shared With", "The user that a resource was shared with.");

  await addTextField(context, newContentType, "SharingType", "Sharing Type", "The type of sharing permissions that were assigned to the user that the resource was shared with. This user is identified by the UserSharedWith parameter.");

  //SharePoint Sharing schema
  await addTextField(context, newContentType, "TargetUserOrGroupName", "Target User Or Group Name", "	Stores the UPN or name of the target user or group that a resource was shared with.");


  await addTextField(context, newContentType, "TargetUserOrGroupType	", "Target User Or Group Type", "Identifies whether the target user or group is a Member, Guest, Group, or Partner.");

  await addTextField(context, newContentType, "AuditItemId", "Audit Item Id", "Unique identifier of an audit record.");

  await addTextField(context, newContentType, "CorrelationId", "Correlation Id", "");

  await addTextField(context, newContentType, "ListId", "List Id", "");


  await addTextField(context, newContentType, "WebId", "Web Id", "");

  await addTextField(context, newContentType, "ListItemUniqueId", "List Item Unique Id", "");

  await addNumberField(context, newContentType, "RecordType", "Record Type", "The type of operation indicated by the record. See the AuditLogRecordType table for details on the types of audit log records.");


  await addTextField(context, newContentType, "OrganizationId", "Organization Id", "The GUID for your organization's Office 365 tenant. This value will always be the same for your organization, regardless of the Office 365 service in which it occurs.");

  await addNumberField(context, newContentType, "UserType", "User Type", "");

  await addTextField(context, newContentType, "Version", "Version", "");

  await addTextField(context, newContentType, "UserKey", "User Key", "An alternative ID for the user identified in the UserId property. For example, this property is populated with the passport unique ID (PUID) for events performed by users in SharePoint, OneDrive for Business, and Exchange. This property may also specify the same value as the UserID property for events occurring in other services and events performed by system accounts.");

  await addTextField(context, newContentType, "Workload", "Workload", "The Office 365 service where the activity occurred.");

  await addTextField(context, newContentType, "ResultStatus", "Result Status", "Indicates whether the action (specified in the Operation property) was successful or not. Possible values are Succeeded, PartiallySucceeded, or Failed. For Exchange admin activity, the value is either True or False. Important: Different workloads may overwrite the value of the ResultStatus property. For example, for Azure Active Directory STS logon events, a value of Succeeded for ResultStatus indicates only that the HTTP operation was successful; it doesn't mean the logon was successful. To determine if the actual logon was successful or not, see the LogonError property in the Azure Active Directory STS Logon schema. If the logon failed, the value of this property will contain the reason for the failed logon attempt.");

  await addTextField(context, newContentType, "ObjectId", "Object Id", "For SharePoint and OneDrive for Business activity, the full path name of the file or folder accessed by the user. For Exchange admin audit logging, the name of the object that was modified by the cmdlet.");



  await addTextField(context, newContentType, "Scope", "Scope", "Was this event created by a hosted O365 service or an on-premises server? Possible values are online and onprem. Note that SharePoint is the only workload currently sending events from on-premises to O365");

  //SharePoint Base schema
  await addTextField(context, newContentType, "Site", "Site Id", "The GUID of the site where the file or folder accessed by the user is located.");




  await addTextField(context, newContentType, "EventSource", "Event Source", "Identifies that an event occurred in SharePoint. Possible values are SharePoint or ObjectModel.");


  await addTextField(context, newContentType, "SourceName", "Source Name", "The entity that triggered the audited operation. Possible values are SharePoint or ObjectModel.");


  await addTextField(context, newContentType, "UserAgent", "User Agent", "Information about the user's client or browser. This information is provided by the client or browser.");


  await addTextField(context, newContentType, "MachineDomainInfo", "Machine Domain Info", "Information about device sync operations. This information is reported only if it's present in the request.");

  await addTextField(context, newContentType, "MachineId", "Machine Id", "Information about device sync operations. This information is reported only if it's present in the request.");






  // "File and folder activities" 
  await addTextField(context, newContentType, "EventData", "Event Data", "Conveys follow-up information about the sharing action that has occurred, such as adding a user to a group or granting edit permissions.");


  //    SharePoint schema
  await addTextField(context, newContentType, "CustomEvent", "Custom Event", "Optional string for custom events.");

  await addTextField(context, newContentType, "ModifiedProperties", "Modified Properties", "Optional payload for custom events.");
}

async function addTextField(context: SP.ClientContext, newContentType: SP.ContentType, fieldName: string, displayName: string, description: string): Promise<any> {
  var fieldXML = getTextFieldSchema(fieldName, description, displayName);
  var field = context.get_site().get_rootWeb().get_fields().addFieldAsXml(fieldXML, true, SP.AddFieldOptions.addFieldToDefaultView);
  context.load(field);
  var fieldLink = new SP.FieldLinkCreationInformation();
  fieldLink.set_field(field);
  fieldLink.get_field().set_required(false);
  fieldLink.get_field().set_hidden(false);
  newContentType.get_fieldLinks().add(fieldLink);
  newContentType.update(false);
  return;
}
function getTextFieldSchema(fieldName: string, displayName: string, description: string): string {
  return `<Field Type="Text" Name="${fieldName}" DisplayName="${description}" Required="FALSE" Group="_Audit Columns" />`;
}

async function addNumberField(context: SP.ClientContext, newContentType: SP.ContentType, fieldName: string, displayName: string, description: string): Promise<any> {
  var fieldXML = getNumberFieldSchema(fieldName, description, displayName);
  var field = context.get_site().get_rootWeb().get_fields().addFieldAsXml(fieldXML, true, SP.AddFieldOptions.addFieldToDefaultView);
  var fieldLink = new SP.FieldLinkCreationInformation();
  fieldLink.set_field(field);
  fieldLink.get_field().set_required(false);
  fieldLink.get_field().set_hidden(false);
  newContentType.get_fieldLinks().add(fieldLink);
  return;
}

function getNumberFieldSchema(fieldName: string, displayName: string, description: string): string {
  return `<Field Type="Number" Name="${fieldName}" DisplayName="${description}" Required="FALSE" Group="_Audit Columns" />`;
}

function executeQuery(context: SP.ClientContext): Promise<any> {


  const promise: Promise<any> = new Promise<any>((resolve, reject) => {
    try {
      context.executeQueryAsync(
        (sender: any, args: SP.ClientRequestSucceededEventArgs) => {
          debugger;
          return resolve(args);
        },
        (sender: any, err: SP.ClientRequestFailedEventArgs) => {
          debugger;
          alert(err.get_message());
          console.timeLog(err.get_errorDetails());
          return reject(err.get_message());
        }
      );
    }
    catch (err) {
      debugger;
      console.log(err);
      debugger;

    }

  });
  return promise;
}