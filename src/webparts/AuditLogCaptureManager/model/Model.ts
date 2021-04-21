
export class CallbackItem {
    public clientId: string;
    public contentCreated: string;
    public contentExpiration: string;

    public contentId: string;
    public contentType: string;
    public contentUri: string;
    public tenantId: string;
}

export class Notification {
    public contentExpireation: string;
    public contentCreated: string;
    public contentId: string;

    public contentType: string;
    public contentUri: string;

    public notificationSent: string;
    public notificationStatus: string;

}
export class CrawledCallbackItem {
    public dateCrawled: string;
    public triggeredBy: string;
    public callbackItem: CallbackItem;

}
export class Subscription {
    public contentType: string;
    public status: string;
    public webhook: Webhook;
}
export class Webhook {
    public address: string;
    public authId: string;
    public expiration: string;
    public status: string;
}
export class SiteToCapture {

    public siteUrl: string;
    public siteId: string;
    public eventsToCapture: string;
    public captureToSiteId: string;
    public captureToListId: string;
}

export class AuditItem {
    public CreationTime: string;
    public Id: string;
    public Operation: string;
    public OrganizationId: string;
    public RecordType: number;
    public UserType: number;
    public UserKey: string;
    public Version: number;
    public Workload: string;
    public ClientIP: string;
    public ObjectId: string;

    public UserId: string;
    public CorrelationId: string;
    public CustomUniqueId: boolean;
    public EventSource: string;
    public ItemType: string;
    public ListId: string;
    public ListItemUniqueId: string;
    public Site: string;
    public UserAgent: string;
    public WebId: string;
    public SourceFileExtension: string;
    public SiteUrl: string;
    public SourceFileName: string;
    public SourceRelativeUrl: string;

    public HighPriorityMediaProcessing: boolean;

    public DoNotDistributeEvent: boolean;
    public FromApp: boolean;
    public IsDocLib: boolean;



}
export interface ISharePointAuditOperation {
    Operation: string;
    Description: string;
}
export var SharePointAuditOperations: Array<ISharePointAuditOperation> = [
    { Operation: "AccessInvitationAccepted", Description: "The recipient of an invitation to view or edit a shared file (or folder) has accessed the shared file by clicking on the link in the invitation." },
    { Operation: "AccessInvitationCreated", Description: "User sends an invitation to another person (inside or outside their organization) to view or edit a shared file or folder on a SharePoint or OneDrive for Business site. The details of the event entry identifies the name of the file that was shared, the user the invitation was sent to, and the type of the sharing permission selected by the person who sent the invitation." },
    { Operation: "AccessInvitationExpired", Description: "An invitation sent to an external user expires. By default, an invitation sent to a user outside of your organization expires after 7 days if the invitation isn't accepted." },
    { Operation: "AccessInvitationRevoked", Description: "The site administrator or owner of a site or document in SharePoint or OneDrive for Business withdraws an invitation that was sent to a user outside your organization. An invitation can be withdrawn only before it's accepted." },
    { Operation: "AccessInvitationUpdated", Description: "The user who created and sent an invitation to another person to view or edit a shared file (or folder) on a SharePoint or OneDrive for Business site resends the invitation." },
    { Operation: "AccessRequestApproved", Description: "The site administrator or owner of a site or document in SharePoint or OneDrive for Business approves a user request to access the site or document." },
    { Operation: "AccessRequestCreated", Description: "User requests access to a site or document in SharePoint or OneDrive for Business that they don't have permission to access." },
    { Operation: "AccessRequestRejected", Description: "The site administrator or owner of a site or document in SharePoint declines a user request to access the site or document." },
    { Operation: "ActivationEnabled", Description: "Users can browser-enable form templates that don't contain form code, require full trust, enable rendering on a mobile device, or use a data connection managed by a server administrator." },
    { Operation: "AdministratorAddedToTermStore", Description: "Term store administrator added." },
    { Operation: "AdministratorDeletedFromTermStore", Description: "Term store administrator deleted." },
    { Operation: "AllowGroupCreationSet", Description: "Site administrator or owner adds a permission level to a SharePoint or OneDrive for Business site that allows a user assigned that permission to create a group for that site." },
    { Operation: "AppCatalogCreated", Description: "App catalog created to make custom business apps available for your SharePoint Environment." },
    { Operation: "AuditPolicyRemoved", Description: "Document LifeCycle Policy has been removed for a site collection." },
    { Operation: "AuditPolicyUpdate", Description: "Document LifeCycle Policy has been updated for a site collection." },
    { Operation: "AzureStreamingEnabledSet", Description: "A video portal owner has allowed video streaming from Azure." },
    { Operation: "CollaborationTypeModified", Description: "The type of collaboration allowed on sites (for example, intranet, extranet, or public) has been modified." },
    { Operation: "ConnectedSiteSettingModified", Description: "User has either created, modified or deleted the link between a project and a project site or the user modifies the synchronization setting on the link in Project Web App." },
    { Operation: "CreateSSOApplication", Description: "Target application created in Secure store service." },
    { Operation: "CustomFieldOrLookupTableCreated", Description: "User created a custom field or lookup table/item in Project Web App." },
    { Operation: "CustomFieldOrLookupTableDeleted", Description: "User deleted a custom field or lookup table/item in Project Web App." },
    { Operation: "CustomFieldOrLookupTableModified", Description: "User modified a custom field or lookup table/item in Project Web App." },
    { Operation: "CustomizeExemptUsers", Description: "Global administrator customized the list of exempt user agents in SharePoint admin center. You can specify which user agents to exempt from receiving an entire Web page to index. This means when a user agent you've specified as exempt encounters an InfoPath form, the form will be returned as an XML file instead of an entire Web page. This makes indexing InfoPath forms faster." },
    { Operation: "DefaultLanguageChangedInTermStore*", Description: "Language setting changed in the terminology store." },
    { Operation: "DelegateModified", Description: "User created or modified a security delegate in Project Web App." },
    { Operation: "DelegateRemoved", Description: "User deleted a security delegate in Project Web App." },
    { Operation: "DeleteSSOApplication", Description: "An SSO application was deleted." },
    { Operation: "eDiscoveryHoldApplied", Description: "An In-Place Hold was placed on a content source. In-Place Holds are managed by using an eDiscovery site collection (such as the eDiscovery Center) in SharePoint." },
    { Operation: "eDiscoveryHoldRemoved", Description: "An In-Place Hold was removed from a content source. In-Place Holds are managed by using an eDiscovery site collection (such as the eDiscovery Center) in SharePoint." },
    { Operation: "eDiscoverySearchPerformed", Description: "An eDiscovery search was performed using an eDiscovery site collection in SharePoint." },
    { Operation: "EngagementAccepted", Description: "User accepts a resource engagement in Project Web App." },
    { Operation: "EngagementModified", Description: "User modifies a resource engagement in Project Web App." },
    { Operation: "EngagementRejected", Description: "User rejects a resource engagement in Project Web App." },
    { Operation: "EnterpriseCalendarModified", Description: "User copies, modifies or delete an enterprise calendar in Project Web App." },
    { Operation: "EntityDeleted", Description: "User deletes a timesheet in Project Web App." },
    { Operation: "EntityForceCheckedIn", Description: "User forces a checkin on a calendar, custom field or lookup table in Project Web App." },
    { Operation: "ExemptUserAgentSet", Description: "Global administrator adds a user agent to the list of exempt user agents in the SharePoint admin center." },
    { Operation: "FileAccessed", Description: "User or system account accesses a file on a SharePoint or OneDrive for Business site. System accounts can also generate FileAccessed events." },
    { Operation: "FileCheckOutDiscarded", Description: "User discards (or undos) a checked out file. That means any changes they made to the file when it was checked out are discarded, and not saved to the version of the document in the document library." },
    { Operation: "FileCheckedIn", Description: "User checks in a document that they checked out from a SharePoint or OneDrive for Business document library." },
    { Operation: "FileCheckedOut", Description: "User checks out a document located in a SharePoint or OneDrive for Business document library. Users can check out and make changes to documents that have been shared with them." },
    { Operation: "FileCopied", Description: "User copies a document from a SharePoint or OneDrive for Business site. The copied file can be saved to another folder on the site." },
    { Operation: "FileDeleted", Description: "User deletes a document from a SharePoint or OneDrive for Business site." },
    { Operation: "FileDeletedFirstStageRecycleBin", Description: "User deletes a file from the recycle bin on a SharePoint or OneDrive for Business site." },
    { Operation: "FileDeletedSecondStageRecycleBin", Description: "User deletes a file from the second-stage recycle bin on a SharePoint or OneDrive for Business site." },
    { Operation: "FileDownloaded", Description: "User downloads a document from a SharePoint or OneDrive for Business site." },
    { Operation: "FileFetched", Description: "This event has been replaced by the FileAccessed event, and has been deprecated." },
    { Operation: "FileModified", Description: "User or system account modifies the content or the properties of a document located on a SharePoint or OneDrive for Business site." },
    { Operation: "FileMoved", Description: "User moves a document from its current location on a SharePoint or OneDrive for Business site to a new location." },
    { Operation: "FilePreviewed", Description: "User previews a document on a SharePoint or OneDrive for Business site." },
    { Operation: "FileRenamed", Description: "User renames a document on a SharePoint or OneDrive for Business site." },
    { Operation: "FileRestored", Description: "User restores a document from the recycle bin of a SharePoint or OneDrive for Business site." },
    { Operation: "FileSyncDownloadedFull", Description: "User establishes a sync relationship and successfully downloads files for the first time to their computer from a SharePoint or OneDrive for Business document library." },
    { Operation: "FileSyncDownloadedPartial", Description: "User successfully downloads any changes to files from SharePoint or OneDrive for Business document library. This event indicates that any changes that were made to files in the document library were downloaded to the user's computer. Only changes were downloaded because the document library was previously downloaded by the user (as indicated by the FileSyncDownloadedFull event)." },
    { Operation: "FileSyncUploadedFull", Description: "User establishes a sync relationship and successfully uploads files for the first time from their computer to a SharePoint or OneDrive for Business document library." },
    { Operation: "FileSyncUploadedPartial", Description: "User successfully uploads changes to files on a SharePoint or OneDrive for Business document library. This event indicates that any changes made to the local version of a file from a document library are successfully uploaded to the document library. Only changes are unloaded because those files were previously uploaded by the user (as indicated by the FileSyncUploadedFull event)." },
    { Operation: "FileUploaded", Description: "User uploads a document to a folder on a SharePoint or OneDrive for Business site." },
    { Operation: "FileViewed", Description: "This event has been replaced by the FileAccessed event, and has been deprecated." },
    { Operation: "FolderCopied", Description: "User copies a folder from a SharePoint or OneDrive for Business site to another location in SharePoint or OneDrive for Business." },
    { Operation: "FolderCreated", Description: "User creates a folder on a SharePoint or OneDrive for Business site." },
    { Operation: "FolderDeleted", Description: "User deletes a folder from a SharePoint or OneDrive for Business site." },
    { Operation: "FolderDeletedFirstStageRecycleBin", Description: "User deletes a folder from the recycle bin on a SharePoint or OneDrive for Business site ." },
    { Operation: "FolderDeletedSecondStageRecycleBin", Description: "User deletes a folder from the second-stage recycle bin on a SharePoint or OneDrive for Business site." },
    { Operation: "FolderModified", Description: "User modifies a folder on a SharePoint or OneDrive for Business site. This event includes folder metadata changes, such as tags and properties." },
    { Operation: "FolderMoved", Description: "User moves a folder from a SharePoint or OneDrive for Business site." },
    { Operation: "FolderRenamed", Description: "User renames a folder on a SharePoint or OneDrive for Business site." },
    { Operation: "FolderRestored", Description: "User restores a folder from the Recycle Bin on a SharePoint or OneDrive for Business site." },
    { Operation: "GroupAdded", Description: "Site administrator or owner creates a group for a SharePoint or OneDrive for Business site, or performs a task that results in a group being created. For example, the first time a user creates a link to share a file, a system group is added to the user's OneDrive for Business site. This event can also be a result of a user creating a link with edit permissions to a shared file." },
    { Operation: "GroupRemoved", Description: "User deletes a group from a SharePoint or OneDrive for Business site." },
    { Operation: "GroupUpdated", Description: "Site administrator or owner changes the settings of a group for a SharePoint or OneDrive for Business site. This can include changing the group's name, who can view or edit the group membership, and how membership requests are handled." },
    { Operation: "LanguageAddedToTermStore", Description: "Language added to the terminology store." },
    { Operation: "LanguageRemovedFromTermStore", Description: "Language removed from the terminology store." },
    { Operation: "LegacyWorkflowEnabledSet", Description: "Site administrator or owner adds the SharePoint Workflow Task content type to the site. Global administrators can also enable work flows for the entire organization in the SharePoint admin center." },
    { Operation: "LookAndFeelModified", Description: "User modifies a quick launch, gantt chart formats, or group formats.  Or the user creates, modifies, or deletes a view in Project Web App." },
    { Operation: "ManagedSyncClientAllowed", Description: "User successfully establishes a sync relationship with a SharePoint or OneDrive for Business site. The sync relationship is successful because the user's computer is a member of a domain that's been added to the list of domains (called the safe recipients list) that can access document libraries in your organization. For more information, see Use SharePoint Online PowerShell to enable OneDrive sync for domains that are on the safe recipients list." },
    { Operation: "MaxQuotaModified", Description: "The maximum quota for a site has been modified." },
    { Operation: "MaxResourceUsageModified", Description: "The maximum allowable resource usage for a site has been modified." },
    { Operation: "MySitePublicEnabledSet", Description: "The flag enabling users to have public MySites has been set by the SharePoint administrator." },
    { Operation: "NewsFeedEnabledSet", Description: "Site administrator or owner enables RSS feeds for a SharePoint or OneDrive for Business site. Global administrators can enable RSS feeds for the entire organization in the SharePoint admin center." },
    { Operation: "ODBNextUXSettings", Description: "New UI for OneDrive for Business has been enabled." },
    { Operation: "OfficeOnDemandSet", Description: "Site administrator enables Office on Demand, which lets users access the latest version of Office desktop applications. Office on Demand is enabled in the SharePoint admin center and requires an Office 365 subscription that includes full, installed Office applications." },
    { Operation: "PageViewed", Description: "User views a page on a SharePoint site or OneDrive for Business site. This does not include viewing document library files from a SharePoint site or One Drive for Business site on a browser." },
    { Operation: "PeopleResultsScopeSet", Description: "Site administrator creates or changes the result source for People Searches for a SharePoint site." },
    { Operation: "PermissionSyncSettingModified", Description: "User modifies the project permission sync settings in Project Web App." },
    { Operation: "PermissionTemplateModified", Description: "User creates, modifies or deletes a permissions template in Project Web App." },
    { Operation: "PortfolioDataAccessed", Description: "User accesses portfolio content (driver library, driver prioritization, portfolio analyses) in Project Web App." },
    { Operation: "PortfolioDataModified", Description: "User creates, modifies, or deletes portfolio data (driver library, driver prioritization, portfolio analyses) in Project Web App." },
    { Operation: "PreviewModeEnabledSet", Description: "Site administrator enables document preview for a SharePoint site." },
    { Operation: "ProjectAccessed", Description: "User accesses project content in Project Web App." },
    { Operation: "ProjectCheckedIn", Description: "User checks in a project that they checked out from a Project Web App." },
    { Operation: "ProjectCheckedOut", Description: "User checks out a project located in a Project Web App. Users can check out and make changes to projects that they have permission to open." },
    { Operation: "ProjectCreated", Description: "User creates a project in Project Web App." },
    { Operation: "ProjectDeleted", Description: "User deletes a project in Project Web App." },
    { Operation: "ProjectForceCheckedIn", Description: "User forces a check in on a project in Project Web App." },
    { Operation: "ProjectModified", Description: "User modifies a project in Project Web App." },
    { Operation: "ProjectPublished", Description: "User publishes a project in Project Web App." },
    { Operation: "ProjectWorkflowRestarted", Description: "User restarts a workflow in Project Web App." },
    { Operation: "PWASettingsAccessed", Description: "User access the Project Web App settings via CSOM." },
    { Operation: "PWASettingsModified", Description: "User modifies the a Project Web App configuration." },
    { Operation: "QueueJobStateModified", Description: "User cancels or restarts a queue job in Project Web App." },
    { Operation: "QuotaWarningEnabledModified", Description: "Storage quota warning modified." },
    { Operation: "RenderingEnabled", Description: "Browser-enabled form templates will be rendered by InfoPath forms services." },
    { Operation: "ReportingAccessed", Description: "User accessed the reporting endpoint in Project Web App." },
    { Operation: "ReportingSettingModified", Description: "User modifies the reporting configuration in Project Web App." },
    { Operation: "ResourceAccessed", Description: "User accesses an enterprise resource content in Project Web App." },
    { Operation: "ResourceCheckedIn", Description: "User checks in an enterprise resource that they checked out from Project Web App." },
    { Operation: "ResourceCheckedOut", Description: "User checks out an enterprise resource located in Project Web App." },
    { Operation: "ResourceCreated", Description: "User creates an enterprise resource in Project Web App." },
    { Operation: "ResourceDeleted", Description: "User deletes an enterprise resource in Project Web App." },
    { Operation: "ResourceForceCheckedIn", Description: "User forces a checkin of an enterprise resource in Project Web App." },
    { Operation: "ResourceModified", Description: "User modifies an enterprise resource in Project Web App." },
    { Operation: "ResourcePlanCheckedInOrOut", Description: "User checks in or out a resource plan in Project Web App." },
    { Operation: "ResourcePlanModified", Description: "User modifies a resource plan in Project Web App." },
    { Operation: "ResourcePlanPublished", Description: "User publishes a resource plan in Project Web App." },
    { Operation: "ResourceRedacted", Description: "User redacts an enterprise resource removing all personal information in Project Web App." },
    { Operation: "ResourceWarningEnabledModified", Description: "Resource quota warning modified." },
    { Operation: "SSOGroupCredentialsSet", Description: "Group credentials set in Secure store service." },
    { Operation: "SSOUserCredentialsSet", Description: "User credentials set in Secure store service." },
    { Operation: "SearchCenterUrlSet", Description: "Search center URL set." },
    { Operation: "SecondaryMySiteOwnerSet", Description: "A user has added a secondary owner to their MySite." },
    { Operation: "SecurityCategoryModified", Description: "User creates, modifies or deletes a security category in Project Web App." },
    { Operation: "SecurityGroupModified", Description: "User creates, modifies or deletes a security group in Project Web App." },
    { Operation: "SendToConnectionAdded", Description: "Global administrator creates a new Send To connection on the Records management page in the SharePoint admin center. A Send To connection specifies settings for a document repository or a records center. When you create a Send To connection, a Content Organizer can submit documents to the specified location." },
    { Operation: "SendToConnectionRemoved", Description: "Global administrator deletes a Send To connection on the Records management page in the SharePoint admin center." },
    { Operation: "SharedLinkCreated", Description: "User creates a link to a shared file in SharePoint or OneDrive for Business. This link can be sent to other people to give them access to the file. A user can create two types of links: a link that allows a user to view and edit the shared file, or a link that allows the user to just view the file." },
    { Operation: "SharedLinkDisabled", Description: "User disables (permanently) a link that was created to share a file." },
    { Operation: "SharingInvitationAccepted*", Description: "User accepts an invitation to share a file or folder. This event is logged when a user shares a file with other users." },
    { Operation: "SharingRevoked", Description: "User unshares a file or folder that was previously shared with other users. This event is logged when a user stops sharing a file with other users." },
    { Operation: "SharingSet", Description: "User shares a file or folder located in SharePoint or OneDrive for Business with another user inside their organization." },
    { Operation: "SiteAdminChangeRequest", Description: "User requests to be added as a site collection administrator for a SharePoint site collection. Site collection administrators have full control permissions for the site collection and all subsites." },
    { Operation: "SiteCollectionAdminAdded*", Description: "Site collection administrator or owner adds a person as a site collection administrator for a SharePoint or OneDrive for Business site. Site collection administrators have full control permissions for the site collection and all subsites." },
    { Operation: "SiteCollectionCreated", Description: "Global administrator creates a new site collection in your SharePoint organization." },
    { Operation: "SiteRenamed", Description: "Site administrator or owner renames a SharePoint or OneDrive for Business site" },
    { Operation: "StatusReportModified", Description: "User creates, modifies or deletes a status report in Project Web App." },
    { Operation: "SyncGetChanges", Description: "User clicks Sync in the action tray on in SharePoint or OneDrive for Business to synchronize any changes to file in a document library to their computer." },
    { Operation: "TaskStatusAccessed", Description: "User accesses the status of one or more tasks in Project Web App." },
    { Operation: "TaskStatusApproved", Description: "User approves a status update of one or more tasks in Project Web App." },
    { Operation: "TaskStatusRejected", Description: "User rejects a status update of one or more tasks in Project Web App." },
    { Operation: "TaskStatusSaved", Description: "User saves a status update of one or more tasks in Project Web App." },
    { Operation: "TaskStatusSubmitted", Description: "User submits a status update of one or more tasks in Project Web App." },
    { Operation: "TimesheetAccessed", Description: "User accesses a timesheet in Project Web App." },
    { Operation: "TimesheetApproved", Description: "User approves timesheet in Project Web App." },
    { Operation: "TimesheetRejected", Description: "User rejects a timesheet in Project Web App." },
    { Operation: "TimesheetSaved", Description: "User saves a timesheet in Project Web App." },
    { Operation: "TimesheetSubmitted", Description: "User submits a status timesheet in Project Web App." },
    { Operation: "UnmanagedSyncClientBlocked", Description: "User tries to establish a sync relationship with a SharePoint or OneDrive for Business site from a computer that isn't a member of your organization's domain or is a member of a domain that hasn't been added to the list of domains (called the safe recipients list) that can access document libraries in your organization. The sync relationship is not allowed, and the user's computer is blocked from syncing, downloading, or uploading files on a document library. For information about this feature, see Use Windows PowerShell cmdlets to enable OneDrive sync for domains that are on the safe recipients list." },
    { Operation: "UpdateSSOApplication", Description: "Target application updated in Secure store service." },
    { Operation: "UserAddedToGroup", Description: "Site administrator or owner adds a person to a group on a SharePoint or OneDrive for Business site. Adding a person to a group grants the user the permissions that were assigned to the group." },
    { Operation: "UserRemovedFromGroup", Description: "Site administrator or owner removes a person from a group on a SharePoint or OneDrive for Business site. After the person is removed, they no longer are granted the permissions that were assigned to the group." },
    { Operation: "WorkflowModified", Description: "User creates, modifies, or deletes an Enterprise Project Type or Workflow phases or stages in Project Web App." }
];