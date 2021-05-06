import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { sp } from "@pnp/sp";
import { ContentTypes, IContentType, IContentTypeAddResult, IContentTypeInfo, IContentTypes } from "@pnp/sp/content-types";
import { Fields, FieldTypes } from "@pnp/sp/fields";
import { ILists, Lists } from "@pnp/sp/lists";
import { IContextInfo, ISite, Site } from "@pnp/sp/sites";
import { IWebs, Web, Webs } from "@pnp/sp/webs";
import { find } from 'lodash';

import { IAuditLogCaptureManagerProps } from '../components/IAuditLogCaptureManagerProps';

import "@pnp/sp/presets/all";
import "@pnp/sp/sites";

export async function getContentType(parentContext: IAuditLogCaptureManagerProps, siteUrl: string): Promise<IContentTypeInfo> {
  debugger;

  try {
    var url: string = decodeURIComponent(siteUrl);
    var rootweb = Web(url);
    debugger;
    const contentType = await rootweb.contentTypes.getById(parentContext.auditItemContentTypeId)();
    return contentType;
    debugger;

  }
  catch (ee) {
    console.log(ee);
    debugger;
  }
}