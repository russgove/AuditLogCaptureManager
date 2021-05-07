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

export function getContentType(parentContext: IAuditLogCaptureManagerProps, siteUrl: string): Promise<IContentTypeInfo> {
  debugger;
  if (!siteUrl) {
    console.log(`site url passed to getContentType is empty`);
    return Promise.reject(`site url passed to getContentTypew is empty`);
  }
  try {
    var url: string = decodeURIComponent(siteUrl);
    console.log(`site url passed to getContentType is ${url}`);
    var rootweb = Web(url);
    debugger;
    const contentType = rootweb.contentTypes.getById(parentContext.auditItemContentTypeId)().then((ctLookupResults) => {
      debugger;
      if (ctLookupResults['odata.null']) {
        console.log(`ContentType not found`);
        return Promise.reject(`ContentType not found`)
      }
      else {
        return ctLookupResults;
      }
    });
    debugger;
    return contentType;
    debugger;

  }
  catch (ee) {
    console.log(ee);
    debugger;
  }
}