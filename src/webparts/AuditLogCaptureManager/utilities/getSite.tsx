import { Site } from "@pnp/sp/sites";

import "@pnp/sp/presets/all";
import "@pnp/sp/sites";

export function getSite(siteUrl: string): Promise<any> {
    debugger;
    if (!siteUrl) {
        console.log(`site url passed to getSite is empty`);
        return Promise.reject(`site url passed to getSite is empty`);
    }
    var url: string = decodeURIComponent(siteUrl);
    var site = Site(url);
    return site.get();

}