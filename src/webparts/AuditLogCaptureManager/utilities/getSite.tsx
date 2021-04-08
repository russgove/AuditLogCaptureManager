import { Site } from "@pnp/sp/sites";

import "@pnp/sp/presets/all";
import "@pnp/sp/sites";

export async function getSite(siteUrl: string): Promise<any> {
    debugger;
    var url: string = decodeURIComponent(siteUrl);
    var site = Site(url);
    return site.get();

}