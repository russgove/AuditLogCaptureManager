import { Site } from "@pnp/sp/sites";

import "@pnp/sp/presets/all";
import "@pnp/sp/sites";

export async function getList(siteUrl: string, listId: string): Promise<any> {
    debugger;
    var url: string = decodeURIComponent(siteUrl);
    return await Site(url).rootWeb.lists.getById(listId).get();
    //   return site.rootWeb.lists.getById(listId).get();
}