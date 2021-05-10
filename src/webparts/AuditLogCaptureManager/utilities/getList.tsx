import { Web } from "@pnp/sp/webs";

import "@pnp/sp/presets/all";
import "@pnp/sp/sites";

export function getList(siteUrl: string, listId: string): Promise<any> {
    var url: string = decodeURIComponent(siteUrl);
    return Web(url).lists.getById(listId).get();
}