import { ILists } from "@pnp/sp/lists";
import { Site } from "@pnp/sp/sites";
import { IWebs, Web, Webs } from "@pnp/sp/webs";

import "@pnp/sp/presets/all";
import "@pnp/sp/sites";

export function getLists(siteUrl: string): Promise<ILists> {
    if (!siteUrl) {
        console.log(`site url passed to getLists is empty`);
        return Promise.reject(`site url passed to getLists is empty`);
    }
    var url: string = decodeURIComponent(siteUrl);
    return Web(url).lists.select("Title,Id").get();
    //   return site.rootWeb.lists.getById(listId).get();
}