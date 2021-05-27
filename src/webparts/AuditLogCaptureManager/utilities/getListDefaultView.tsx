import { IList, IViewInfo } from "@pnp/sp/presets/all";
import { Web } from "@pnp/sp/webs";

import "@pnp/sp/presets/all";
import "@pnp/sp/sites";

export function getListDefaultView(siteUrl: string, listId: string): Promise<IViewInfo> {
    var url: string = decodeURIComponent(siteUrl);
    return Web(url).lists.getById(listId).defaultView.get();
}