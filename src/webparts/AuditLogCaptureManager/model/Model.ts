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

