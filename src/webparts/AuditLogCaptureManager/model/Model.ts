export class Subscription {
    public contentType: string;
    public status: string;
    public webhook: Webhook;
}
export class Webhook {
    public address: string;
    public authId: string;
    public expriration: string;
    public status: string;
}
