import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
export function fetchAZFunc(client: AadHttpClient, url: string, method: string, body?: string): Promise<any> {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');
    return client.fetch(url,
        AadHttpClient.configurations.v1,
        {
            method: method,
            headers: requestHeaders,
            body: body
        })
        .then(async (response: HttpClientResponse) => {
            debugger;
            if (response.ok) {
                return response.json().then((results) => {
                    return results;
                });
            }
            else {
                throw (response.statusText);
            }
        });
    //let caller handle it to show nice message
    // .catch((err) => {
    //     alert(err);
    //     console.log(err);
    //     return null;
    // });
}


