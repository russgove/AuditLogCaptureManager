import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

export function callManagementApi(client: AadHttpClient, url: string, method: string, body?: string): Promise<any> {
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

            if (response.ok) {
                return response.json()
                    .then((results) => {
                        return results;
                    })
                    .catch((e) => {
                        console.log(`the call to ${url} returned invalid json`);
                        console.log(`${response.text}`);
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


