import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

export function fetchAZFunc(client: AadHttpClient, url: string): Promise<any> {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');
    return client.fetch(url,
        AadHttpClient.configurations.v1,
        {
            method: "GET",
            headers: requestHeaders,
        })
        .then(async (response: HttpClientResponse) => {
            debugger;
            //  setCaptures(response)
            return response.json().then((results) => {
                debugger;
                return results
            });

        }).catch((err) => {
            alert(err);
            console.log(err);
            return null;
        });
}


