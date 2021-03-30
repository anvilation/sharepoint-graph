import { MsalConfig } from "../src/app/msgraph";


export const msalConfig: MsalConfig = {
    auth: {
        clientId: 'clientid',
        authority: 'https://login.microsoftonline.com/{tenentid}',
        clientSecret: 'clientSecret',
        graphendpoint: 'https://graph.microsoft.com/'
    }
}