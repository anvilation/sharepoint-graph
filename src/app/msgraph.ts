import "isomorphic-fetch";
import * as msal from "@azure/msal-node";
import { AuthenticationProvider, Client, ClientOptions, LargeFileUploadTask, FileObject } from "@microsoft/microsoft-graph-client";
import { statSync, readFileSync } from 'fs';

export class DLAuthenticationProvider implements AuthenticationProvider {
    config: MsalConfig = {
        auth: {
            clientId: '',
            authority: '',
            clientSecret: '',
            graphendpoint: '',
        }
    };
    tokenrequest: any;
    cca: msal.ConfidentialClientApplication;

    constructor(config: MsalConfig) {
        this.config = config;
        this.tokenrequest = {
            scopes: [config.auth.graphendpoint + '.default'],
        };
        this.cca = new msal.ConfidentialClientApplication(this.config);
    }

    public async getAccessToken(): Promise<string> {
        return new Promise((ok, fail) => {
            this.cca.acquireTokenByClientCredential(this.tokenrequest)
                .then((response) => {
                    ok(response.accessToken);
                })
                .catch((error) => {
                    fail(error);
                })
        })
    }

}



// TODO: DlMSGraphClient
//  * Add Query Options
//  * Add Search Options

export class DlMSGraphClient {
    private config: MsalConfig = {
        auth: {
            clientId: '',
            authority: '',
            clientSecret: '',
            graphendpoint: '',
        }
    };;
    client: Client;

    constructor(config: MsalConfig) {
        this.config = config;
        const clientOptions: ClientOptions = {
            defaultVersion: 'v1.0',
            debugLogging: true,
            authProvider: new DLAuthenticationProvider(this.config)
        };
        this.client = Client.initWithMiddleware(clientOptions);
    }


    async get(url: string, stream = false) {
        try {
            if (stream) {
                const response = await this.client.api(url).getStream();
                return response;
            } else {
                const response = await this.client.api(url).get();
                return response;
            }
        } catch (error) {
            throw new Error(error);
        }
    }

    async post(url: string, body: any) {
        try {
            const response = await this.client.api(url).post(body);
            return response;
        } catch (error) {
            throw new Error(error);
        }
    }

    async put(url: string, body: any, stream = false) {
        try {
            if (stream) {
                const response = await this.client.api(url).putStream(body);
                return response;
            } else {
                const response = await this.client.api(url).put(body);
                return response
            }
        } catch (error) {
            throw new Error(error);
        }
    }

    async addLargeFile(url: string, filename: string, filesize: number, file: File) {
        try {
            const payload = {
                item: {
                    "@microsoft.graph.conflictBehavior": "fail",
                    name: filename,
                },
            };
            const fileObject: FileObject = {
                size: filesize,
                content: file,
                name: filename
            };

            const uploadSession = await LargeFileUploadTask.createUploadSession(this.client, url, payload);
            const uploadTask = await new LargeFileUploadTask(this.client, fileObject, uploadSession);
            const response = await uploadTask.upload();
            return response;

        } catch (error) {
            throw new Error(error);
        }
    }

    pathToFile(filepath: string, filename: string): File {
        // Create Buffer
        const buff = readFileSync(filepath);

        // Get File Stats
        const stats = statSync(filepath);

        // Convert Buffer to Blob
        let file: any = JSON.stringify({ blob: buff.toString("base64") })
        file = JSON.parse(file);

        let blob: any = Buffer.from((<any>file).blob, "base64");
        blob.lastModifiedDate = stats.mtime;
        blob.name = filename;
        return <File>blob;
    }
}

export interface MsalConfig {
    auth: {
        clientId: string;
        authority: string;
        clientSecret: string;
        graphendpoint: string;
    }
}
