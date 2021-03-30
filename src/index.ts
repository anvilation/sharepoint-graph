import * as inquirer from 'inquirer';
import { v4 as uuidv4 } from 'uuid';
import { readFile, statSync } from 'fs';
import { basename } from 'path';

// Project Imports 
import { DlMSGraphClient } from './app/msgraph';
import { msalConfig } from './environments/piggles';

// Demo Constants
const sharePointHost = 'sample.sharepoint.com';
const sharePointSiteAddress = '/sites/SampleSharePoint';
const listName = 'SimpleList';
const uploadFile = './temp/SmallDocument.docx';

function main() {
    const prompt = [
        {
            type: 'list',
            name: 'action',
            message: 'What action to perform',
            choices: [
                'Access SharePoint Information',
                'Access SharePoint List',
                'Add item to SharePoint List',
                'Create Folder',
                'Upload Document'
            ]
        }
    ];

    inquirer.prompt(prompt)
        .then((answers: any) => {
            switch (answers.action) {
                case 'Access SharePoint Information':
                    sharePointInfo();
                    break;
                case 'Access SharePoint List':
                    sharePointList();
                    break;
                case 'Add item to SharePoint List':
                    addItemToList();
                    break;
                case 'Create Folder':
                    createFolder();
                    break;
                case 'Upload Document':
                    createDocument();
                    break;
                default:
                    break;
            }
        })
}


main();


// SharePoint Operations
async function sharePointInfo() {
    try {
        const graph = new DlMSGraphClient(msalConfig);
        const sharePointUrl = `/sites/${sharePointHost}:${sharePointSiteAddress}`;
        const getSiteId: any = await graph.get(sharePointUrl);
        console.log(getSiteId);

    } catch (error) {
        console.error('sharePointInfo: An error occured');
        console.error(error);
    }

}

async function sharePointList() {
    try {
        const graph = new DlMSGraphClient(msalConfig);
        const sharePointUrl = `/sites/${sharePointHost}:${sharePointSiteAddress}`;
        const getSiteId: any = await graph.get(sharePointUrl);

        // Get List Information
        const getListsUrl = `/sites/${getSiteId.id}/lists`;
        const getListsFromSite = await graph.get(getListsUrl);
        const lists = (<any>getListsFromSite).value;

        for (let l = 0; l < lists.length; l++) {
            const list = lists[l];
            if (list.displayName === listName) {
                console.log(list);
            }
        }

    } catch (error) {
        console.error('sharePointList: An error occured');
        console.error(error);
    }
}

async function addItemToList() {
    try {
        const graph = new DlMSGraphClient(msalConfig);
        const sharePointUrl = `/sites/${sharePointHost}:${sharePointSiteAddress}`;
        const getSiteId: any = await graph.get(sharePointUrl);

        // Get List Information
        const getListsUrl = `/sites/${getSiteId.id}/lists`;
        const getListsFromSite = await graph.get(getListsUrl);
        const lists = (<any>getListsFromSite).value;

        // Get List Id
        let listId;
        for (let l = 0; l < lists.length; l++) {
            const list = lists[l];
            if (list.displayName === listName) {
                listId = list.id
            }
        }


        // Create Body
        const listItem = {
            fields: {
                Title: 'SharePoint using Graph',
                Value: uuidv4()
            }
        };
        const addListItemURL = `${getListsUrl}/${listId}/items`;
        const addListItem = await graph.post(addListItemURL, listItem);
        console.log(addListItem);

    } catch (error) {
        console.error('addItemToList: An error occured');
        console.error(error);
    }
}

async function createFolder() {
    try {
        const graph = new DlMSGraphClient(msalConfig);
        const sharePointUrl = `/sites/${sharePointHost}:${sharePointSiteAddress}`;
        const getSiteId: any = await graph.get(sharePointUrl);

        // Create Folder
        const libraryUrl = `/sites/${getSiteId.id}/drive/root/children`
        const foldername = {
            "name": uuidv4(),
            "folder": {}
        };

        const newFolder = await graph.post(libraryUrl, foldername);
        console.log(newFolder);


    } catch (error) {
        console.error('createFolder: An error occured');
        console.error(error);
    }
}

async function createDocument() {
    try {
        // Graph Client
        const graph = new DlMSGraphClient(msalConfig);

        // File Information
        const fileName = basename(uploadFile);
        const fSize = statSync(uploadFile);

        // Get SharePoint Information
        const sharePointUrl = `/sites/${sharePointHost}:${sharePointSiteAddress}`;
        const getSiteId: any = await graph.get(sharePointUrl);

        // Create Folder
        const getRootIdUrl = `/sites/${getSiteId.id}/drive/root/`
        const getRootId: any = await graph.get(getRootIdUrl);

        if ((fSize.size / (1024 * 1024)) < 4.096) {
            console.log('Small File Upload');

            // Use Upload Small File Method
            const smallUploadUrl = `/sites/${getSiteId.id}/drive/items/${getRootId.id}:/${fileName}:/content`

            readFile(uploadFile, 'utf8', async (err, data) => {
                if (err) {
                    console.error('Error Reading File');
                    console.error(err);

                } else {
                    const smallFileUpload = await graph.put(smallUploadUrl, data);
                    console.log(smallFileUpload);
                }
            });
        } else {
            console.log('Large File Upload');
            // Use Upload Large File Method
            // based upon https://docs.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0
            // Create a upload Session
            const largeFileUploadSessionUrl = `/sites/${getSiteId.id}/drive/items/${getRootId.id}:/${fileName}:/createUploadSession`

            const largeFile = graph.pathToFile(uploadFile, fileName);
            const largeFileUpload = await graph.addLargeFile(largeFileUploadSessionUrl, fileName, fSize.size, largeFile);
            console.log(largeFileUpload);
        }
    } catch (error) {
        console.error('createDocument: An error occured');
        console.error(error);
    }
}

