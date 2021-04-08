import * as inquirer from 'inquirer';
import { v4 as uuidv4 } from 'uuid';
import { readFile, statSync } from 'fs';
import { basename } from 'path';

import { DLMSGraphClient, DLMSALConfig } from '@driverlane/sharepoint-msgraph-wrapper';

// Project Imports 
// import { DlMSGraphClient } from './app/msgraph';
import { msalConfig } from '../test/config';

// Demo Constants
const sharePointHost = 'piggles.sharepoint.com';
const sharePointSiteAddress = '/sites/PigglesSharePoint';
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
        const graph = new DLMSGraphClient(msalConfig);
        const sharePointUrl = `/sites/${sharePointHost}:${sharePointSiteAddress}`;
        const getSiteId: any = await graph.request('GET', sharePointUrl);
        console.log(getSiteId);

    } catch (error) {
        console.error('sharePointInfo: An error occured');
        console.error(error);
    }

}

async function sharePointList() {
    try {
        const graph = new DLMSGraphClient(msalConfig);
        const sharePointUrl = `/sites/${sharePointHost}:${sharePointSiteAddress}`;
        const getSiteId: any = await graph.request('GET', sharePointUrl);

        // Get List Information
        const getListsUrl = `/sites/${getSiteId.id}/lists`;
        const getListsFromSite = await graph.request('GET', getListsUrl);
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
        const graph = new DLMSGraphClient(msalConfig);
        graph.sp_setSharePointSettings(sharePointHost, sharePointSiteAddress, '');
        const listItem = {
            fields: {
                Title: 'SharePoint using Graph',
                Value: uuidv4()
            }
        };

        const addListItem = await graph.sp_addItemToList('listName', listItem);
        console.log(addListItem);


    } catch (error) {
        console.error('addItemToList: An error occured');
        console.error(error);
    }
}

async function createFolder() {
    try {
        const graph = new DLMSGraphClient(msalConfig);
        graph.sp_setSharePointSettings(sharePointHost, sharePointSiteAddress, 'Documents');
        const newfolder = await graph.sp_createFolder('', uuidv4(), 'fail');
        console.log(newfolder);
    } catch (error) {
        console.error('createFolder: An error occured');
        console.error(error);
    }
}

async function createDocument() {
    try {
        // Graph Client
        const graph = new DLMSGraphClient(msalConfig);
        graph.sp_setSharePointSettings(sharePointHost, sharePointSiteAddress, 'Documents');

        // File Information
        const fileName = basename(uploadFile);
        const fSize = statSync(uploadFile);

        if ((fSize.size / (1024 * 1024)) < 4.096) {
            console.log('Small File Upload');
            readFile(uploadFile, 'utf8', async (err, data) => {
                if (err) {
                    console.error('Error Reading File');
                    console.error(err);

                } else {
                    const smallFileUpload = await graph.sp_createSmallFile('', fileName, data);
                    console.log(smallFileUpload);
                }
            });
        } else {
            console.log('Large File Upload');
            const largeFileUploadSessionUrl = await graph.sp_generateSessionUrl('', fileName);
            const largeFile = graph.helper_pathToFile(uploadFile, fileName);
            const uploadLargeFile = graph.sp_addLargeFile(largeFileUploadSessionUrl, fileName, fSize.size, 'fail', largeFile);
            console.log(uploadLargeFile)
        }
    } catch (error) {
        console.error('createDocument: An error occured');
        console.error(error);
    }
}

