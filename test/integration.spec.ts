import { msalConfig } from '../test/config';
import { DlMSGraphClient } from '../src/app/msgraph';
import { expect } from 'chai';

const tests: any[] = [
    {
        name: 'my-profile',
        url: 'https://graph.microsoft.com/v1.0/me',
        method: 'GET',
        expect: 'mail'
    },
    {
        name: 'get-sites',
        url: 'https://graph.microsoft.com/v1.0/sites/root',
        method: 'GET',
        expect: 'webUrl'
    }
];


function main(tests: Array<any>) {

    tests.forEach(test => {
        describe(`sharepoint-graph: test = ${test.name}`, () => {
            it(`Performing Graph request to ${test.url}`, function (done):void {
                this.timeout(30 * 1000);
                const graph = new DlMSGraphClient(msalConfig);
                const url = test.url;

                if (test.method === 'GET') {
                    graph.get(url)
                    .then((result:any) => {
                        const checkResult = result[test.expect]
                        expect(checkResult).to.not.null;
                        done();
                    })
                    .catch((error) => {
                        done(error);
                    })
                }
    
                
            })
        })
    })

}
main(tests);