# Using SharePoint with MS graph

A simple project which demonstrates how to create a Node.Js console app that connects the MSGraph and accesses SharePoint. 


## Medium Blog post
https://medium.com/@mbarben/sharepoint-on-the-graph-e813e838f604


## To Setup 

 * You will need to setup Azure AD Application first, that will give you the `clientId` `clientSecret` `tenentid` values
 * Copy `/test/test.config.ts` to `/test/config.ts` and update the corresponding values
 * Clone the report `git clone https://github.com/anvilation/sharepoint-graph` 
 * browse to the folder and install modules `npm install`
 * Update the `/src/index.ts` file and update the following constants:
    - `sharePointHost` the host name for your tenant
    - `sharePointSiteAddress` the relative location of your SharePoint Site
    - `listName` The name of a list that you want to query
    - `uploadFile` the location of a file that you want to upload with
 * To run run `npm run dev`
 * There are is also an additional integrationt tests by running `npm run test`
