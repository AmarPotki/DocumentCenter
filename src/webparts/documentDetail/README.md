# Document Detail Web Part

## Summary

SharePoint Framework client-side web part built using React for displaying the information related to a document or documents inside a document set.

## Used SharePoint Framework Version
1.3.1

## Solution

Solution Name
Document Center

Solution|Author(s)
Ramin Ahmadi,Ben Weeks

## Version history

Version|Date|Comments
-------|----|--------
1.0.0| 11/28/2017 |Initial release

## Prerequisites
* Sharepoint Framework 1.3.1
* TypeScript 2.4.2 or later
* React
* sp-pnp-js 2.0.8
* react-taxonomypicker 0.0.35
* gulp-spsync-creds 2.3.6
* node-sppkg-deploy 1.1.1

### How-to

  - in your web browser navigate to the page you want to add this web part
  - add the Document Detail web part from Cielo Costa group
  - in the configuration specify the **Search result page**,also optionally the **No document found message** and **Show description?** and **Select desire Properties** and **Email body message**.
  - confirm the changes by clicking the **Apply** button


## Features

This project contains client-side web parts built on the SharePoint Framework using React illustrating.

This sample illustrates the following concepts on top of the SharePoint Framework:

- using React for building SharePoint Framework client-side web parts
- using React spread operator for passing multiple properties to React components
- conditionally rendering React components
- managing state in a parent component
- styling React applications using Office UI Fabric
- chaining multiple ES6 promises
- reading SharePoint list items using the pnp
- using MockService for displaying data on local machine
- using TypeScript features like interfaces,classes,generics,etc

#### Local Mode
A browser in local mode (localhost) will be opened.
https://localhost:4321/temp/workbench.html

#### SharePoint Mode
If you want to try on a real environment, open:
https://your-domain.sharepoint.com/_layouts/15/workbench.aspx