require("dotenv").config();
const sharepoint = require("sharepointplus/dist");
const { JSDOM } = require('jsdom');

// Create a fake window object
const dom = new JSDOM('<!DOCTYPE html><html><body></body></html>');
global.window = dom.window;
global.document = dom.window.document;

// define the SharePoint credentials
const credentials = {
  username: process.env.SHAREPOINT_USERNAME,
  password: process.env.SHAREPOINT_PASSWORD,
};

// create a new SharePointPlus instance with the credentials
const sp = sharepoint(process.env.SHAREPOINT_SITE_URL).auth(credentials);

// retrieve all items from the list
sp.list(process.env.SHAREPOINT_LIST_NAME).get((data) => {
  console.log(data);
}).catch((error) => {
	console.log(error);
});