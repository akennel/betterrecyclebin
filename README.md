## SharePoint Better Recycle Bin

A SharePoint Online WebPart that displays the Second-Stage (Site Collection) Recycle Bin in a sortable and filterable table.

The default Recycle Bin does not make it easy to find deleted files. This web part allows your search for files by name, path, deleted date, deleted by, or a combination of all of the above.

Tested on a production system, but use at your own risk. Tested with 25k+ items in the Recycle Bin.

Uses PNPJS and the SharePoint REST API.
The WebPart requires that the user be a member of the Site Collection Admins group. Other users will be shown an Access Denied message. No other configuration required. 

Add the WebPart to your App Catalog (either download the code and build, or use the link below to download a pre-built package). Once the WebPart is in your App Catalog, add it to a page on your Site Collection. 

You can download a packaged version [here](https://github.com/akennel/betterrecyclebin/blob/master/better-recycle-bin.sppkg?raw=true)

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```


