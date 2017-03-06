# gsheets_mock_example

Proof of concept of a local development / push toolchain for Google Apps Scripting. Please note this is currently under active development. Please watch for updates, which are expected for the next month (as of March 6th 2017).

Requirements for development:

* node.js

Requirements for deployment:

* Production code for Google App Script lives in the cloud as a standalone script
* Your account has write access to this script
* node.js + gas-local

This repo is a proof of concept for following workflow:

* Authors/collaborators develop locally in node.js ecosystem
* Push development code via git the dev directory
* Production code is stored in Google Drive (as a standalone script container of other scripts)
* "Mock" Google App Script calls for testing

Development Setup:
Doing the following, you can download the source used for the Google Apps Script, run it on the node.js environment, and contribute code back into the repo:

* Clone/fork this repo
* Install node.js
* npm install
* npm test
* Write code and test, if applicable

Deployment Setup:
If your account can be used to push the changes to the cloud (to the actual standalone script)

* Create and save a (new standalone script)['https://script.google.com'] with empty contents. Keep the package ID for below
* Authenticate with [gapps](https://www.npmjs.com/package/node-google-apps-script) with the user that created the script document above. 
* In the same directory as the cloned repo: `gapps init "your script project id from above"` If you get script ID error, it may be due to lack of permissions on the script (which happened to me while I was juggling between two accounts). If you need more than one attempt at `gapps init`, you can safely `rm -r src` directory and `rm gas.package.json` file.
* Copy the contents of the /dev folder into /src folder (or manually delete /src contents and use a symlink)
* `gapps upload`

Workflow summary:

* Contribute code to the /dev folder, which is version controlled with git
* Easy way to deploy; for single devs (with combo dev and deployment setup) or for a team (where the dev and deployment setups are differentiated)
* Other Open Source contributors could also do pull requests, all from within node.js!

