# POC Excel Add-In

Simple example showing how to create an office addon for windows, mac and web.

Generated via `yo office` and enabled autoloading via:

https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime

## How to run

for a web hosted sheet you can start and load the add-in with just one command:
`npm run start -- web --document https://1drv.ms/....document...`

If you want to open excel with the addin loaded:

`npm run start -- desktop --document path`

## Observations


- It appears that there are no content security policies preventing once javascript
- However it is required that services that you access run over https.
- Certificates for https://localhost:3000 need to be trusted, on windows it appears to just work. (or run `google-chrome --ignore-certificate-errors`)
- the addin api is not able to control the host's window  or open/close documents


## To Verify: How are addons installed?

Whilst `npm run start` automatically installs and runs office with the add-in loaded it would be good to know how to install the addon on a permanent basis and without nodejs.


**On Windows**

`%LOCALAPPDATA%\Microsoft\Office\Addins` ?

https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins#specify-the-shared-folder-as-a-trusted-catalog

**On MacOS**

The manifest is copied to ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/.

source: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac#sideload-an-add-in-in-office-on-mac

**Office on the Web**

The manifest is hosted by the dev server.

When you access the web-based Office app, it automatically(???) uses the manifest URL to load your add-in.
