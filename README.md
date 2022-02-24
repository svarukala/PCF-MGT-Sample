# PCF-MGT-Sample
Shows how to use Microsoft Graph Toolkit components in Power Apps Component Framework (PCF) custom controls. This sample is shows how to authenticate using MGT MSAL provider and then shows how to use Agend and FileList components. 
The FileList component also supports Upload of files to SharePoint Onlines. Based on my testing, it can support large files upload too. This will help get around the upload limit of 100MB in Power Apps and Power Automate.

#Demo
Here is the sample demo. Once you add the component to your app, you need to set the Client Id, Scopes and Redirect Url.
Client Id - Azure Ad App Id
Scopes - Comma delimited scopes
Redirecr Url - The Power App published URL

![Demo](./PCFwithMGT/mgt-in-pcf.gif)