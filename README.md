# Connect, Read, and Update SharePoint Online Using Python

Welcome to the SharePoint Python Connector repository! This project provides a Python-based solution for connecting to, reading, and updating files in SharePoint Online. Leveraging the Office365 Python packages, the code is organized using classes and methods, ensuring a modular and user-friendly structure.

## Key Features
- **Effortless Connection**: Seamlessly establish a secure connection to SharePoint Online.
- **File Management**: Read and update files effortlessly to streamline your data workflows.
- **Modularity**: The codebase is thoughtfully organized with classes and methods for easy customization and integration.

## Getting Started
**Creating SharePoint Client ID and Secrets**

***Important: To generate app-based credentials, you must be a site owner of the SharePoint site.***
- Go to the ```appregnew.aspx``` page in your SharePoint Online tenant. For example,       
        ```https://<Your_Company>.sharepoint.com/teams/<Teams_Name>/_layouts/15/appregnew.aspx```
  
- Navigate to this page and locate the "**Generate**" buttons positioned alongside the Client ID and Client Secret fields. Proceed to input the required information as illustrated in the accompanying screenshot.

![](https://github.com/MishraSubash/connect_sp_online/blob/main/img/sp_online_1.png)

*Title: You can give whatever title you want*

*App Domain: In case you are developing this application for your workplace or educational institution's SharePoint site, consider using your company's DNS as the App Domain.*
For more info: [SharePoint Administration](https://learn.microsoft.com/en-us/sharepoint/administration/configure-an-environment-for-apps-for-sharepoint)

- Next, proceed to grant ```site-scoped``` permissions for the recently created principal. Depending on your access types, use either link to grant permissions:
  
  ```https://<Your_Company>.sharepoint.com/teams/<Team_Name>/_layouts/15/appinv.aspx```
  or
  
  ```https://<Your_Company>-admin.sharepoint.com/_layouts/15/appinv.aspx```

  After the page loads, input your client ID into the "**App Id**" box and select "**lookup**," as indicated in the screen below:

![](https://github.com/MishraSubash/connect_sp_online/blob/main/img/sp_online_2.png)
  
- To grant permissions, you will be required to provide the permissions (if you're a site admin - "FullControl", if you're an owner -"manage") XML that outlines the necessary permissions. Copy the provided permission XML into the "**Permission Request XML**" box and proceed to select "**Create**."
 ```
 <AppPermissionRequests AllowAppOnlyPolicy="true"> 
 <AppPermissionRequest Scope="http://sharepoint/content/sitecollection" Right="Manage" />
 </AppPermissionRequests>
 ```

- Upon selecting "**Create**", a permission consent dialog will appear. Click "**Trust It**" to grant the necessary permissions.

Safeguard the created client id/secret combination as would it be your administrator account. This client ID/secret holds the capability to read/update all data within your SharePoint Online environment.

Configuration settings are now complete. Next, transition to the Terminal and install the ```office365``` library using the following command.

Use pip: 
```pip install Office365-REST-Python-Client```

Alternatively, the latest version could be directly installed via GitHub:

```pip install git+https://github.com/vgrem/Office365-REST-Python-Client.git```

## Working with Code:
Leverage the code available in this repository to interact with SharePoint Online. Code will provide basic actions such as creating directories, reading files, and updating files seamlessly. Feel free to explore, adapt, and integrate this project into your own Python projects. Whether you are a developer, data professional, or SharePoint enthusiast, this repository serves as a valuable resource for enhancing your SharePoint integration experience.

## Contributing
Welcome contributions! If you have suggestions for improvements, feel free to submit a pull request or open an issue.
