<div align="center">
    <a href="http://aptitudetech.net/">
        <img src="https://22698e.p3cdn1.secureserver.net/wp-content/uploads/2017/05/Logo_AptitudeTechnologies.png" height="128">
    </a>
    <h2>Aptitude Technologies</h2>
</div>

## Microsoft Office 365 Groups

<p>Groups in Microsoft Office 365 let you choose a set of people that you wish to collaborate with and easily set up a collection of resources for those people to share. Resources such as a shared Outlook inbox, shared calendar or a document library for collaborating on files.</p>

## Group types in Azure AD and Microsoft Graph

<p>Groups are collections of principals with shared access to resources in Microsoft services or in your app. Different principals such as users, other groups, devices, and applications can be part of groups.</p>

<p>Azure Active Directory (Azure AD) supports the following types of groups.</p>

1. Microsoft 365 groups
2. Security groups
3. Mail-enabled security groups
4. Distribution groups

<p>Note: Only the Microsoft 365 groups can be managed through our app right now.</p>

---

## Main Features

1. Create and Manage multiple Office 365 Groups
2. Manage Group Members
3. One SharePoint site per Group
4. Manage file storage
5. Supports Version 13 and Version 14

### How to Install

#### Self Hosting:

1. `bench get-app https://github.com/Aptitudetech/microsoft_integration.git`
2. `bench --site [your.site.name] install-app microsoft_integration`
3. `bench --site [your.site.name] migrate`
4. `bench restart`

---

### Bug report:

Please Create Github Issue [here](https://github.com/Aptitudetech/microsoft_integration/issues/new)

---

### Dependencies:

- [Frappe](https://github.com/frappe/frappe)
- [Erpnext](https://github.com/frappe/erpnext)

---

### Setup and Use:

#### In Frappe M365

1. Go to → M365 Groups 
    
    Enter Group Information
    
    - When creating an M365 Group, if no Group Members are specified, the Group will be created without specific permissions.
    - When a user is associated with the Group, user added as a member in the Group.  Afterward, if a user is removed from the Group, it will be also removed from the Group

<img src="/Aptitudetech/frappe-m365/raw/develop/.github/m365_groups.png" height="480">

2. Go to → M365 Groups Settings ->
    
    Setup and define following settings:

    1. App Oauth Information.
    2. Enable and Disable file sync.
        - Only new files will be synchronized.
    3. Replace File Link with SharePoint site web url.
        - If unchecked, file will reside on both sides.
        - If checked, file resides ONLY in SharePoint.
    4. Default Group for all your files.
        - Default group with which files will be synchronized.
    5. Module Settings:
        1. Define Module and Default Group for file sync.
        2. Add Role based file sync.
        
        - On this table, the app provides the flexibility to override the "Default M365 Group" and configure specific Group per Module.
        - If a Role is specified, only users part of this role will synchronize over the specified group.  Other users will default to the "Default M365 Group".

<img src="/Aptitudetech/frappe-m365/raw/develop/.github/m365_settings.png" height="480">


#### In Microsoft Azure Active Directory

1. Go to → Your Azure Portal -> Create New App Registration and use it for Connected Apps

<img src="/Aptitudetech/frappe-m365/raw/develop/.github/app_registration.png" height="480">

2. Add some extra delegated permissions in your App permissions list
    1. offline_access
    2. Files.ReadWrite.All
    3. Sites.FullControl.All
    4. User.Read.All
    5. Group.ReadWrite.All

---

### License

MIT
