# **MGT-APP**

This project demonstrates how to use Microsoft Graph Toolkit and MSAL.js to display emails and chat information. The application fetches emails and chat data from the Microsoft Graph API and displays them on a webpage.

## **Features**

- Log in to a Microsoft account
- Display the latest emails
- Display chat information, including members and the latest messages

## **Installation and Usage**

### Prerequisites

- Node.js and npm installed
- A registered application in Microsoft Azure to obtain a `clientId`

### Clone the Repository

```sh
git clone https://github.com/yjchenbe/MGT-APP.git
cd MGT-APP
```
### **Configuration**

Register an application in the Azure portal and obtain the clientId. Insert it into the index.html file's clientId attribute.

```html
<mgt-msal2-provider client-id="YOUR_CLIENT_ID_HERE"></mgt-msal2-provider>
```

## **Run the Application**
Since this project consists of static HTML files, you can use any local HTTP server to run it. Here are a few simple methods:

### Using Live Server (VS Code Extension)


1.   Open the project in Visual Studio Code.
2.   Install and start the Live Server extension.

## **Usage**
### Log In

*   Click the "Log In" button on the page to log in with your Microsoft account.


### View Emails

*   After logging in, you will see the latest 5 emails, including the subject, body preview, and sender.

### View Chat Information

*   You will see all non-meeting type chats, including the member list and the latest messages.


## **Technical Details**
### Key Technologies


*   HTML5
*   JavaScript
*   Microsoft Graph Toolkit
*   MSAL.js
*   清單項目

### API Endpoints


*   /me/messages - Fetch emails
*   /me/chats - Fetch chat information

## **Contributing**
Any form of contribution is welcome! Please send a Pull Request or submit an issue report.

## **License**
This project is licensed under the MIT License.

## **Reference**

[Get started with Microsoft Graph Toolkit ](https://learn.microsoft.com/en-us/training/modules/msgraph-toolkit-intro/1-introduction)

