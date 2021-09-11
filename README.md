# PyTeamsExporter
> A web application for exporting personal Microsoft Teams chats.

This is a simple web application for retrieving all personal Microsoft Teams chats and exporting them for backup. Microsoft do not support this natively, but it is possible to use [Microsoft Graph](https://docs.microsoft.com/en-us/graph) to achieve this.

Tech stack with Python `3.9.7`:

* `flask` for creating and serving a RESTful API and frontend.
* `requests_oauthlib` for talking to the Microsoft Graph API.

## Usage
### Registering an Azure Application
To use the Microsoft Graph API, you need to have a registered Azure application. You can find instructions [here](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app) to get started.

Once registered, you will need to make some small tweaks for a web application. Under the `Authentication` tab, add a new platform:

* Add a new `Web` platform and set the redirect URIs to `http://localhost:5000` and `http://localhost:5000/login/authorized`.
* Set the `Supported account types` to `Accounts in any organizational directory (Any Azure AD directory - Multitenant`).

Go to the `Overview` tab and note down the `Application (client) ID` value.

Next, under the `Certificates & secrets` tab, create a new client secret and **immediately note the key's value field down somewhere**.

In `app.py`, set `CLIENT_ID` and `CLIENT_SECRET` to the two above values above, respectively.

**Do not share these values with anybody else. They are private.**

### Starting Flask
If you want to deploy this live, please look at [this page](https://flask.palletsprojects.com/en/1.1.x/deploying/#deployment) for help.

To deploy this locally, navigate your terminal to the root directory (where `app.py` lives) and run either `python app.py`.

### Using the Web App
In your browser, navigate to `http://localhost:5000`. Press `Sign in` and you will be redirected to Microsoft's login portal. Login with the same account you use on Microsoft Teams.

Once signed in, press the `Get Chat History` button which will populate a table with all of your chats. Use the checkboxes on the left of each row to select which chats you would like to export.

Once desired chats are selected, press the `Download Selected Chats` button and you will be prompted to download all of the chats in a pretty HTML format.

### Bugs
For some reason, chats seem to be limited in count, as do messages. I'm not sure why this is yet, but it's something I'm investigating.

### Limitations
* Only HTML and text messages are currently supported. System messages will not be included.
* Attachments are not considered. This includes images, which are stored on Microsoft's servers. I am considering a separate pass to download, compress, and send the images too, but only if there's a need for it and I have time to do so.
* I'm sure I'm missing plenty of other rich information that the API lets me gather. This is a quick and dirty project.
