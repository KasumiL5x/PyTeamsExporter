# PyTeamsExporter
> A web application for exporting Microsoft Teams chats.

This is a simple web application for retrieving all personal Microsoft Teams chats and exporting them for backup. Microsoft do not support this natively, but it is possible to use [Microsoft Graph](https://docs.microsoft.com/en-us/graph) to achieve this.

Tech stack with Python `3.9.7`:

* `flask` for creating and serving a RESTful API and frontend.
* `requests_oauthlib` for talking to the Microsoft Graph API.

## Usage
### Setting Your Azure Application Keys (self-hosting/developers)
To use the Microsoft Graph API, you need to have a registered Azure application. You can find instructions [here](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app) to get started.

Your application will need to have the following configuration:

* Add a `Web` platform and set the redirect URIs to `http://localhost:5000` and `http://localhost:5000/login/authorized`.
* Set the `Supported account types` to `Accounts in any organizational directory (Any Azure AD directory - Multitenant`).
* Create and note down a new secret key.

In `app.py`, set `CLIENT_ID` and `CLIENT_SECRET` to your Azure Application's own `Application (client) ID` and secret key respectively.

### Deploying Flask (self-hosting or developers only)
If you want to deploy this live, please look at [this page](https://flask.palletsprojects.com/en/1.1.x/deploying/#deployment) for help. You may need to make some tweaks to `app.py` as it is currently setup for a local development server only.

### Using the Web App (all users)
If you are running this locally, navigate a terminal to the root directory (where `app.py` lives) and run `python app.py`.

In your browser, navigate to `http://localhost:5000` (or the hosting URL you have been given). Press `Sign in` and you will be redirected to Microsoft's login portal. Login with the same account you use on Microsoft Teams.

Once signed in, press the `Get Chat History` button which will populate a table with all of your chats. Use the checkboxes on the left of each row to select which chats you would like to export.

Once desired chats are selected, press the `Download Selected Chats` button and you will be prompted to download all of the selected chats once they finish processing.

There are two toggle buttons labeled `Chat + Attachments` and `Chat Only`. If the former is active, both the chat itself and all attachments will be processed. This is the recommended option. Note that SharePoint files are currently **not** downloaded due to permission issues, but rather are inserted back into the chat as hyperlinks. If the latter is active, then only the chat is downloaded and attachments aren't processed at all, which means they will simply be missing from the chat history, even in text.

When a chat is ready to download, you will receive a `zip` file made up of the following:

* `attachments` - A folder with all attachments for the chat.
* `dist` - A folder for local CSS and JS used by the HTML file below.
* `attachments.json` - A complete record of attachments in JSON format.
* `chat.html` - A complete record of the chat history in a pretty HTML format. **You likely want this.**
* `chat.json` - A complete record of the chat history in JSON format.
* `metadata.json` - Basic information about the chat.

If you are dealing with large chats, the download process may take quite a while. This is because retrieving the data requires many calls to Microsft's Graph API which seems to have a built-in speed limiter for frequent requests. If you are running the app locally, then you can see the rough progress in the terminal window, which is quite handy, particularly for large files.

### Notes / Bugs / Limitations
* There is code present to 'properly' download hosted contents (e.g., pasted images) using the API, but it gets throttled too frequently to be practical. There is alternative code that manually scrapes image tags to achieve _almost_ the same thing. Feel free to alter this behavior in `app.py` by changing `get_hosted_contents_through_api`.
* Only HTML and text messages are currently supported. At the time of writing, events (call started, call ended, etc.) don't seem to be functional in the API.
* Attachments that use SharePoint URLs are protected with a different login token and will instead be inserted as a link. It may be possible to derive a drive link and use [this](https://docs.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0&tabs=http) call to read it.
* I'm sure I'm missing plenty of other rich information that the API lets me gather. This is a quick and dirty project!
