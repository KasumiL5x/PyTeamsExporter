# PyTeamsExporter
> A web application for exporting personal Microsoft Teams chats.

This is a simple web application for retrieving all personal Microsoft Teams chats and exporting them for backup. Microsoft do not support this natively, but it is possible to use [Microsoft Graph](https://docs.microsoft.com/en-us/graph) to achieve this.

Tech stack with Python `3.9.7`:

* `flask` for creating and serving a RESTful API and frontend.
* `requests_oauthlib` for talking to the Microsoft Graph API.

## Usage
### Starting Flask
If you want to deploy this live, please look at [this page](https://flask.palletsprojects.com/en/1.1.x/deploying/#deployment) for help.

To deploy this locally, navigate your terminal to the root directory (where `app.py` lives) and run either `python app.py`.

### Using the Web App
In your browser, navigate to `http://localhost:5000`. Press `Sign in` and you will be redirected to Microsoft's login portal. Login with the same account you use on Microsoft Teams.

Once signed in, press the `Get Chat History` button which will populate a table with all of your chats. Use the checkboxes on the left of each row to select which chats you would like to export.

Once desired chats are selected, press the `Download Selected Chats` button and you will be prompted to download all of the chats in a pretty HTML format.

### Limitations
* Only HTML and text messages are currently supported. System messages will not be included.
* Attachments are not considered. This includes images, which are stored on Microsoft's servers. I am considering a separate pass to download, compress, and send the images too, but only if there's a need for it and I have time to do so.
* I'm sure I'm missing plenty of other rich information that the API lets me gather. This is a quick and dirty project.
