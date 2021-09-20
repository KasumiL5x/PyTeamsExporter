# Fill these in from your Azure app (see https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app).
CLIENT_ID = 'YOUR_CLIENT_ID'
CLIENT_SECRET = 'YOUR_CLIENT_SECRET'

# App redirect URI and allowed scopes.
REDIRECT_URI = 'http://localhost:5000/login/authorized'
SCOPES = [
  "User.Read",
  "Chat.Read",
  "Files.Read",
  "offline_access"
]

# URLs and endpoints for authorization.
AUTHORITY_URL = 'https://login.microsoftonline.com/common'
AUTH_ENDPOINT = '/oauth2/v2.0/authorize'
TOKEN_ENDPOINT = '/oauth2/v2.0/token'

# Microsoft Graph API configuration.
RESOURCE = 'https://graph.microsoft.com/'
# NOTE: the 'beta' channel is required for the API calls we need to make.
API_VERSION = 'beta'
