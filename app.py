import json
import uuid
import flask
from flask.json import jsonify
from requests_oauthlib import OAuth2Session

from functools import wraps

# Fill these in from your Azure app (see https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app).
CLIENT_ID = 'YOUR_CLIENT_ID'
CLIENT_SECRET = 'YOUR_CLIENT_SECRET'
REDIRECT_URI = 'http://localhost:5000/login/authorized'
SCOPES = [
  "User.Read",
  "Chat.Read"
]
# URLs and endpoints for authorization.
AUTHORITY_URL = 'https://login.microsoftonline.com/common'
AUTH_ENDPOINT = '/oauth2/v2.0/authorize'
TOKEN_ENDPOINT = '/oauth2/v2.0/token'
# Microsoft Graph API configuration.
RESOURCE = 'https://graph.microsoft.com/'
API_VERSION = 'beta' # NOTE: the 'beta' channel is required for the API calls we need to make.


# Flask initialization.
APP = flask.Flask(__name__, static_folder='static')
APP.debug = True
APP.secret_key = 'development'
APP.config['SESSION_TYPE'] = 'filesystem'
if APP.secret_key == 'development':
  import os
  os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'  # allows http requests
  os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'  # allows tokens to contain additional permissions

# Create initial connection to Microsoft Graph.
MSGRAPH = OAuth2Session(CLIENT_ID, redirect_uri=REDIRECT_URI, scope=SCOPES)

@APP.route('/')
def homepage():
  """Render the homepage."""

  # If user is already logged in, redirect them to the data page.
  if 'access_token' in flask.session:
    return flask.redirect('/mydata')

  return flask.render_template('index.html')
#end

@APP.route('/login')
def login():
  """Redirects the user to the Microsoft login page (which in turn redirects back to this app)."""

  flask.session.clear()
  auth_url, state = MSGRAPH.authorization_url(AUTHORITY_URL + AUTH_ENDPOINT)
  flask.session['state'] = state
  return flask.redirect(auth_url)
#end

@APP.route('/login/authorized')
def authorized():
  """Obtains and stores the access token for an authorized login."""

  if flask.session.get('state') and str(flask.session['state']) != str(flask.request.args.get('state')):
    raise Exception('state returned to redirect URL does not match!')
  
  token = MSGRAPH.fetch_token(
    AUTHORITY_URL + TOKEN_ENDPOINT,
    client_secret=CLIENT_SECRET,
    authorization_response=flask.request.url
  )
  flask.session['access_token'] = token

  return flask.redirect('/')
#end

@APP.route('/logout')
def logout():
  """Does what it says on the tin."""

  flask.session.clear()
  return flask.redirect('/')
#end

def requires_auth(f):
  """Defines a decorator so that functions can require an authorized user to work."""

  @wraps(f)
  def decorated(*args, **kwargs):
    if 'access_token' not in flask.session:
      return flask.redirect('/login')
    if not MSGRAPH.authorized:
      return flask.redirect('/login')
    return f(*args, **kwargs)
  #end
  return decorated
#end

@APP.route('/mydata')
@requires_auth
def my_data():
  """Renders the 'my data' page if the user is logged in."""

  # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-beta&tabs=http
  base_url = RESOURCE + API_VERSION + '/'
  user_profile = MSGRAPH.get(base_url + 'me', headers=request_headers()).json()
  # If the request gives an error, try to get the user to login again.
  if 'error' in user_profile:
    return flask.redirect('/login')

  username = user_profile['displayName']
  email = user_profile['userPrincipalName']
  return flask.render_template('mydata.html', username=username, email=email)
#end

@APP.route('/get_all_chats', methods=['GET'])
@requires_auth
def get_all_chats():
  """Returns a JSON dictionary of all chats and associated metadata. This is NOT messages, but the chats themselves."""
  # NOTE: See `all_chats` below for the structure of the returned JSON.

  print('Getting all chats...')

  # https://docs.microsoft.com/en-us/graph/api/chat-list?view=graph-rest-beta&tabs=http
  next_link_key = '@odata.nextLink'
  raw_chats = []
  base_url = RESOURCE + API_VERSION + '/'
  next_chat_url = base_url + 'me/chats?$expand=members'
  req_index = 0
  while True:
    print(f'\trequest {req_index}')
    # Get this round's chats.
    tmp_chats = MSGRAPH.get(next_chat_url, headers=request_headers()).json()
    if ('error' in tmp_chats) or ('value' not in tmp_chats):
      break

    # Update raw chats with this request's response.
    raw_chats.extend(tmp_chats['value'])

    # Process the next link if it's available.
    if next_link_key not in tmp_chats or tmp_chats[next_link_key] is None:
      break
    else:
      next_chat_url = tmp_chats[next_link_key]
    
    req_index += 1
  #end while

  all_chats = []

  total_chats = len(raw_chats)
  print(f'Total chats: {total_chats}')
  for chat in raw_chats:
    chat_id = chat['id']
    chat_topic = chat['topic']
    chat_type = chat['chatType']
    chat_link = chat['webUrl']

    # Concatenate members with a maximum number.
    members_data = chat['members']
    total_members = len(members_data)
    max_members = 4
    members = [x['displayName'] for x in members_data[:min(total_members, max_members)]]
    members_str = ', '.join(members)
    if total_members > max_members:
      members_str += ', ...'
    
    all_chats.append({
      'id': chat_id,
      'topic': chat_topic,
      'type': chat_type,
      'members': members_str,
      'link': chat_link
    })
  #end for

  return flask.Response(json.dumps(all_chats), mimetype='application/json')
#end

def json_to_html_chat(data):
  """Takes a data array from within `get_chat` and converts it into a pretty HTML document."""
  html_string = ''

  chat_data = data['chat']
  chat_topic = chat_data['topic']
  chat_type = chat_data['type']
  members = chat_data['members']
  chat_when = chat_data['when']
  chat_link = chat_data['link']

  header_text = f'{chat_topic} &mdash; {chat_type} with {len(members)} participants'
  members_list = ', '.join(members)

  messages_data = data['messages']

  # Header.
  html_string += "<!DOCTYPE html>\n"
  html_string += "<html>\n"
  html_string += "<head>\n"
  html_string += f"\t<title>{header_text} | PyTeamsExporter</title>\n"
  html_string += "\t<meta charset=\"utf-8\">\n"
  html_string += "\t<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, shrink-to-fit=no\">\n"
  html_string += "\t<link href=\"https://cdn.jsdelivr.net/npm/bootstrap@5.1.1/dist/css/bootstrap.min.css\" rel=\"stylesheet\" integrity=\"sha384-F3w7mX95PdgyTmZZMECAngseQB83DfGTowi0iMjiWaeVhAn4FJkqJByhZMI3AhiU\" crossorigin=\"anonymous\">\n"
  html_string += "</head>\n"
  html_string += "<body>\n"

  # Navbar with title and details button.
  html_string += "\t<nav class=\"navbar navbar-light bg-light\">\n"
  html_string += "\t\t<div class=\"container-fluid\">\n"
  html_string += f"\t\t\t<span class=\"navbar-brand mb-0 h1\">{header_text}</span>\n"
  html_string += "\t\t\t<button class=\"btn btn-dark\" type=\"button\" data-bs-toggle=\"collapse\" data-bs-target=\"#navbarToggleExternalContent\" aria-controls=\"navbarToggleExternalContent\" aria-expanded=\"false\" aria-label=\"See details\">\n"
  html_string += "\t\t\t\tSee details\n"
  html_string += "\t\t\t</button>\n"
  html_string += "\t</nav>\n"
  # The hidden details panel.
  html_string += "\t<div class=\"collapse\" id=\"navbarToggleExternalContent\">\n"
  html_string += "\t\t<div class=\"bg-dark p-4\">\n"
  #
  html_string += "\t\t\t<h5 class=\"text-white h4\">Messages</h5>\n"
  html_string += f"\t\t\t<span class=\"text-light\">{len(messages_data)}</span>\n"
  html_string += "\t\t\t<br/><br/>\n"
  #
  html_string += "\t\t\t<h5 class=\"text-white h4\">Created</h5>\n"
  html_string += f"\t\t\t<span class=\"text-light\">{chat_when}</span>\n"
  html_string += "\t\t\t<br/><br/>\n"
  #
  html_string += "\t\t\t<h5 class=\"text-white h4\">Teams Link</h5>\n"
  html_string += f"\t\t\t<a href=\"{chat_link}\" target=\"_blank\" class=\"text-light\">Link</a>\n"
  html_string += "\t\t\t<br/><br/>\n"
  #
  html_string += "\t\t\t<h5 class=\"text-white h4\">Members</h5>\n"
  html_string += f"\t\t\t<span class=\"text-light\">{members_list}</span>\n"
  html_string += "\t\t\t<br/><br/>\n"
  #
  html_string += "\t\t</div>\n"
  html_string += "\t</div>\n"

  # Messages.
  html_string += "\t<ul class=\"list-group\">\n"
  for msg in messages_data:
    msg_from = msg['from']
    msg_content = msg['content']
    msg_timestamp = msg['when']

    html_string += "\t\t<li class=\"list-group-item list-group-item-action d-flex justify-content-between align-items-start\">\n"
    html_string += "\t\t\t<div class=\"ms-2 me-auto\">\n"
    html_string += f"\t\t\t\t<div class=\"fw-bold\">{msg_from}</div>\n"
    html_string += f"\t\t\t\t{msg_content}\n"
    html_string += "\t\t\t</div>\n"
    html_string += f"\t\t\t<span class=\"badge bg-primary rounded-pill\">{msg_timestamp}</span>\n"
    html_string += "\t\t</li>\n"
    html_string += "\t</ul>\n\n"
  #end

  # Footer.
  html_string += "\t<script src=\"https://cdn.jsdelivr.net/npm/bootstrap@5.1.1/dist/js/bootstrap.bundle.min.js\" integrity=\"sha384-/bQdsTh/da6pkI1MST/rWKFNjaCP5gBSY4sEBT38Q/9RBh9AH40zEOg7Hlq2THRZ\" crossorigin=\"anonymous\"></script>\n"
  html_string += "</body>\n"
  html_string += "</html>"

  return html_string
#end

@APP.route('/get_chat', methods=['POST'])
@requires_auth
def get_chat():
  """
  Given a chat id, retrieves chat metadata and all messages. This is stored in JSON and pretty HTML
  on the server, and the generated HTML file is then sent back to the browser as a file to download.
  """

  # Broadly sanity check input data type.
  if not flask.request.is_json:
    return jsonify(message='Input was not JSON.'), 400

  # Make sure the appropriate JSON entry is present.
  chat_id_key = 'chat_id'
  if chat_id_key not in flask.request.json:
    return jsonify(message='Failed to find chat_id entry.'), 400
  
  # Make sure the passed in value actually has data.
  chat_id = flask.request.json[chat_id_key]
  if not len(chat_id):
    return jsonify(message='chat_id value was empty.'), 400

  print(f'Processing chat {chat_id}...')

  # https://docs.microsoft.com/en-us/graph/api/chat-get?view=graph-rest-beta&tabs=http
  chat_base_url = RESOURCE + API_VERSION + '/'
  raw_chat = MSGRAPH.get(chat_base_url + f'/me/chats/{chat_id}?$expand=members', headers=request_headers()).json()
  # Double check that we can get the chat data (i.e., is the ID correct and does it return valid data).
  if 'error' in raw_chat:
    print('Chat response contains an error.')
    return jsonify(message='Unable to retrieve chat metadata.'), 400
  
  # Build chat metadata dictionary.
  chat_members = []
  for member in raw_chat['members']:
    member_name = member['displayName']
    member_email = member['email']
    chat_members.append(f'{member_name} ({member_email})')
  #end
  chat_data = {
    'topic': 'Unnamed Chat' if raw_chat['topic'] is None else raw_chat['topic'],
    'type': raw_chat['chatType'],
    'when': raw_chat['createdDateTime'],
    'link': raw_chat['webUrl'],
    'members': chat_members
  }

  print(f'\tprocessing messages for chat {chat_id}...')

  # https://docs.microsoft.com/en-us/graph/api/chat-list-messages?view=graph-rest-beta&tabs=http
  next_link_key = '@odata.nextLink'
  raw_messages = []
  base_url = RESOURCE + API_VERSION + '/'
  last_msg_url = base_url + f'me/chats/{chat_id}/messages?$top=50'
  req_index = 0
  while True:
    print(f'\t\trequest {req_index}')

    # Get this round's messages.
    tmp_messages = MSGRAPH.get(last_msg_url, headers=request_headers()).json()
    # If any response fails, just exit out.
    if ('error' in tmp_messages) or ('value' not in tmp_messages):
      break

    # Update the raw messages with this request's response.
    raw_messages.extend(tmp_messages['value'])

    if next_link_key not in tmp_messages or tmp_messages[next_link_key] is None:
      # If there are no more next links available, we're done.
      break
    else:
      # Otherwise, update the url for the next request.
      last_msg_url = tmp_messages[next_link_key]
    #end if

    req_index += 1
  #end while

  all_messages = []
  for msg in raw_messages:
    # Only process user messages... for now.
    if msg['messageType'] != 'message':
      continue

    msg_entry = {}
    msg_entry['from'] = msg['from']['user']['displayName']
    msg_entry['when'] = msg['createdDateTime']
    msg_entry['type'] = msg['body']['contentType']
    # Only handle certain types of content for now.
    if msg['body']['contentType'] in ['text', 'html']:
      msg_entry['content'] = msg['body']['content']
    else:
      msg_entry['content'] = ''

    all_messages.append(msg_entry)
    
    # TODO: Consider attachments.
    # https://docs.microsoft.com/en-us/graph/api/attachment-get?view=graph-rest-beta&tabs=http
    # Maybe I can have them downloaded in advance then inserted in with a link or something?
    # Perhaps I can download them all, zip them up, then send the zip back?
    # Perhaps this is a separate button in the UI in each row?
    # Need to think about this. Messages for now, as those are the focus.
  #end for

  # If no messages are present, return an error.
  if not len(all_messages):
    return jsonify(message='No valid messages found in chat.'), 400

  # Concatenate the final data that will be converted and saved out.
  final_data = {
    'chat': chat_data,
    'messages': all_messages
  }

  # TODO: Come up with a more intelligent filename than a random uuid.
  random_filename = str(uuid.uuid4())

  print('\tWriting files to disk...')

  # Write out the dictionary to a raw JSON file.
  with open('static/files/' + random_filename + '.json', 'w+', encoding="utf-8") as out_file:
    out_file.write(json.dumps(final_data, indent=2))

  # Convert the dictionary into a pretty HTML page and write that out.
  with open('static/files/' + random_filename + '.html', 'w+', encoding="utf-8") as out_file:
    out_file.write(json_to_html_chat(final_data))

  print('\tDone!')

  # If there's a format, respect it. Otherwise, default to HTML.
  should_return_html = True
  format_key = 'format'
  if format_key in flask.request.json:
    should_return_html = flask.request.json[format_key] == 'html'
  
  if should_return_html:
    # Return the HTML file from disk.
    return flask.send_file('static/files/' + random_filename + '.html', as_attachment=True)
  else:
    # Return the JSON file from disk.
    return flask.send_file('static/files/' + random_filename + '.json', as_attachment=True)

  # NOTE: I read online that this is safer and I shouldn't use `send_file`. Will look into it later.
  # return flask.send_from_directory('static/files', random_filename, as_attachment=True, max_age=0)
#end

def request_headers(headers=None):
  """Returns a dictionary of default HTTP headers for Graph API calls."""

  default_headers = {
    'SdkVersion': 'sample-python-flask',
    'x-client-SKU': 'sample-python-flask',
    'client-request-id': str(uuid.uuid4()),
    'return-client-request-id': 'true'
  }
  if headers:
    default_headers.update(headers)
  return default_headers
#end

if __name__ == '__main__':
  APP.run()
