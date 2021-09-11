import json
import uuid
import datetime
from logging import exception
import flask
from flask.json import jsonify
from oauthlib.oauth2.rfc6749.clients import base
from requests_oauthlib import OAuth2Session

from functools import wraps

AUTHORITY_URL = 'https://login.microsoftonline.com/common'
AUTH_ENDPOINT = '/oauth2/v2.0/authorize'
TOKEN_ENDPOINT = '/oauth2/v2.0/token'
CLIENT_ID = 'YOUR_CLIENT_ID_HERE'
REDIRECT_URI = 'http://localhost:5000/login/authorized'
CLIENT_SECRET = 'YOUR_CLIENT_SECRET_HERE'
SCOPES = [
  "User.Read",
  "Chat.Read"
]

RESOURCE = 'https://graph.microsoft.com/'
API_VERSION = 'beta' # NOTE: the 'beta' channel is required for the API calls we need to make.


APP = flask.Flask(__name__, static_folder='static')
APP.debug = True
APP.secret_key = 'development'
APP.config['SESSION_TYPE'] = 'filesystem'

if APP.secret_key == 'development':
  import os
  os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'  # allows http requests
  os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'  # allows tokens to contain additional permissions

MSGRAPH = OAuth2Session(CLIENT_ID, redirect_uri=REDIRECT_URI, scope=SCOPES)

# Render the homepage.
@APP.route('/')
def homepage():
  if 'access_token' in flask.session:
    return flask.redirect(flask.url_for('my_data'))

  return flask.render_template('index.html')

@APP.route('/login')
def login():
  flask.session.clear()
  auth_url, state = MSGRAPH.authorization_url(AUTHORITY_URL + AUTH_ENDPOINT)
  flask.session['state'] = state
  print(auth_url)
  return flask.redirect(auth_url)

@APP.route('/login/authorized')
def authorized():
  if flask.session.get('state') and str(flask.session['state']) != str(flask.request.args.get('state')):
    raise Exception('state returned to redirect URL does not match!')
  
  token = MSGRAPH.fetch_token(
    AUTHORITY_URL + TOKEN_ENDPOINT,
    client_secret=CLIENT_SECRET,
    authorization_response=flask.request.url
  )

  flask.session['access_token'] = token

  return flask.redirect('/')

def requires_auth(f):
  @wraps(f)
  def decorated(*args, **kwargs):
    if 'access_token' not in flask.session:
      return flask.redirect('/login')
    if not MSGRAPH.authorized:
      return flask.redirect('/login')
    return f(*args, **kwargs)
  return decorated

@APP.route('/mydata')
@requires_auth
def my_data():
  base_url = RESOURCE + API_VERSION + '/'
  user_profile = MSGRAPH.get(base_url + 'me', headers=request_headers()).json()
  if 'error' in user_profile:
    return flask.redirect('/login')

  username = user_profile['displayName']
  email = user_profile['userPrincipalName']
  return flask.render_template('mydata.html', username=username, email=email)

@APP.route('/get_all_chats')
@requires_auth
def get_all_chats():
  base_url = RESOURCE + API_VERSION + '/'
  raw_chats = MSGRAPH.get(base_url + 'me/chats?$expand=members', headers=request_headers()).json()

  all_chats = []

  if 'value' in raw_chats:
    index = 0
    total_chats = len(raw_chats['value'])
    print(f'Total chats: {total_chats}')
    for chat in raw_chats['value']:
      chat_id = chat['id']
      chat_index = index+1
      chat_topic = chat['topic']
      chat_type = chat['chatType']
      chat_link = chat['webUrl']

      total_members = len(chat['members'])
      chat_members = f'({total_members})'
      max_members = 4
      processed_members = 0
      for member in chat['members']:
        member_name = member['displayName']
        chat_members += f' {member_name}'

        if processed_members != total_members-1:
          chat_members += ','

        processed_members += 1
        if processed_members >= max_members:
          break
      if processed_members != total_members:
        chat_members += ' ...'
      
      all_chats.append({
        'id': chat_id,
        'index': chat_index,
        'topic': chat_topic,
        'type': chat_type,
        'members': chat_members,
        'link': chat_link
      })
      
      index += 1

  return flask.Response(json.dumps(all_chats), mimetype='application/json')

def json_to_html_chat(data):
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

  # Header
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

  # Messages
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

  # Footer
  html_string += "\t<script src=\"https://cdn.jsdelivr.net/npm/bootstrap@5.1.1/dist/js/bootstrap.bundle.min.js\" integrity=\"sha384-/bQdsTh/da6pkI1MST/rWKFNjaCP5gBSY4sEBT38Q/9RBh9AH40zEOg7Hlq2THRZ\" crossorigin=\"anonymous\"></script>\n"
  html_string += "</body>\n"
  html_string += "</html>"

  return html_string

@APP.route('/get_chat', methods=['POST'])
@requires_auth
def get_chat():
  if not flask.request.is_json:
    return jsonify(message='Error. Input was not JSON.'), 400

  chat_id_key = 'chat_id'
  if chat_id_key not in flask.request.json:
    return jsonify(message='Error. Failed to find chat_id entry.'), 400
  
  chat_id = flask.request.json[chat_id_key]
  if not len(chat_id):
    return jsonify(message='Error: chat_id key was empty.'), 400

  # Get and build chat metadata.
  chat_base_url = RESOURCE + API_VERSION + '/'
  raw_chat = MSGRAPH.get(chat_base_url + f'/me/chats/{chat_id}?$expand=members', headers=request_headers()).json()
  if 'error' in raw_chat:
    return jsonify(message='Error: Unable to retrieve chat metadata.'), 400
  chat_data = {
    'topic': 'Unnamed Chat' if raw_chat['topic'] is None else raw_chat['topic'],
    'type': raw_chat['chatType'],
    'when': raw_chat['createdDateTime'],
    'link': raw_chat['webUrl']
  }
  chat_members = []
  for member in raw_chat['members']:
    member_name = member['displayName']
    member_email = member['email']
    chat_members.append(f'{member_name} ({member_email})')
  chat_data['members'] = chat_members

  base_url = RESOURCE + API_VERSION + '/'
  raw_messages = MSGRAPH.get(base_url + f'me/chats/{chat_id}/messages', headers=request_headers()).json()

  all_messages = []
  if 'value' in raw_messages:
    index = 0
    total_chats = len(raw_messages['value'])
    print(f'Total messages: {total_chats}')
    for msg in raw_messages['value']:
      # Only process user messages... for now.
      if msg['messageType'] != 'message':
        continue

      msg_entry = {}

      msg_entry['from'] = msg['from']['user']['displayName']
      msg_entry['when'] = msg['createdDateTime']

      msg_entry['type'] = msg['body']['contentType']

      # Only handle certain types of content.
      if msg['body']['contentType'] in ['text', 'html']:
        msg_entry['content'] = msg['body']['content']
      else:
        msg_entry['content'] = ''
      
      # TODO: Consider attachments.
      # https://docs.microsoft.com/en-us/graph/api/attachment-get?view=graph-rest-beta&tabs=http
      # Maybe I can have them downloaded in advance then inserted in with a link or something?

      all_messages.append(msg_entry)
    #end for
  #end if

  # If no messages are present, return an error.
  if not len(all_messages):
    return jsonify(message='Error: No valid messages found in chat.'), 400

  final_data = {
    'chat': chat_data,
    'messages': all_messages
  }

  random_filename = str(uuid.uuid4())
  with open('static/files/' + random_filename + '.json', 'w+', encoding="utf-8") as out_file:
    out_file.write(json.dumps(final_data))

  with open('static/files/' + random_filename + '.html', 'w+', encoding="utf-8") as out_file:
    out_file.write(json_to_html_chat(final_data))
  
  return flask.send_file('static/files/' + random_filename + '.html', as_attachment=True)
  # return flask.send_from_directory('static/files', random_filename, as_attachment=True, max_age=0)
  # return flask.Response(result_str, mimetype='application/json')

@APP.route('/logout')
def logout():
  flask.session.clear()
  return flask.redirect('/')

# Returns a dictionary of default HTTP headers for Graph API calls.
def request_headers(headers=None):
  default_headers = {
    'SdkVersion': 'sample-python-flask',
    'x-client-SKU': 'sample-python-flask',
    'client-request-id': str(uuid.uuid4()),
    'return-client-request-id': 'true'
  }
  if headers:
    default_headers.update(headers)
  return default_headers

if __name__ == '__main__':
  APP.run()
