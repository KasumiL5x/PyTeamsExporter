import re
import os
import shutil
import json
import uuid
import flask
from flask.json import jsonify
from requests_oauthlib import OAuth2Session

from functools import wraps

# Fill these in from your Azure app (see https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app).
CLIENT_ID = '25c4a000-e26f-40a6-897c-3c79a5ef583a'
CLIENT_SECRET = '29eZPOzTtsFO.c9PMD0XF~XGphLf~bK6~F'
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

def get_blank_oauth(existing_state=None):
  """Creates a new blank OAuth2Session object with no token. This should be used for creation and authorization."""
  return OAuth2Session(CLIENT_ID, redirect_uri=REDIRECT_URI, scope=SCOPES, state=existing_state)

def get_authorized_oauth():
  """
  Creates a new OAuth2Session object based on the user's current access token.
  This should only be called from places where @requires_auth has passed to ensure a valid token.
  """

  refresh_url = AUTHORITY_URL + TOKEN_ENDPOINT
  extra = {
    'client_id': CLIENT_ID,
    'client_secret': CLIENT_SECRET
  }

  def save_new_token(tok):
    flask.session['access_token'] = tok
  
  current_token = flask.session['access_token']
  current_state = flask.session['state']

  return OAuth2Session(
    CLIENT_ID,
    token=current_token,
    redirect_uri=REDIRECT_URI,
    scope=SCOPES,
    state=current_state,
    auto_refresh_url=refresh_url,
    auto_refresh_kwargs=extra,
    token_updater=save_new_token
  )
#end

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

  oauth = get_blank_oauth()

  flask.session.clear()
  auth_url, state = oauth.authorization_url(AUTHORITY_URL + AUTH_ENDPOINT)
  flask.session['state'] = state
  return flask.redirect(auth_url)
#end

@APP.route('/login/authorized')
def authorized():
  """Obtains and stores the access token for an authorized login."""

  if flask.session.get('state') and str(flask.session['state']) != str(flask.request.args.get('state')):
    raise Exception('state returned to redirect URL does not match!')

  oauth = get_blank_oauth(existing_state=flask.session['state'])
  
  token = oauth.fetch_token(
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
    # Doesn't have a valid access token.
    if 'access_token' not in flask.session:
      return flask.redirect('/login')
    # Check if it's authorized. This should always work unless the token has expired.
    if not get_authorized_oauth().authorized:
      return flask.redirect('/login')
    return f(*args, **kwargs)
  #end
  return decorated
#end

@APP.route('/mydata')
@requires_auth
def my_data():
  """Renders the 'my data' page if the user is logged in."""

  oauth = get_authorized_oauth()

  # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-beta&tabs=http
  base_url = RESOURCE + API_VERSION + '/'
  user_profile = oauth.get(base_url + 'me', headers=request_headers()).json()
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

  oauth = get_authorized_oauth()

  print('Getting all chats...')

  # Get this user's display name so that we can find out who they are in a oneOnOne chat.
  # I'd prefer to do this with email, but we have two emails and it seems to be different for account vs. Teams.
  # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-beta&tabs=http
  base_url = RESOURCE + API_VERSION + '/'
  user_profile = oauth.get(base_url + 'me', headers=request_headers()).json()
  user_name = user_profile['displayName']

  # https://docs.microsoft.com/en-us/graph/api/chat-list?view=graph-rest-beta&tabs=http
  next_link_key = '@odata.nextLink'
  raw_chats = []
  base_url = RESOURCE + API_VERSION + '/'
  next_chat_url = base_url + 'me/chats?$expand=members'
  req_index = 0
  while True:
    if req_index == 0:
      print(f'  Request {req_index} ', end='', flush=True)
    else:
      print(f'{req_index} ', end='', flush=True)

    # Get this round's chats.
    tmp_chats = oauth.get(next_chat_url, headers=request_headers()).json()
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
  print(f'\nTotal chats: {total_chats}')
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
    members = [x for x in members if x is not None]
    members_str = ', '.join(members)
    if total_members > max_members:
      members_str += ', ...'
    
    # If the chat is a oneOnOne, then there's no topic. Find the other user and make the topic their name.
    # # Check for 2 explicitly as you can technically chat with yourself, which is 1.
    if (chat_type == 'oneOnOne') and len(members_data) >= 2:
      user_index = [idx for idx, el in enumerate(members_data) if el['displayName'] == user_name]
      if len(user_index):
        user_index = user_index[0]
        other_name = members_data[1 - user_index]['displayName']
        chat_topic = f'Chat with {other_name}'
    
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
    html_string += f"\t\t\t<span class=\"badge bg-primary\">{msg_timestamp}</span>\n"
    html_string += "\t\t</li>\n"
  #end
  html_string += "\t</ul>\n\n"

  # Footer.
  html_string += "\t<script src=\"https://cdn.jsdelivr.net/npm/bootstrap@5.1.1/dist/js/bootstrap.bundle.min.js\" integrity=\"sha384-/bQdsTh/da6pkI1MST/rWKFNjaCP5gBSY4sEBT38Q/9RBh9AH40zEOg7Hlq2THRZ\" crossorigin=\"anonymous\"></script>\n"
  html_string += "</body>\n"
  html_string += "</html>"

  return html_string
#end

supported_image_types = ['.png', '.jpg', '.gif', '.svg', '.webp']
def content_type_to_file_ext(content_type):
  if content_type == 'image/png':
    return '.png'
  elif content_type == 'image/jpeg':
    return '.jpg'
  elif content_type == 'image/gif':
    return '.gif'
  elif content_type == 'image/svg+xml':
    return '.svg'
  elif content_type == 'image/webp':
    return '.webp'
  else:
    return None

@APP.route('/get_chat', methods=['POST'])
@requires_auth
def get_chat():
  """
  Given a chat id, retrieves chat metadata and all messages. This is stored in JSON and pretty HTML
  on the server, and the generated HTML file is then sent back to the browser as a file to download.

  Expected JSON input to this request is:
  {
    "chat_id": "the_chat_id",
    "include_attachments": true/false
  }
  """

  # Broadly sanity check input data type.
  if not flask.request.is_json:
    return jsonify(message='Input was not JSON.'), 400

  # Make sure all required keys are present.
  chat_id_key = 'chat_id'
  include_attachments_key = 'include_attachments'
  if chat_id_key not in flask.request.json:
    return jsonify(message=f'Failed to find {chat_id_key} entry.'), 400
  if include_attachments_key not in flask.request.json:
    return jsonify(message=f'Failed to find {include_attachments_key} entry.'), 400
  
  # Make sure the keys have valid data.
  chat_id = flask.request.json[chat_id_key]
  if not len(chat_id):
    return jsonify(message=f'{chat_id_key} value was empty.'), 400
  include_attachments = flask.request.json[include_attachments_key]
  if not isinstance(include_attachments, bool):
    return jsonify(message=f'{include_attachments_key} value was the wrong type (expected bool).'), 400

  oauth = get_authorized_oauth()

  print(f'Processing chat {chat_id}')

  # Retrieve the chat metadata.
  # https://docs.microsoft.com/en-us/graph/api/chat-get?view=graph-rest-beta&tabs=http
  chat_base_url = RESOURCE + API_VERSION + '/'
  raw_chat = oauth.get(chat_base_url + f'me/chats/{chat_id}?$expand=members', headers=request_headers()).json()
  if 'error' in raw_chat:
    print('Chat response contains an error.')
    return jsonify(message='Unable to retrieve chat metadata.'), 400
  
  # Major failure points are now over, so we can safely create the root folder for this chat now.
  random_filename = str(uuid.uuid4())
  root_folder = 'static/files/' + random_filename + '/'
  # https://stackoverflow.com/a/50901481
  old_umask = os.umask(0o666)
  os.makedirs(root_folder, exist_ok=True)
  # Repeat for attachments if needed.
  attachments_folder = 'attachments'
  attachments_root_folder = root_folder + attachments_folder + '/'
  os.makedirs(attachments_root_folder, exist_ok=True) # NOTE: Always make this as hosted images use the same folder... for now.
  os.umask(old_umask)

  # Strings in attachment URLs that will force them to be not downloaded and just linked to.
  attachment_ignores = ['sharepoint.com']

  
  # Build chat metadata.
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

  # Write out the chat metadata to file.
  print('  Writing metadata...', end='')
  with open(root_folder + 'metadata.json', 'w', encoding="utf-8") as out_file:
    out_file.write(json.dumps(chat_data, indent=2))
  print('done!')

  print(f'  Processing messages...')

  # https://docs.microsoft.com/en-us/graph/api/chat-list-messages?view=graph-rest-beta&tabs=http
  next_link_key = '@odata.nextLink'
  raw_messages = []
  base_url = RESOURCE + API_VERSION + '/'
  last_msg_url = base_url + f'me/chats/{chat_id}/messages?$top=50'
  req_index = 0
  while True:
    if req_index == 0:
      print(f'  Request {req_index} ', end='', flush=True)
    else:
      print(f'{req_index} ', end='', flush=True)

    # Get this round's messages.
    tmp_messages = oauth.get(last_msg_url, headers=request_headers()).json()
    # If any response fails, just exit out.
    if ('error' in tmp_messages) or ('value' not in tmp_messages):
      break

    # NOTE: The way the below messages are appended is awkward, but here's why it's like this.
    # There seems to be a bug where the @odata.nextLink repeats forever, creating infinite requests.
    # Because of this bug, we can't guarantee that data needs adding NOW before the below key check.
    # If there is no next key, we do need to add the current response and break from the loop.
    # However, if there is a key, we need to conditionally add the current response based on the above bug.
    # That's why the below code is a bit fugly, but it does work.
    # If this is not done, then the latest message will be duplicated at the start...?
    # TODO: Investigate this in more detail as it may be an oversight in other unused dictionary values.

    if next_link_key not in tmp_messages or tmp_messages[next_link_key] is None:
      # Update the raw messages with this request's response.
      raw_messages.extend(tmp_messages['value'])
      # If there are no more next links available, we're done.
      break
    else:
      # There seems to be a bug where the same 'next link' is returned. Exit out if this happens.
      if last_msg_url == tmp_messages[next_link_key]:
        break
      # Otherwise, update the url for the next request.
      last_msg_url = tmp_messages[next_link_key]
      # Update the raw messages with this request's response.
      raw_messages.extend(tmp_messages['value'])
    #end if

    req_index += 1
  #end while
  print('')

  attachment_lookup = {} # {id:index_in_all_attachments}
  all_attachments = []
  all_messages = []
  total_hosted_images = 0
  for msg in raw_messages:
    # Only process user messages... for now.
    if msg['messageType'] != 'message':
      continue

    # Build message data.
    msg_entry = {}
    #
    msg_user = msg['from']['user']
    is_user_valid = (msg_user is not None) and ('displayName' in msg_user) and (msg_user['displayName'] is not None)
    msg_entry['from'] = 'INVALID USER' if not is_user_valid else msg_user['displayName']
    #
    msg_entry['when'] = msg['createdDateTime']
    msg_entry['type'] = msg['body']['contentType']
    if msg_entry['type'] in ['text', 'html']:
      msg_entry['content'] = msg['body']['content']
    else:
      msg_entry['content'] = ''

    # Messages can host content such as images (and perhaps videos).
    # To display these, they need to be downloaded locally and the HTML needs swapping out.
    # The below code handles <img> tags with hosed content.
    all_img_tags = re.findall(r"<img\s*.*?>", msg_entry['content'])
    # Download each of the images from the raw bytes that MSGRAPH provides.
    hosted_img_index = 0
    for img_tag in all_img_tags:
      # This will pull all images (including emoji), so only process those with a MSGRAPH url.
      if 'graph.microsoft.com' not in img_tag:
        continue

      print(f'  Downloading hosted image {hosted_img_index+1}...', end='', flush=True)

      # Extract and request the actual data.
      img_src = re.findall(r"src=\"(.+?)\"", img_tag)[0]
      img_data = oauth.get(img_src, headers=request_headers())

      # Create a random filename for the file (reading the type from content-type).
      img_name = str(uuid.uuid4())
      img_type = img_data.headers['content-type']
      file_ext = content_type_to_file_ext(img_type)
      if file_ext is None:
        print(f'Warning: Undetected <img> type: {img_type}')
        continue
      else:
        img_name += file_ext

      # Write out the file.
      with open(attachments_root_folder + img_name, 'wb') as out_file:
        out_file.write(img_data.content)

      # Swap the src for the local version.
      img_tag_new = re.sub(r"src=\"(.+?)\"", f"src=\"{attachments_folder}/{img_name}\"", img_tag)
      # These tags don't have any class, so replace the <img at the start with an appended Bootstrap tag.
      img_tag_new = re.sub("^<img", "<img class=\"img-fluid img-thumbnail\"", img_tag_new)
      # Replace the original string.
      msg_entry['content'] = msg_entry['content'].replace(img_tag, img_tag_new)

      print('done!')

      hosted_img_index += 1
      total_hosted_images += 1
    #end for
    
    # Each message determines which attachments it includes. We also want them.
    if include_attachments:
      # Build a list of all attachments for this message.
      for attachment in msg['attachments']:
        # Thumbnails are used for link previews. I'm ignoring these.
        # Ignoring code snippets for now. Will process them later.
        ignore_types = ['application/vnd.microsoft.card.thumbnail', 'application/vnd.microsoft.card.codesnippet', 'application/vnd.microsoft.card.adaptive']
        if attachment['contentType'] in ignore_types:
          continue

        # TODO: Process code snippets.

        attachment_entry = {
          'id': attachment['id'],
          'link': attachment['contentUrl'],
          'name': attachment['name']
        }

        # Links to tab pages of a chat count as attachments and always start with "tab::" in their ID. Ignore them.
        if attachment_entry['id'].startswith('tab::'):
          continue

        # NOTE: Sometimes the name is empty. I should be ignoring those that cause it, but debug printing just in case.
        if(attachment_entry['name'] is None):
          print('! WARNING ! Attachment name is null.')
          print(json.dumps(attachment, indent=2))
          continue

        # Local path to the file relative to the root folder.
        attachment_entry['path'] = attachments_folder + '/' + attachment_entry['name']

        # Add this attachment to the list of all attachments and add a lookup entry into that based on its ID.
        # These values are used below and later when actually processing and downloading the attachments.
        all_attachments.append(attachment_entry)
        attachment_lookup[attachment_entry['id']] = len(all_attachments)-1
      #end for

      # Attachments are inserted with a custom <attachment> tag. This code replaces those tags accordingly.
      all_attachment_tags = re.findall(r"<attachment\s*.*?><\/attachment>", msg_entry['content'])
      # Build a new HTML tag for each of the attachment entries.
      new_attachment_tags = []
      for tag in all_attachment_tags:
        tag_id = re.findall(r"id=\"(.+?)\"", tag)[0]
        # We don't capture all attachments, so ignore those that don't have a lookup value.
        if tag_id not in attachment_lookup:
          continue

        # Lookup the actual attachment that we saved above based on the ID.
        original_attachment = all_attachments[attachment_lookup[tag_id]]
        # Build out several properties for the URL.
        if original_attachment['link'] is None:
          print(f'ERROR: Attachment has no link: {original_attachment}')
          continue
        use_original_link = any([x in original_attachment['link'] for x in attachment_ignores])
        attachment_path = original_attachment['link'] if use_original_link else original_attachment['path']
        attachment_name = original_attachment['name']
        # If an image extension is in the name, use an <img> tag, otherwise use a standard <a> tag.
        if any([attachment_path.endswith(x) for x in supported_image_types]):
          new_attachment_tags.append(f'<img src=\"{attachment_path}\" class=\"img-fluid img-thumbnail\">')
        else:
          new_attachment_tags.append(f'<a href=\"{attachment_path}\" target=\"_blank\">ATTACHMENT: {attachment_name}</a>')
      #end for

      # Replace the original tags with the new tags.
      for tag, new_tag in zip(all_attachment_tags, new_attachment_tags):
        msg_entry['content'] = msg_entry['content'].replace(tag, new_tag)
      #end for
    #end if include_attachments

    all_messages.append(msg_entry)
  #end for

  print(f'  Total messages: {len(all_messages)}; Total attachments: {len(all_attachments)}; Total hosted content: {total_hosted_images}')
  if len(all_attachments):
    print(f'  Processing attachments...')
    print(f'    Writing attachments.json...', end='', flush=True)
    with open(root_folder + 'attachments.json', 'w', encoding="utf-8") as out_file:
      out_file.write(json.dumps(all_attachments, indent=2))
    print('done!')
    for idx, attachment in enumerate(all_attachments):
      print(f'    {idx+1}/{len(all_attachments)}...', end='', flush=True)

      should_ignore = any([x in original_attachment['link'] for x in attachment_ignores])
      if should_ignore:
        print('ignored (protected link).')
        continue

      # TODO: Download attachments here with failsafe checking (404, 401, etc.).
      # att_req = oauth.get(attachment['link'], headers=request_headers())
      # print(att_req)
      # print(att_req.content)
      print('done!')
    #end for
    print(f'  Done!')
  #end if

  # Concatenate the final data that will be converted and saved out.
  final_data = {
    'chat': chat_data,
    'messages': all_messages
  }

  # Write out the dictionary to a raw JSON file.
  with open(root_folder + 'chat.json', 'w', encoding="utf-8") as out_file:
    out_file.write(json.dumps(final_data, indent=2))

  # Convert the dictionary into a pretty HTML page and write that out.
  with open(root_folder + 'chat.html', 'w', encoding="utf-8") as out_file:
    out_file.write(json_to_html_chat(final_data))

  # Create a zip of the root folder.
  print('  Compressing data...', end='', flush=True)
  shutil.make_archive(f'static/files/{random_filename}', 'zip', root_folder)
  print('done!')

  print('Done!')

  return flask.send_file(f'static/files/{random_filename}' + '.zip', as_attachment=True, mimetype='application/octet-stream')

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
  # Make sure the static/files directory definitely exists upon start.
  static_files_path = 'static/files/'
  # https://stackoverflow.com/a/50901481
  old_umask = os.umask(0o666)
  os.makedirs(static_files_path, exist_ok=True)
  os.umask(old_umask)

  # Run the server.
  APP.run()
