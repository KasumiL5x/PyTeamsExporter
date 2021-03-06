import re
import os
import shutil
import json
import uuid
import flask
from flask.json import jsonify
from flask_session import Session
from oauthlib import oauth2
from requests_oauthlib import OAuth2Session
from functools import wraps

import utilities as utils
from app_config import *

# Flask initialization.
APP = flask.Flask(__name__, static_folder='static')
APP.debug = True
APP.secret_key = 'development'
APP.config['SESSION_TYPE'] = 'filesystem'
if APP.secret_key == 'development':
  import os
  os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'  # allows http requests
  os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'  # allows tokens to contain additional permissions

# Configure the server-side session object.
Session(APP)

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

def query_endpoint(oauth, endpoint, json=True):
  try:
    res = oauth.get(endpoint, headers=request_headers())
  except oauth2.TokenExpiredError:
    print(f'Warning: Token expired!')
    return None

  if not json:
    return res
  
  res = res.json()
  return None if 'error' in res else res
#end

def query_endpoint_recursive(oauth, endpoint, callback, print_request_index=False):
  next_link_key = '@odata.nextLink'
  next_url = endpoint
  index = 0
  while True:
    if print_request_index:
      if index == 0:
        print(f'Request {index} ', end='', flush=True)
      else:
        print(f'{index} ', end='', flush=True)
    #end

    try:
      # Get this round's data.
      res = oauth.get(next_url, headers=request_headers()).json()
    except oauth2.TokenExpiredError:
      print(f'Warning: Token expired!')
      break
    
    # Check for a bad response.
    if 'error' in res:
      break

    # NOTE: The way the below messages are appended is awkward, but here's why it's like this.
    # There seems to be a bug where the @odata.nextLink repeats forever, creating infinite requests.
    # Because of this bug, we can't guarantee that data needs adding NOW before the below key check.
    # If there is no next key, we do need to add the current response and break from the loop.
    # However, if there is a key, we need to conditionally add the current response based on the above bug.
    # That's why the below code is a bit fugly, but it does work.
    # If this is not done, then the latest message will be duplicated at the start...?
    # TODO: Investigate this in more detail as it may be an oversight in other unused dictionary values.
    
    if res.get(next_link_key) is None:
      callback(res)
      # If there are no more next links available, we're done.
      break
    else:
      # There seems to be a bug where the same 'next link' is returned. Exit out if this happens.
      if next_url == res[next_link_key]:
        break
      # Otherwise, update the url for the next request.
      next_url = res[next_link_key]
      callback(res)
    
    index += 1
  #end
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

@APP.route('/')
def homepage():
  """Render the homepage."""

  # If user is already logged in, redirect them to the data page.
  if 'access_token' in flask.session:
    return flask.redirect('/mydata')

  return flask.render_template('index.html')
#end

import time
@APP.route('/debug_request', methods=['POST'])
@requires_auth
def debug_request():
  flask.session['access_token']['expires_at'] = time.time() - 10

  oauth = get_authorized_oauth()
  data = flask.request.json['data']

  out = []
  def the_callback(res):
    out.extend(res['value'])

  out = query_endpoint(oauth, f'{RESOURCE}{API_VERSION}/me/chats/{data}/messages?$top=50')

  for msg in out['value']:
    msg_id = msg['id']
    msg_content = msg['body']['content']
    hosted_content = query_endpoint(oauth, f'{RESOURCE}{API_VERSION}/chats/{data}/messages/{msg_id}/hostedContents')
    if hosted_content.get('@odata.count', 0) > 0:
      for hc_value in hosted_content['value']:
        hc_id = hc_value['id']
        hc_url = f'"https://graph.microsoft.com/beta/chats/{data}/messages/{msg_id}/hostedContents/{hc_id}/$value'
        msg_content = msg_content.replace(hc_url, f'NEWURL{hc_id}')
        # hc_content = query_endpoint(oauth, f'{RESOURCE}{API_VERSION}/chats/{data}/messages/{msg_id}/hostedContents/{hc_id}')
        # hc_url = hc_content['@odata.context'] + '/$value'
        # print(hc_url)
      print(msg_content)
      print()
      # find all 'value' entries of hosted_content in body.content. Can I just download them then store and .replace() the urls?
    # print(msg['body'])
    # print()
  print('done')

  # query_endpoint_recursive(oauth, f'{RESOURCE}{API_VERSION}/me/chats/{data}/messages?$top=50', the_callback, True)

  # for msg in out:
  #   if len(msg['attachments']) > 0:
  #     print(json.dumps(msg, indent=2))

  # print(json.dumps(out, indent=2))
  return ''
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

  existing_state = flask.session.get('state')

  if existing_state is None:
    print('state key was missing from request.')
    return flask.redirect('/')

  if str(existing_state) != str(flask.request.args.get('state')):
    raise Exception('state returned to redirect URL does not match!')

  oauth = get_blank_oauth(existing_state=existing_state)
  
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

@APP.route('/mydata')
@requires_auth
def my_data():
  """Renders the 'my data' page if the user is logged in."""

  oauth = get_authorized_oauth()

  # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-beta&tabs=http
  user_profile = query_endpoint(oauth, f'{RESOURCE}{API_VERSION}/me')
  # If the request gives an error, try to get the user to login again.
  if user_profile is None:
    return flask.redirect('/login')

  username = user_profile.get('displayName', 'INVALID USER')
  email = user_profile.get('userPrincipalName', 'INVALID EMAIL')
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
  user_profile = query_endpoint(oauth, f'{RESOURCE}{API_VERSION}/me')
  if user_profile is None or 'displayName' not in user_profile:
    return jsonify(message='Unable to retrieve user profile.'), 400
  user_name = user_profile['displayName']

  raw_chats = []
  def chats_callback(res):
    if 'value' in res:
      raw_chats.extend(res['value'])
      
  # https://docs.microsoft.com/en-us/graph/api/chat-list?view=graph-rest-beta&tabs=http
  query_endpoint_recursive(oauth, f'{RESOURCE}{API_VERSION}/me/chats?$expand=members', chats_callback, True)

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

@APP.route('/get_chat', methods=['POST'])
@requires_auth
def get_chat():
  """
  Given a chat id, retrieves chat metadata and all messages. This is stored in JSON and pretty HTML
  on the server, and the generated HTML file is then sent back to the browser as a file to download.

  Expected JSON input to this request is:
  {
    "chat_id": "the_chat_id"
  }
  """

  # Broadly sanity check input data type.
  if not flask.request.is_json:
    print('Error: Input was not JSON.')
    return flask.Response('Input was not JSON.', status=400)

  # Make sure all required keys are present.
  chat_id_key = 'chat_id'
  chat_id = flask.request.json.get(chat_id_key)
  if not isinstance(chat_id, str) or not len(chat_id):
    print(f'Error: Bad key {chat_id_key}.')
    return flask.Response(f'Bad key {chat_id_key}.', status=400)

  # Make sure extra files to be copied are present.
  if not utils.are_extra_files_valid():
    print('Error: Local extra files were not present.')
    return flask.Response('Error: Local extra files were not present.', status=400)

  # Get an authorized oauth instance.
  oauth = get_authorized_oauth()

  print(f'Processing chat {chat_id}')

  # Get this user's display name so that we can find out who they are in a oneOnOne chat.
  # I'd prefer to do this with email, but we have two emails and it seems to be different for account vs. Teams.
  # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-beta&tabs=http
  user_profile = query_endpoint(oauth, f'{RESOURCE}{API_VERSION}/me')
  if user_profile is None or 'displayName' not in user_profile:
    print('Error: Unable to retrieve user profile.')
    return flask.Response('Unable to retrieve user profile.', status=400)
  user_name = user_profile['displayName']

  # Retrieve the chat metadata.
  # https://docs.microsoft.com/en-us/graph/api/chat-get?view=graph-rest-beta&tabs=http
  raw_chat = query_endpoint(oauth, f'{RESOURCE}{API_VERSION}/me/chats/{chat_id}?$expand=members')
  if raw_chat is None:
    print('Error: Unable to retrieve chat metadata.')
    return flask.Response('Unable to retrieve chat metadata.', status=400)
  
  # Create the folders for this chat.
  # https://stackoverflow.com/a/50901481
  old_umask = os.umask(0o000)
  #
  random_filename = str(uuid.uuid4())
  root_folder = 'static/files/' + random_filename + '/'
  os.makedirs(root_folder, exist_ok=True)
  #
  attachments_folder = 'attachments'
  attachments_root_folder = root_folder + attachments_folder + '/'
  os.makedirs(attachments_root_folder, exist_ok=True)
  #
  os.umask(old_umask)
  
  # Build chat metadata.
  chat_topic = raw_chat['topic'] or 'Unnamed Chat'
  chat_members = []
  for member in raw_chat.get('members', []):
    member_name = member.get('displayName', 'INVALID NAME')
    member_email = member.get('email', 'INVALID EMAIL')
    chat_members.append(f'{member_name} ({member_email})')

    # If this is a oneOnOne chat and we are NOT adding the current user, generate a new title.
    if (raw_chat['chatType'] == 'oneOnOne') and (member_name != user_name):
      chat_topic = f'Chat with {member_name}'

  chat_data = {
    'version': EXPORT_VERSION,
    'topic': chat_topic,
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

  raw_messages = []
  def messages_callback(res):
    if 'value' in res:
      raw_messages.extend(res['value'])

  # https://docs.microsoft.com/en-us/graph/api/chat-list-messages?view=graph-rest-beta&tabs=http
  query_endpoint_recursive(oauth, f'{RESOURCE}{API_VERSION}/me/chats/{chat_id}/messages?$top=50', messages_callback, True)
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


    # NOTE: The 'proper' way of handling hostedContents is to use the API, but it frequently hangs when many requests are made.
    # If this bool is true, it will use this API despite it being slow. If it is false, it will manually scrape and request them.
    get_hosted_contents_through_api = False

    if get_hosted_contents_through_api:
      msg_id = msg['id']
      msg_hosted_content_url = f'{RESOURCE}{API_VERSION}/me/chats/{chat_id}/messages/{msg_id}/hostedContents'
      # print(f'Querying hosted content...', end='', flush=True)
      msg_hosted_content = query_endpoint(oauth, msg_hosted_content_url)
      # print('done!')
      if msg_hosted_content is not None and msg_hosted_content.get('@odata.count', 0) > 0:
        msg_hosted_content_items = msg_hosted_content.get('value', [])
        for hc_item in msg_hosted_content_items:
          print(f'Downloading hosted content {total_hosted_images}...', end='', flush=True)

          hc_id = hc_item['id']
          hc_url = f'https://graph.microsoft.com/beta/me/chats/{chat_id}/messages/{msg_id}/hostedContents/{hc_id}/$value'

          # Download the actual data.
          print('query...', end='', flush=True)
          hc_data = query_endpoint(oauth, hc_url, json=False)
          if hc_data is None:
            print(f'Warning: Failed to download hosted content with ID {hc_id}.')
            continue

          # Try to figure out the file type.
          hc_type = hc_data.headers['content-type']
          hc_ext = utils.img_mime_to_ext(hc_type)
          if hc_ext is None:
            hc_ext = utils.video_mime_to_ext(hc_type)
            if hc_ext is None:
              print(f'Warning: Unsupported hosted content mime type {hc_type} with ID {hc_id}.')
              continue
            #end
          #end
          hc_name = str(uuid.uuid4()) + hc_ext

          # Write out the actual file.
          with open(attachments_root_folder + hc_name, 'wb') as out_file:
            out_file.write(hc_data.content)

          # Replace the content URL with this new local URL.
          msg_entry['content'] = msg_entry['content'].replace(hc_url, f'{attachments_folder}/{hc_name}')

          total_hosted_images += 1
          print('done!')
        #end
      #end
    #end (get_hosted_contents_through_api)
    else:
      msg_img_tags = re.findall(r"<img\s*.*?>", msg_entry['content'])
      for img_tag in msg_img_tags:
        # Only process images that have a MSGRAPH url, as hosted content would.
        if 'graph.microsoft.com' not in img_tag:
          continue
        
        print(f'  Downloading hosted image {total_hosted_images}...', end='', flush=True)

        # Extract and request the actual data.
        img_src = re.findall(r"src=\"(.+?)\"", img_tag)[0]
        img_data = query_endpoint(oauth, img_src, json=False)
        if img_data is None:
          print('Warning: Hosted image failed to query. Skipping.')
          continue

        img_type = img_data.headers['content-type']

        # It is possible to get an actual error (like 403 FORBIDDEN), so catch those too.
        is_json = img_type == 'application/json'
        img_error_message = None if not is_json else img_data.json().get('error').get('message')
        if img_error_message is not None:
          print(f'failed ({img_error_message}). Skipping.')
          continue

        # Try to figure out the file type.
        img_ext = utils.img_mime_to_ext(img_type)
        if img_ext is None:
          print(f'Warning: Undetected hosted image type: {img_type}')
          continue
        img_name = str(uuid.uuid4()) + img_ext

        # Write out the actual file.
        with open(attachments_root_folder + img_name, 'wb') as out_file:
          out_file.write(img_data.content)

        # Swap the src for the local version.
        img_tag_new = re.sub(r"src=\"(.+?)\"", f"src=\"{attachments_folder}/{img_name}\"", img_tag)
        # These tags don't have any class, so replace the <img at the start with an appended Bootstrap tag.
        img_tag_new = re.sub("^<img", "<img class=\"img-fluid img-thumbnail\"", img_tag_new)
        # Replace the original string.
        msg_entry['content'] = msg_entry['content'].replace(img_tag, img_tag_new)

        total_hosted_images += 1
        print('done!')
      #end for
    #end (get_hosted_contents_through_api)
    
    # Build a list of all attachments for this message.
    for attachment in msg['attachments']:
      # Thumbnails are used for link previews. I'm ignoring these.
      # Ignoring code snippets for now. Will process them later.
      ignore_types = ['application/vnd.microsoft.card.thumbnail', 'application/vnd.microsoft.card.adaptive', 'application/vnd.microsoft.card.announcement']
      if attachment['contentType'] in ignore_types:
        continue

      # print(json.dumps(attachment, indent=2))

      # Links to tab pages of a chat count as attachments and always start with "tab::" in their ID. Ignore them.
      if attachment['id'].startswith('tab::'):
        continue

      attachment_entry = {
        'id': attachment['id'],
        # TODO: Once I add downloading, I'll likely need to add and later use:
        # ? 'path': 'path_to_local_file_after_downloading',
        # Something like: attachment_entry['path'] = attachments_folder + '/' + attachment_entry['name']
        # If this is true, 'url' should also be set and valid.
        'should_download': False,
        'url': ''
      }

      # Handle external links. This includes special handling for images and videos which can be embedded (but not downloaded).
      if attachment['contentType'] == 'reference':
        url = attachment['contentUrl']
        if url is None:
          print('Warning: Attachment is a link but has no valid URL.')
          print(json.dumps(attachment, indent=2))
          continue
        
        # Handle video links by embedding them.
        if any([url.endswith(x) for x in utils.supported_video_types.values()]):
          file_type = os.path.splitext(url)[-1]
          mime_type = utils.video_ext_to_mime(file_type)
          if mime_type is None:
            print('Warning: Attachment has valid video URL but mime type was not found.')
            print(json.dumps(attachment, indent=2))
            continue
          attachment_entry['html'] = f'<video controls><source src=\"{url}\" type=\"{mime_type}\"/></video></br><a href=\"{url}\" target=\"_blank\">Video Link</a>'
        #end if (video)
        # Handle image links by embedding them.
        elif any([url.endswith(x) for x in utils.supported_image_types.values()]):
          attachment_entry['html'] = f'<img src=\"{url}\" class=\"img-fluid img-thumbnail\" />'
        #end if (image)
        # Handle other links.
        else:
          name = attachment['name']
          if name is None:
            print('Warning: Attachment has a valid URL but no name. Will supplement an unknown name.')
            print(json.dumps(attachment, indent=2))
          link_name = name or 'INVALID NAME'
          attachment_entry['html'] = f'<a href=\"{url}\" target=\"_blank\">ATTACHMENT: {link_name}</a>'
        #end if (other)
      #end if (reference)

      # Process message references.
      if attachment['contentType'] == 'messageReference':
        msgref_content = None if 'content' not in attachment else attachment['content']
        if msgref_content is None:
          print('Warning: Encountered messageReference but content was null. Skipping.')
          print(json.dumps(attachment, indent=2))
          continue
        msgref_content = json.loads(msgref_content) # Convert JSON string into JSON object.
        msgref_preview = None if 'messagePreview' not in msgref_content else msgref_content['messagePreview']
        if msgref_preview is None:
          print('Warning: Encountered messageReference but content.messagePreview was null. Skipping.')
          print(json.dumps(attachment, indent=2))
          continue
        msgref_sender = msgref_content.get('messageSender', {}).get('user', {}).get('displayName', 'INVALID USER')

        # Build HTML for a nested conversation element. This should play nice with the convert to HTML function.
        msgref_html  = "<ul class=\"list-group\">\n"
        msgref_html += "\t<li class=\"list-group-item list-group-item-action d-flex justify-content-between align-items-start list-group-item-secondary\">\n"
        msgref_html += "\t\t<div class=\"me-auto\">\n"
        msgref_html += f"\t\t\t<div class=\"fw-bold\">{msgref_sender}</div>\n"
        msgref_html += f"\t\t\t<p>{msgref_preview}</p>\n"
        msgref_html += "\t\t</div>\n"
        msgref_html += "\t</li>\n"
        msgref_html += "</ul>\n"
        attachment_entry['html'] = msgref_html

      # Process code snippets.
      if attachment['contentType'] == 'application/vnd.microsoft.card.codesnippet':
        snip_content = attachment.get('content')
        if snip_content is None:
          print('Warning: Encountered code snippet but content was null. Skipping.')
          print(json.dumps(attachment, indent=2))
          continue
        snip_content = json.loads(snip_content) # Convert JSON string into JSON object.

        snip_url = snip_content.get('codeSnippetUrl')
        if snip_url is None:
          print('Warning: Encountered code snippet but codeSnippetUrl was null. Skipping.')
          print(json.dumps(attachment, indent=2))
          continue

        # Get the actual snippet data.
        snip_data = query_endpoint(oauth, snip_url, json=False)
        if snip_data is None:
          print('Warning: Failed to query snippet content. Skipping.')
          continue
        snip_code = snip_data.content.decode('utf-8') # Hopefully it's all utf-8 compatible!

        # Replace troublesome characters.
        snip_code = snip_code.replace('&', '&amp;') # NOTE: This MUST be firstg otherwise it will strip other fixes.
        snip_code = snip_code.replace('<', '&lt;')
        snip_code = snip_code.replace('>', '&gt;')
        snip_code = snip_code.replace('\"', '&quot;')
        snip_code = snip_code.replace('\'', '&apos;')
        # TODO: Handle more when I can be bothered (e.g., https://wonko.com/post/html-escaping)

        snip_html = f"<pre><code>\n{snip_code}</code></pre>\n"
        attachment_entry['html'] = snip_html
      # end if (code snippet)

      # Check for unhandled items (should have 'html' by now)
      if 'html' not in attachment_entry:
        print('WARNING: Unhandled attachment type.')
        print(json.dumps(attachment, indent=2))
        continue

      # Add this attachment to the list of all attachments and add a lookup entry into that based on its ID.
      # These values are used below and later when actually processing and downloading the attachments.
      all_attachments.append(attachment_entry)
      attachment_lookup[attachment_entry['id']] = len(all_attachments)-1
    #end for

    # Attachments are inserted with a custom <attachment> tag. This code replaces those tags accordingly.
    all_attachment_tags = re.findall(r"<attachment\s*.*?><\/attachment>", msg_entry['content'])
    for tag in all_attachment_tags:
      tag_id = re.findall(r"id=\"(.+?)\"", tag)[0]
      # We don't capture all attachments, so ignore those that don't have a lookup value.
      if tag_id not in attachment_lookup:
        continue

      # Lookup the actual attachment that we saved above based on the ID.
      attachment = all_attachments[attachment_lookup[tag_id]]
      # Get the new html from the attachment.
      attachment_html = attachment['html']
      # Replace old attachment HTML with the new one.
      msg_entry['content'] = msg_entry['content'].replace(tag, attachment_html)
    #end for

    all_messages.append(msg_entry)
  #end for

  print(f'  Total messages: {len(all_messages)}; Total attachments: {len(all_attachments)}; Total hosted content: {total_hosted_images}')

  print(f'    Writing attachments.json...', end='', flush=True)
  with open(root_folder + 'attachments.json', 'w', encoding="utf-8") as out_file:
    out_file.write(json.dumps(all_attachments, indent=2))
  print('done!')

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
    out_file.write(utils.json_to_html_chat(final_data))

  # Write out the extra files that are needed for the HTML to render and behave properly.
  utils.copy_extra_to_dst(root_folder)

  # Create a zip of the root folder.
  print('  Compressing data...', end='', flush=True)
  shutil.make_archive(f'static/files/{random_filename}', 'zip', root_folder)
  print('done!')

  print('Done!')

  return flask.send_file(f'static/files/{random_filename}' + '.zip', as_attachment=True, mimetype='application/octet-stream')

  # NOTE: I read online that this is safer and I shouldn't use `send_file`. Will look into it later.
  # return flask.send_from_directory('static/files', random_filename, as_attachment=True, max_age=0)
#end

if __name__ == '__main__':
  # Make sure the static/files directory definitely exists upon start.
  static_files_path = './static/files/'
  # https://stackoverflow.com/a/50901481
  # https://www.systutorials.com/in-python-os-makedirs-with-0777-mode-does-not-give-others-write-permission/
  old_umask = os.umask(0o000)
  os.makedirs(static_files_path, exist_ok=True)
  os.umask(old_umask)

  # Run the server.
  APP.run()
