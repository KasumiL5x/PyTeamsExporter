supported_image_types = {
  'image/png': '.png',
  'image/jpeg': '.jpg',
  'image/gif': '.gif',
  'image/svg+xml': '.svg',
  'image/webp': '.webp'
}

def img_mime_to_ext(mime):
  return supported_image_types.get(mime)

def img_ext_to_mime(ext):
  res = [k for k,v in supported_image_types.items() if v == ext]
  return None if not len(res) else res[0]

supported_video_types = {
  'video/x-flv': '.flv',
  'video/mp4': '.mp4',
  'video/mpeg': '.m1v',
  'video/ogg': '.ogg',
  'video/webm': '.webm',
  'video/x-m4v': '.m4v',
  'video/quicktime': '.mov',
  'video/x-msvideo': '.avi',
  'video/x-ms-wmv': 'wmv'
}

def video_mime_to_ext(mime):
  return supported_video_types.get(mime)

def video_ext_to_mime(ext):
  res = [k for k,v in supported_video_types.items() if v == ext]
  return None if not len(res) else res[0]

def get_custom_css():
  """Returns custom CSS used in the pretty HTML chat page."""

  return """
body {
  background-color: #DBDBD4;
}
.speech-wrapper {
	background-color: #E5DDD5;
  padding: 30px 40px;
  display: flex;
  flex-direction: column;
  align-items: flex-start;
}
.speech-wrapper .bubble {
  height: auto;
  display: inline-block;
  background: #f5f5f5;
  border-radius: 4px;
  box-shadow: 2px 8px 5px rgba(0, 0, 0, 0.15);
  position: relative;
  margin: 10px 0 20px 25px;
	min-width: 350px;
	/* max-width: 80%; */
}
.speech-wrapper .bubble.alt {
  margin: 10px 25px 20px 0;
	margin-left: auto;
}
.speech-wrapper .bubble.continue {
  margin: 0 0 20px 60px;
}
.speech-wrapper .bubble .txt {
  padding: 8px 55px 8px 14px;
}
.speech-wrapper .bubble .txt .name {
  font-weight: 600;
  font-size: 1rem;
  margin: 0 0 4px;
  color: #3498db;
}
.speech-wrapper .bubble .txt .name span {
  font-weight: normal;
  color: #b3b3b3;
}
.speech-wrapper .bubble .txt .name.alt {
  color: #2ecc71;
}
.speech-wrapper .bubble .txt .message {
  font-size: 0.9rem;
  margin: 0;
  color: #2b2b2b;
  padding-right: 10px;
}
.speech-wrapper .bubble .txt .timestamp {
  font-size: 0.9rem;
  position: absolute;
  top: 8px;
  right: 10px;
  text-transform: uppercase;
  color: #999;
}
.speech-wrapper .bubble .bubble-arrow {
  position: absolute;
  width: 0;
  bottom: 42px;
  left: -16px;
  height: 0;
}
.speech-wrapper .bubble .bubble-arrow.alt {
  right: -2px;
  bottom: 40px;
  left: auto;
}
.speech-wrapper .bubble .bubble-arrow:after {
  content: "";
  position: absolute;
  border: 0 solid transparent;
  border-top: 9px solid #f5f5f5;
  border-radius: 0 20px 0;
  width: 15px;
  height: 30px;
  transform: rotate(145deg);
}
.speech-wrapper .bubble .bubble-arrow.alt:after {
  transform: rotate(45deg) scaleY(-1);
}
"""
#end

def json_to_html_chat(data):
  """
  Takes a data array and converts it into a pretty HTML document.
  Expected format of data is:
  {
    'chat': {
      'topic': str,
      'type': str,
      'members': [str, str, ...],
      'when': str,
      'link': str,
      'messages': [
        {
          'from': str,
          'content': str,
          'when': str
        },
        ...
      ]
    }
  }
  """

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

  # Custom CSS for chat style. Inspired by https://codepen.io/8eni/pen/YWoRGm.
  html_string += "<style>"
  html_string += get_custom_css()
  html_string += "</style>"

  html_string += "<body>\n"

  # Wrapper.
  html_string += "<div class=\"container text-break\">\n"
  # Navbar with title and details button.
  html_string += "\t<nav class=\"navbar navbar navbar-dark shadow-sm\" style=\"background-color: #00BFA5;\">\n"
  html_string += "\t\t<div class=\"container-fluid\">\n"
  html_string += f"\t\t\t<span class=\"navbar-brand mb-0 h1\">{header_text}</span>\n"
  html_string += "\t\t\t<button class=\"btn btn-light\" type=\"button\" data-bs-toggle=\"collapse\" data-bs-target=\"#navbarToggleExternalContent\" aria-controls=\"navbarToggleExternalContent\" aria-expanded=\"false\" aria-label=\"See details\">\n"
  html_string += "\t\t\t\tSee details\n"
  html_string += "\t\t\t</button>\n"
  html_string += "\t</nav>\n"
  # The hidden details panel.
  html_string += "\t<div class=\"collapse\" id=\"navbarToggleExternalContent\">\n"
  html_string += "\t\t<div class=\"p-4\" style=\"background-color: #2A2F32;\">\n"
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

  # Messages (v2).
  html_string += "\t<div class=\"speech-wrapper\">\n"
  previous_from = ''
  use_alt = True
  for msg in messages_data:
    msg_from = msg['from']
    msg_content = msg['content']
    msg_timestamp = msg['when']

    if previous_from != msg_from:
      use_alt = not use_alt
      previous_from = msg_from

    alt_str = ' alt' if use_alt else ''

    html_string += f"\t\t<div class=\"bubble{alt_str}\">\n"
    html_string += "\t\t\t<div class=\"txt\">\n"
    html_string += f"\t\t\t\t<div class=\"name{alt_str}\">{msg_from}</div>\n"
    html_string += "\t\t\t\t<div class=\"message\">\n"
    html_string += msg_content
    html_string += "\t\t\t\t</div>\n"
    html_string += f"<span class=\"timestamp\">{msg_timestamp}</span>"
    html_string += "\t\t\t</div>\n"
    html_string += f"<div class=\"bubble-arrow{alt_str}\"></div>"
    html_string += "\t\t</div>\n"
  #end
  html_string += "\t</div>\n"

  # # Messages.
  # html_string += "\t<ul class=\"list-group\">\n"
  # for msg in messages_data:
  #   msg_from = msg['from']
  #   msg_content = msg['content']
  #   msg_timestamp = msg['when']

  #   html_string += "\t\t<li class=\"list-group-item list-group-item-action d-flex justify-content-between align-items-start\">\n"
  #   html_string += "\t\t\t<div class=\"ms-2 me-auto\">\n"
  #   html_string += f"\t\t\t\t<div class=\"fw-bold\">{msg_from}</div>\n"
  #   html_string += f"\t\t\t\t{msg_content}\n"
  #   html_string += "\t\t\t</div>\n"
  #   html_string += f"\t\t\t<span class=\"badge bg-primary\">{msg_timestamp}</span>\n"
  #   html_string += "\t\t</li>\n"
  # #end
  # html_string += "\t</ul>\n\n"

  # Wrapper end.
  html_string += "</div>"

  # Footer.
  html_string += "\t<script src=\"https://cdn.jsdelivr.net/npm/bootstrap@5.1.1/dist/js/bootstrap.bundle.min.js\" integrity=\"sha384-/bQdsTh/da6pkI1MST/rWKFNjaCP5gBSY4sEBT38Q/9RBh9AH40zEOg7Hlq2THRZ\" crossorigin=\"anonymous\"></script>\n"
  html_string += "</body>\n"
  html_string += "</html>"

  return html_string
#end
