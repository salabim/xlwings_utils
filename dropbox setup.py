"""
This Python script shows how to create a Dropbox app and generate the required environment variables.
"""

import sys

try:
    import dropbox
except ImportError:
    print("dropbox not installed. Use pip install dropbox to install")
    sys.exit()

print("""\
Instructions how to get the required DROPBOX.APP_KEY, DROPBOX.APP_SECRET and DROPBOX.REFRESH_TOKEN 
to access Dropbox in xlwings_utils or synchronista.

First, make an app by going to
    https://www.dropbox.com/developers/apps
There, click "Create app"
    Select "Full dropbox access"
    Name your app (name is not important, but it should be unique)
Click "Create"
In "Select scopes",
    Tick "files.metadata.write"
    Tick "files.content.write"
    Tick "files.content.read"
    Save
      
Under Settings you'll see the App key""")
APP_KEY = input("Enter App key ==> ").strip()
print("and you'll see the App secret (by clicking 'show')")
APP_SECRET = input("Enter App secret ==> ").strip()

auth_flow = dropbox.DropboxOAuth2FlowNoRedirect(
    consumer_key=APP_KEY,
    consumer_secret=APP_SECRET,
    token_access_type="offline",  # This is key to getting a refresh token
    scope=["files.metadata.read", "files.content.read", "files.content.write"],  # Add your needed scopes
)

authorize_url = auth_flow.start()
print("Go to: " + authorize_url)
print()
print("Click 'Allow' (you might have to log in first)")
print("and copy the shown autorization code.")
input("Enter when done")

auth_code = input("Enter the authorization code here ==> ").strip()
oauth_result = auth_flow.finish(auth_code)

ACCESS_TOKEN = oauth_result.access_token
REFRESH_TOKEN = oauth_result.refresh_token

print("Now set environment variables")
print("          DROPBOX.APP_KEY", APP_KEY)
print("       DROPBOX.APP_SECRET", APP_SECRET)
print("    DROPBOX.REFRESH_TOKEN", REFRESH_TOKEN)
print()
print("All done.")
