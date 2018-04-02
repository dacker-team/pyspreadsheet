import argparse
# Path to get client_secrets.json and to store credentials
import httplib2
from googleapiclient.discovery import build
from oauth2client import client, file
from oauth2client import tools

from pyspreadsheet.path import credential_path


def get_api_account(project, version='v4'):
    google_client_secret_path, google_credentials_path = credential_path(project)
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    discovery_uri = 'https://sheets.googleapis.com/$discovery/rest'
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        parents=[tools.argparser])
    flags = parser.parse_args([])

    # Set up a Flow object to be used if we need to authenticate.
    flow = client.flow_from_clientsecrets(
        google_client_secret_path,
        scope=scopes,
        message=tools.message_if_missing(google_client_secret_path))

    path_storage = google_credentials_path + "/spreadsheet.json"
    storage = file.Storage(path_storage)
    credentials = storage.get()
    if credentials is None or credentials.invalid:
        credentials = tools.run_flow(flow, storage, flags)
    http = credentials.authorize(http=httplib2.Http())

    # Build the service object.
    if version == 'v3':
        account = build('analytics', version, http=http)
    else:
        account = build('analytics', version, http=http, discoveryServiceUrl=discovery_uri)

    return account.spreadsheets()
