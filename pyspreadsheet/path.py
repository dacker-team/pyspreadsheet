import os


def credential_path(project):
    google_client_secret_path = os.environ.get("GOOGLE_%s_CLIENT_SECRET_PATH" % project)
    google_credentials_path = os.environ.get("GOOGLE_%s_CREDENTIALS_PATH" % project)

    if not google_client_secret_path:
        google_client_secret_path = './'

    if not google_credentials_path:
        google_credentials_path = './'

    if google_client_secret_path[-1] == '/':
        google_client_secret_path = google_client_secret_path[:-1]

    if google_credentials_path[-1] == '/':
        google_credentials_path = google_credentials_path[:-1]

    google_client_secret_path = google_client_secret_path + '/client_secrets.json'
    return google_client_secret_path, google_credentials_path