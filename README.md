A script that edits a google document table

# Setup

1. Enable Google Docs API
Goto [Google Python Quickstart](https://developers.google.com/docs/api/quickstart/python)
and press "Enable the Google Docs API". Save `credentials.json` to this directory.

2. Install Google Client Library
```bash
pip3.6 install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

# Running the script

Edit `DOCUMENT_ID`, `TARGETCELLTEXT`, and `TEXTTOINSERT` in `modifyGoogleDocTable.py`

`./modifyGoogleDocTable.py`
