from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
import gspread
import numpy as np
import pandas as pd

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents'
]

creds = ServiceAccountCredentials.from_json_keyfile_name(
    "./venv/creds.json", scopes=SCOPES)
docs = build("docs", "v1", credentials=creds)
gc = gspread.authorize(creds)
drive = build("drive", "v3", credentials=creds)


# https://docs.google.com/spreadsheets/d/1rK5BB9KUawSJMotOAmZIheB38QAGk5tln1pROFKk6UY/edit#gid=1791625530
gsheet = gc.open_by_key("1rK5BB9KUawSJMotOAmZIheB38QAGk5tln1pROFKk6UY")
data_sheet = gsheet.worksheet("Form Responses 1").get_all_values()
df = pd.DataFrame(data_sheet[1:], columns=data_sheet[0]).replace(
    r'^\s*$', np.nan, regex=True)
familias = df.loc[df["Tipe Keanggotaan"] == "Familia"].dropna(
    axis=1, how="all").to_dict("records")


# https://docs.google.com/document/d/1CJ-H8BFRcuw8qGJqpuobmAsmah8Ct6RbHu_fwPBAJj8/edit

sheet_familia = gsheet.worksheet("Familia")
familia_values = sheet_familia.get_all_values()
if (not familia_values):
    familia_values = [list(familias[0].keys()) + ["Merged", "Link Merge"]]
df_familia = pd.DataFrame(familia_values[1:], columns=familia_values[0])
for familia in familias:
    if not df_familia.loc[df_familia["Nama Familia"] == familia["Nama Familia"]].empty and df_familia.loc[df_familia["Nama"] == familia["Nama"]]["Merged"].values[0] == "TRUE":
        continue

    body = {
        'name': familia["Nama Familia"] + " - " + familia["Nama Ketua"],
        'parents': ['15Y2R_ljg_izilDk99RLF9jj9uhNl6xcU']
    }
    copy_response = drive.files().copy(
        fileId='1CJ-H8BFRcuw8qGJqpuobmAsmah8Ct6RbHu_fwPBAJj8',
        body=body
    ).execute()
    copy_id = copy_response.get('id')

    request = [
        {
            'replaceAllText': {
                'containsText': {
                    'text': '<<Nama Familia>>',
                    'matchCase': 'true'
                },
                'replaceText': familia["Nama Familia"]
            }
        }, {
            'replaceAllText': {
                'containsText': {
                    'text': '<<Anggota>>',
                    'matchCase': 'true'
                },
                'replaceText': "\n".join(familia["Anggota"].split(';'))
            }
        }
    ]
    response = docs.documents().batchUpdate(
        documentId=copy_id,
        body={
            'requests': request
        }
    ).execute()

    if df_familia.loc[df_familia["Nama Familia"] == familia["Nama Familia"]].empty:
        familia["Link Merge"] = "https://docs.google.com/document/d/" + \
            copy_id + "/edit"
        familia["Merged"] = True
        df_familia = df_familia.append(familia, ignore_index=True)
    else:
        df_familia.loc[df_familia["Nama Familia"] == familia["Nama Familia"]
                       ]["Link Merge"] = "https://docs.google.com/document/d/" + copy_id + "/edit"
        df_familia.loc[df_familia["Nama Familia"] ==
                       familia["Nama Familia"]]["Merged"] = True
sheet_familia.update([df_familia.columns.values.tolist()
                      ] + df_familia.values.tolist())
print("Familia Updated")
