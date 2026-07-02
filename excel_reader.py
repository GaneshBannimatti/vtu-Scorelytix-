import pandas as pd

def read_usns_from_excel(filepath):

    df = pd.read_excel(filepath)

    usns = []

    for value in df.iloc[:, 0]:

        if pd.notna(value):
            usn = str(value).strip().upper()

            if usn:
                usns.append(usn)

    return usns