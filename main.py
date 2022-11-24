import pygsheets
import pandas as pd

gc = pygsheets.authorize(service_account_file='local/auth.json')

gsheetURL= 'https://docs.google.com/spreadsheets/d/1rUWc3EMay7j3_GJMFikF8Px3VoFgdcubBPRVxQOYxbw/edit?usp=sharing'
sh = gc.open_by_url(gsheetURL)
ws = sh.worksheet_by_title('main')
# write the sheet=====
# df1 = pd.DataFrame({'a': [1, 2], 'b': [3, 4]})
# ws.set_dataframe(df1, 'A1', copy_index=True, nan='')

# Read the file=====
def ws_to_df(self, **kwargs):
    worksheet_title = self.title
    self.export(filename=worksheet_title + '_df')
    df = pd.read_csv(worksheet_title + '_df.csv', **kwargs)
    return df
pygsheets.worksheet.Worksheet.ws_to_df = ws_to_df
df4 = ws.ws_to_df()
print(df4)
