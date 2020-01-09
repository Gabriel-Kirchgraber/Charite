import pandas as pd
import ctypes
import os

def main():
    path = os.getcwd()
    import_file = 'animallist.csv'
    report_file = 'Test.xlsm'
    transfer_file = 'animallist.xlsx'
    
    if not os.path.isfile(os.path.join(path, import_file)):
        message_box('{} nicht in {} gefunden.'.format(import_file, path), 'Datei nicht gefunden')
        return
    if not os.path.isfile(os.path.join(path, report_file)):
        message_box('{} nicht in {} gefunden.'.format(report_file, path), 'Datei nicht gefunden')
        return
    if os.path.isfile(os.path.join(path, transfer_file)):
        os.remove(os.path.join(path, transfer_file))
    
    df = pd.read_csv(os.path.join(path, import_file), engine='python', sep='\t', encoding='utf_16')
    df = df.dropna(axis='columns', how='all')
    
    columns = list(df)
    
    for column in columns:
        if (df[column].dtypes == 'int64')|(df[column].dtypes == 'float64'):
            if (df[column].sum()) == 0:
                df.drop([column], axis='columns', inplace=True)
                
    df.to_excel(os.path.join(path, transfer_file), index=False)

def message_box(text, title):
    ctypes.windll.user32.MessageBoxW(0, text, title, 1)


if __name__== "__main__":
     main()
    

