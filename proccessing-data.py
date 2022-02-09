import requests
import scrapy
import os
import pandas as pd

def get_directories(url):
    """
        Scrapes the directories from the given url using scrapy lib.
    """
    response = requests.get(url)
    response.raise_for_status()
    response.encoding = 'utf-8'
    # Get <td><a> tags
    links = scrapy.Selector(response).xpath('//td/a')
    # Filter links to only directories excluding Parent Directory
    directories = [link.xpath('text()').extract_first() for link in links if link.xpath('text()').extract_first() != 'Parent Directory']
    return directories

def download_file(url, file_name, destination):
    """
        Downloads the file from the given url and saves it to the given destination.
    """
    response = requests.get(url)
    response.raise_for_status()
    # Save file to destination
    with open(destination + file_name, 'wb') as file:
        file.write(response.content)

def download_get_salaries_files_from_duran():
    """
        Downloads the salaries files from Duran's web.
    """
    months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
    url = 'https://duran.gob.ec/wp-content/uploads/transparencia/Lotaip/'
    destination = './data/'
    # Exception file name list
    dictFileName = {'default':      'C%20-%20Remuneracion%20mensual%20por%20puesto.pdf',
                    '2021_enero.pdf':   'C%20Remuneracion%20mensual%20por%20puesto.pdf',
                    '2021_febrero.pdf': 'C%20Remuneracion%20mensual%20por%20puesto.pdf',
                    '2021_junio.pdf':   'C%20-%20Remuneraci%c3%b3n%20mensual%20por%20puesto.pdf',
                    '2021_julio.pdf':   'C%20-%20Remuneración%20mensual%20por%20puesto.pdf',
                    '2021_octubre.pdf':  'C%20-%20Remuneración%20menusal%20por%20puesto.pdf',
                    '2021_diciembre.pdf':'C%20-%20Remuneración%20mensual%20por%20puesto.pdf',
                    '2022_enero.pdf':   'C%20-%20Remuneración%20mensual%20por%20puesto.pdf',}

    # Get directories
    directories = get_directories(url)
    # Download files
    for directory in directories:
        for month in months:
            # Get year from name of directory (e.g. '2018')
            year = directory[:4]
            file_name = year + '_' + month + '.pdf'
            # If file name is in key exception list, use the value
            # Otherwise, use default value
            if file_name in dictFileName:
                file_name_source = dictFileName[file_name]
            else:
                file_name_source = dictFileName['default'] 

            # if not exists, download file
            if not os.path.isfile(destination + file_name):
                try:
                    download_file(url + directory + month + '/' + file_name_source, file_name, destination)
                except Exception as e:
                    print('Error downloading file: ' + file_name)
                    print(e)

def get_df_to_excel(path, file_name):
    """
        Reads the given file and returns a dataframe.
    """
    print('Reading file: ' + file_name)

    files_with_id = ['2015_julio.xlsx', '2015_mayo.xlsx']
    irregular_files = ['2015_junio.xlsx', '2016_febrero.xlsx']

    # Get year and month_name from file name. Split by '_'
    # e.g. '2018_enero.xlsx' -> ['2018', 'enero']
    year, month_name = file_name.split('_')
    month_name = month_name.split('.')[0] 

    # List of dataframes
    df_list = []
    # Get sheet names from excel file
    sheet_names = pd.ExcelFile(path + file_name).sheet_names
    # Get dataframe from each sheet without column names
    for sheet_name in sheet_names:
        print('  Reading Sheet: ' + sheet_name)
        # If sheet name is 'Table 1', skip 4 first rows
        if sheet_name == 'Table 1':
            df = pd.read_excel(path + file_name, sheet_name, skiprows=3)
        # If sheet name is 'Table 2', skip 2 first rows
        elif sheet_name == 'Table 2':
            df = pd.read_excel(path + file_name, sheet_name, skiprows=1)
        else:
            df = pd.read_excel(path + file_name, sheet_name)

        if file_name in irregular_files:
            if sheet_name == 'Table 2':
                # Drop last 3 rows
                df = df.drop(df.index[-3:])
        else:
            # If sheet is last sheet, drop 7 lastest rows
            if sheet_name == sheet_names[-1]:
                df = df.drop(df.index[-7:])

        if file_name not in files_with_id:
            # Add column after firts column with null values
            df.insert(1, 'id', None)

        # Add column 'year' and 'month_name'
        df['year'] = year
        df['month_name'] = month_name
        
        # Add dataframe to list
        df_list.append(df)
        # Return list of dataframes
    return df_list

def unified_excel_files(path):
    """
        Generate salaries.csv file from all excel files in the given path.
    """
    PATH_CSV_FILE = './input/salaries.csv'
    # Drop existing csv file
    if os.path.isfile(PATH_CSV_FILE):
        os.remove(PATH_CSV_FILE)

    # Get list of excel files
    excel_files = [file for file in os.listdir(path) if file.endswith('.xlsx')]
    for file in excel_files:
        # Get dataframes from excel files
        df_list = get_df_to_excel(path, file)
        for df in df_list:
            if os.path.isfile(PATH_CSV_FILE):
                df.to_csv( PATH_CSV_FILE, mode='a', header=False, index=False)
            else:
                df.to_csv(PATH_CSV_FILE, header=False, index=False)
    


if __name__ == '__main__':
    print('Downloading files from Duran\'s web...')
    download_get_salaries_files_from_duran()

    unified_excel_files('./data/')        
    
    print('Done!')
