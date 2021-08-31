"""
This module contains class AccessUnzip.
Class is designed to convert .accdb files to .xlsx and send it via SFTP.
"""
import os
import logging
import sys

import pandas as pd
import pyodbc


class AccessToExcel:

    def __init__(self, source_file, target_dir):
        self.source_file = source_file
        self.target_dir = target_dir
        logging.info(f"Initialized with: source_file: {source_file}, target_dir: {target_dir}")
    

    def __call__(self):
        logging.info("Starting the process of reading accdb file and converting to xlsx")
        self.write_df_to_xlsx()
    

    def _find_all_tables(self) -> pd.DataFrame:
        """Generator connect to accdb database, read all tables and convert it to Pandas DF"""

        driver = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)}'
        conn = pyodbc.connect(f'{driver};DBQ={self.source_file};')
        cursor = conn.cursor()
        tables = [table[2] for table in cursor.tables(tableType='TABLE')]
        logging.info(f"Tables found in file {self.source_file}: {tables}")

        for table in tables:

            try:
                df = pd.read_sql(f'select * from "{table}"', conn)
                df.name = table
                logging.info(f"Passing df: {df.name} to by write into Excel file.")
                yield df

            except pyodbc.ProgrammingError as error:
                logging.error(f'Cannot read table named: {table}')
    

    def write_df_to_xlsx(self):
        """Method uses _find_all_tables generator to generate DF's and load ot to xlsx file with multiple sheets"""

        self.file_path = os.path.join(self.target_dir, os.path.basename(self.source_file).replace('accdb', 'xlsx'))

        with pd.ExcelWriter(self.file_path) as target_file:
            logging.info(f"Created Excel file: {self.file_path}.")

            for df in self._find_all_tables():
                logging.info(f"Loading sheet: {df.name} into {self.file_path}")
                df.to_excel(target_file, sheet_name=df.name, index=False)


if __name__ == '__main__':
    kwargs = dict(arg.split('=') for arg in sys.argv[1:])

    if 'logs' in kwargs:
        logging.basicConfig(filename=kwargs.get('logs'), filemode='a', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        kwargs.pop('logs')

    obj = AccessToExcel(**kwargs)
    obj()
    logging.info(f"FINISHED")
