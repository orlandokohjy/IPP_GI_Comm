# standard import
from datetime import date
import datetime
import pandas as pd
import numpy as np
import os
import glob
import PyPDF2
import docx


class Find_Cashbook:
    
    def __init__(self, folder_path, all_cb, previous_wk):
        self.folder_path = folder_path
        self.all_cb = all_cb
        self.previous_wk = previous_wk


    # merge with cashbook to find the cashbook and payment type
    def merge_cb(self, df):
        
        # sum the Commission with gst in a new column to find the cashbook
        df['Sum_Comm'] = np.round(df['Comm.Recd (with GST)'].sum(), 2)

        df = pd.merge(df, self.all_cb[['Extracted_Insurer', 'Num', 'Receipts \nChq No.', 'Debit']].rename(columns={'Num':'Cashbook ref. no.', 'Receipts \nChq No.':'Cheque/GIRO'}), how='left', left_on=['Insurer_Cashbook', 'Sum_Comm'],
                    right_on=['Extracted_Insurer', 'Debit'])
        
        return df
    

class GI_commission(Find_Cashbook):
    
    def __init__(self, folder_path, all_cb, previous_wk):
        super().__init__(folder_path, all_cb, previous_wk)
    
    # create a function for aig
    def run_aig_1(self, pattern):
        
        file_pattern = pattern + '*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            aig = []

        else:

            aig_read = []
            for i, j in enumerate(matching_files):
                aig_read.append(pd.read_excel(matching_files[i], sheet_name='Commission_Final', skipfooter=1))
                
            aig = pd.concat(aig_read)
                

            #aig = pd.read_excel(matching_files[0], sheet_name='Commission_Final', skipfooter=1)
            #aig_2 = pd.read_excel(matching_files[0], sheet_name='Sheet1', header=1, usecols=[1, 2], skipfooter=1)


            # name the GST Type column
            #aig_2 = aig_2.rename(columns={'Unnamed: 2': 'GST Type'})

            # merge both dfs
            #aig = pd.concat([aig_1, aig_2[['GST Type']]], axis=1)

            # remove "IPPFA - " from the ADVISER column (so that it could improve the FAR names matching algorithms)
            aig['ADVISER'] = aig['ADVISER'].str.replace('IPPFA - ', '')

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST Type']
            columns = ['ADVISER', 'POLICY/ENDT', 'POLICY EFF DATE', 'DESCRIPTION/PARTICULARS', 'COMM AMT']

            rename_col = ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)']

            label = dict(zip(columns, rename_col))

            aig = aig.rename(columns=label)

            # reformat date
            aig['Pol Date'] = pd.to_datetime(aig['Pol Date'])

            # create "Insurer" column
            aig['Insurer'] = 'AIG-GI'
            
            # a new column to merge with cashbook
            aig['Insurer_Cashbook'] = 'aig'
            
            # merge with cashbook to find the cashbook and payment type
            aig = super().merge_cb(aig)
        
        return aig



    # create a function for aia

    def run_aia_1(self, pattern):
        
        file_pattern = pattern + '*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            aia = []

        else:
            
            aia_read = []
            for i, j in enumerate(matching_files):
                aia_read.append(pd.read_excel(matching_files[0], header=2, skipfooter=3))
                
            aia = pd.concat(aia_read)

            #aia = pd.read_excel(matching_files[0], header=2, skipfooter=3)

            # forward fill the NaN with the previous valid observation
            aia = aia.fillna(method='ffill')

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST Type']
            columns = ['POLICY NAME', 'polno', 'Sum of TOTAL AMOUNT', 'TRANDTE']

            rename_col = ['Insured Name', 'Policy no.', 'Comm.Recd (with GST)', 'Pol Date']

            label = dict(zip(columns, rename_col))

            aia = aia.rename(columns=label)

            # reformat date
            aia['Pol Date'] = pd.to_datetime(aia['Pol Date'])

            # create "Insurer" column
            aia['Insurer'] = 'AIA-GI'
            
            # KJY comments
            aia['JY_comment'] = 'Unable to find TFAR even from all the sheets in the working file'
            
            # a new column to merge with cashbook
            aia['Insurer_Cashbook'] = 'aia'
            
            # merge with cashbook to find the cashbook and payment type
            aia = super().merge_cb(aia)
            
        return aia



    # create a function for Allianz

    def run_allianz_1(self, pattern):
        file_pattern = pattern + '*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            allianz = []

        else:

            allianz_read = []
            for i, j in enumerate(matching_files):
                allianz_read.append(pd.read_excel(matching_files[0], skipfooter=1))
                    
            allianz = pd.concat(allianz_read)

            #allianz = pd.read_excel(matching_files[0], skipfooter=1)

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['Policy Number', 'Agent Name', 'Policy Holder Name', 'Effective Date', 'Total Commission']

            rename_col = ['Policy no.', 'TFAR', 'Insured Name', 'Pol Date', 'Comm.Recd (with GST)']

            label = dict(zip(columns, rename_col))

            allianz = allianz.rename(columns=label)

            # reformat date
            allianz['Pol Date'] = pd.to_datetime(allianz['Pol Date'])

            # create "Insurer" column
            allianz['Insurer'] = 'ALLIANZ-GI'
            
            # a new column to merge with cashbook
            allianz['Insurer_Cashbook'] = 'allianz'
            
            # merge with cashbook to find the cashbook and payment type
            allianz = super().merge_cb(allianz)
        
        return allianz



    # create a function for Allied

    def run_allied_1(self, pattern):
        file_pattern = pattern + '*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            allied = []

        else:
            allied_read = []
            for i, j in enumerate(matching_files):
                allied_read.append(pd.read_excel(matching_files[0], header=6, skipfooter=13))
                    
            allied = pd.concat(allied_read)

            #allied = pd.read_excel(matching_files[0], header=6, skipfooter=13)

            # remove NaN in Policy No.
            allied = allied.dropna(subset=['Commission'])

            # forward fill NaN
            allied.loc[:, ['Account No.', 'Account Name', 'Currency', 'Payable']] = allied.loc[:, ['Account No.', 'Account Name', 'Currency', 'Payable']].fillna(method='ffill')

            # remove 'Unnamed' columns
            allied = allied.loc[:, ~allied.columns.str.contains('Unnamed')]

            # remove those payable as "N" entries
            allied = allied[allied['Payable'] == 'Y']

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['Policy No.', 'Reference', 'Commission GST', 'Commencement Date']

            rename_col = ['Policy no.', 'Insured Name', 'GST amt', 'Pol Date']

            label = dict(zip(columns, rename_col))

            allied = allied.rename(columns=label)

            # reformat date
            allied['Pol Date'] = pd.to_datetime(allied['Pol Date'])

            # create 'Insurer' column
            allied['Insurer'] = 'ALLIED WORLD-GI'
            allied['Comm.Recd (with GST)'] = allied[['Commission', 'GST amt']].sum(axis=1)*(-1)
            
            
            # remove the row without Policy no.
            allied = allied.dropna(subset=['Policy no.'])
            
            # merge the allied with the P0.Working file to get the TFAR based on policy number
            allied = pd.merge(allied, self.previous_wk[['Policy No', 'TFAR', 'Insured ']].rename(columns={'Policy No':'Policy no.', 'Insured ':'Insured Name'}), how='left', on='Policy no.')
            
            # insert KJY comment
            allied['JY_comment'] = np.where(allied['TFAR'].isna(), 'Unable to find TFAR even from the working files', np.NaN)
            
            # a new column to merge with cashbook
            allied['Insurer_Cashbook'] = 'allied'
            
            # merge with cashbook to find the cashbook and payment type
            allied = super().merge_cb(allied)
        
        return allied



    # create a function for Allied

    def run_allied_2(self, pattern):
        
        file_pattern = pattern + '*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            allied = []

        else:
            allied_read = []
            for i, j in enumerate(matching_files):
                allied_read.append(pd.read_excel(matching_files[0], header=6))
                    
            allied = pd.concat(allied_read)

            #allied = pd.read_excel(matching_files[0], header=6, skipfooter=13)

            # remove NaN in Policy No.
            allied = allied.dropna(subset=['Our Ref'])

            # remove 'Unnamed' columns
            allied = allied.loc[:, ~allied.columns.str.contains('Unnamed')]

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['Our Ref', 'Assured', 'Balance / Unallocated SGD', 'Brokerage GST']

            rename_col = ['Policy no.', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']

            label = dict(zip(columns, rename_col))

            allied = allied.rename(columns=label)


            # create 'Insurer' column
            allied['Insurer'] = 'ALLIED WORLD-GI'
            
            # change the sign for Comm.Recd (with GST)
            allied['Comm.Recd (with GST)'] = allied['Comm.Recd (with GST)'] * (-1)
            
            
            # merge the allied with the P0.Working file to get the TFAR based on policy number
            allied = pd.merge(allied, self.previous_wk[['Policy No', 'TFAR', 'Insured ']].rename(columns={'Policy No':'Policy no.', 'Insured ':'Insured Name'}), how='left', on='Policy no.')
            
            # insert KJY comment
            allied['JY_comment'] = np.where(allied['TFAR'].isna(), 'Unable to find TFAR even from the working files', np.NaN)
            
            # a new column to merge with cashbook
            allied['Insurer_Cashbook'] = 'allied'
            
            # merge with cashbook to find the cashbook and payment type
            allied = super().merge_cb(allied)
        
        return allied



    # create a function for Chubb

    def run_chubb_1(self, pattern):
        
        file_pattern = pattern + '*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            chubb =[]

        else:
            
            chubb_read = []
            for i, j in enumerate(matching_files):
                chubb_read.append(pd.read_excel(matching_files[0], header=7, skipfooter=1))
                    
            chubb = pd.concat(chubb_read)

            #chubb = pd.read_excel(matching_files[0], header=7, skipfooter=1)

            # remove 'Unnamed' columns
            chubb = chubb.loc[:, ~chubb.columns.str.contains('Unnamed')]

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['Policy No.', 'Agent', 'Comm', 'Issuance Date']

            rename_col = ['Policy no.', 'TFAR', 'Comm.Recd (with GST)', 'Pol Date']

            label = dict(zip(columns, rename_col))

            chubb = chubb.rename(columns=label)

            # reformat date
            chubb['Pol Date'] = pd.to_datetime(chubb['Pol Date'])

            # create 'Insurer' column
            chubb['Insurer'] = 'CHUBB-GI'
            
            # a new column to merge with cashbook
            chubb['Insurer_Cashbook'] = 'chubb'
            
            # merge with cashbook to find the cashbook and payment type
            chubb = super().merge_cb(chubb)
            
        return chubb



    # create a function for Ergo

    def run_ergo_1(self, pattern):

        file_pattern = pattern + '*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            ergo = []

        else:
            
            ergo_read = []
            for i, j in enumerate(matching_files):
                ergo_read.append(pd.read_excel(matching_files[0], header=3))
                    
            ergo = pd.concat(ergo_read)

            #chubb = pd.read_excel(matching_files[0], header=7, skipfooter=1)

            # remove 'Unnamed' columns
            ergo = ergo.loc[:, ~ergo.columns.str.contains('Unnamed')]

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['Policy No.', 'Total Commission']

            rename_col = ['Policy no.', 'Comm.Recd (with GST)']

            label = dict(zip(columns, rename_col))

            ergo = ergo.rename(columns=label)

            # create 'Insurer' column
            ergo['Insurer'] = 'ERGO-GI'
            
            # a new column to merge with cashbook
            ergo['Insurer_Cashbook'] = 'ergo'
            
            # merge with cashbook to find the cashbook and payment type
            ergo = super().merge_cb(ergo)
            
        return ergo



    # create a function for FWD

    def run_fwd_1(self, pattern):
        file_pattern = pattern + '*.xls*'
        advisor_file_pattern = 'W5*.xls*'

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))
        advisor_matching_files = glob.glob(os.path.join(self.folder_path, advisor_file_pattern))

        if matching_files == []:
            fwd = []

        else:
            
            fwd_read = []
            for i, j in enumerate(matching_files):
                fwd_read.append(pd.read_excel(matching_files[0]))
                    
            fwd = pd.concat(fwd_read).drop_duplicates()

            fwd_adviser = pd.read_excel(advisor_matching_files[0])
            #fwd = pd.read_excel(matching_files[0])

            fwd['agent_id_number'] = pd.to_numeric(fwd['agent_id_number'], errors='coerce')

            # merge with fwd advisor code list for the adviser details
            fwd = pd.merge(fwd, fwd_adviser, how='left', left_on='agent_id_number', right_on='FWD Life code')

            # remove rows that are without policy_number
            fwd = fwd.dropna(subset=['policy_number'])
            
            # drop duplicates based on policy number, 
            fwd = fwd.drop_duplicates(subset=['policy_number', 'policy_status_description', '$ txn commission'])
            
            # sum the commission and gst amount
            fwd['Comm.Recd (with GST)'] = fwd['$ txn commission'] + fwd['$txn gst commission']

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['policy_number', 'policyholder_full_name', 'GI Advisers', 'policy effective date']

            rename_col = ['Policy no.', 'Insured Name', 'TFAR', 'Pol Date']

            label = dict(zip(columns, rename_col))

            fwd = fwd.rename(columns=label)

            # reformat date
            fwd['Pol Date'] = pd.to_datetime(fwd['Pol Date'])

            # create 'Insurer' column
            fwd['Insurer'] = 'FWD-GI'
            
            # a new column to merge with cashbook
            fwd['Insurer_Cashbook'] = 'fwd'
            
            # merge with cashbook to find the cashbook and payment type
            fwd = super().merge_cb(fwd)
            
        return fwd


    # create a function for GE

    def run_ge_1(self, pattern):
        file_pattern = pattern + '*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            ge = []

        else:
            
            currentdate = datetime.datetime.now()
            sheet_name = currentdate.strftime("%b %y").upper()
            
            ge_read = []
            for i, j in enumerate(matching_files):
                ge_read.append(pd.read_excel(matching_files[0], sheet_name=sheet_name, skipfooter=2))
                    
            ge = pd.concat(ge_read)

            #ge = pd.read_excel(matching_files[0], sheet_name='MAY 23', skipfooter=2)

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['Policy Number', 'Agent Name', 'Particulars', 'Total Net Amount in Accounting Currency', 'Transaction Date']

            rename_col = ['Policy no.', 'TFAR', 'Insured Name', 'Comm.Recd (with GST)', 'Pol Date']

            label = dict(zip(columns, rename_col))

            ge = ge.rename(columns=label)

            # reformat date
            ge['Pol Date'] = pd.to_datetime(ge['Pol Date'])

            # create 'Insurer' column
            ge['Insurer'] = 'GE-GI'
            
            # remove the row without Policy no.
            ge = ge.dropna(subset=['Policy no.'])
            
            # merge with cashbook to find the cashbook and payment type
            ge = super().merge_cb(ge)
            
        return ge

    # create a function for ge_life

    def run_ge_life_1(self, pattern):
        
        file_pattern = pattern + '*.xls*'

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            ge_life = []

        else:
            
            ge_life_read = []
            for i, j in enumerate(matching_files):
                ge_life_read.append(pd.read_excel(matching_files[0]))
                    
            ge_life = pd.concat(ge_life_read)

            #ge_life = pd.read_excel(matching_files[0])
            
            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['POLICY NO.', 'AGENT \nNAME', 'TOTAL\nCOMM', 'GST ON  COMM']

            rename_col = ['Policy no.', 'TFAR', 'Comm.Recd (with GST)', 'GST amt']

            label = dict(zip(columns, rename_col))

            ge_life = ge_life.rename(columns=label)
            
            # remove rows without policy no.
            ge_life = ge_life.dropna(subset=['Policy no.'])
            
            # create 'Insurer' column
            ge_life['Insurer'] = 'GELIFE-GI'
            
            # a new column to merge with cashbook
            ge_life['Insurer_Cashbook'] = 'great'
            
            # merge with cashbook to find the cashbook and payment type
            ge_life = super().merge_cb(ge_life)
            
        return ge_life


    # create a function for HLA

    def run_hla_1(self, pattern):

        file_pattern = pattern + '*.xls*'
        advisor_file_pattern = 'W6*.xls*'

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))
        advisor_matching_files = glob.glob(os.path.join(self.folder_path, advisor_file_pattern))

        if matching_files == []:
            hla = []

        else:
            
            hla_read = []
            for i, j in enumerate(matching_files):
                hla_read.append(pd.read_excel(matching_files[0], header=3))
                    
            hla = pd.concat(hla_read)
            
            hla_adviser = pd.read_excel(advisor_matching_files[0])

            # merge to get advisor details
            hla = pd.merge(hla, hla_adviser, how='left', left_on='Staff ID', right_on='Code')
            
            
            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['Policy No.', 'Name Of Insured /Description', 'Adviser', 'Amount Due', 'GST on Commission', 'Transaction Date']

            rename_col = ['Policy no.', 'Insured Name', 'TFAR', 'Comm.Recd (with GST)', 'GST amt', 'Pol Date']

            label = dict(zip(columns, rename_col))

            hla = hla.rename(columns=label)

            # remove rows that have no Policy no.
            hla = hla.dropna(subset=['Policy no.'])

            # create 'Insurer' column
            hla['Insurer'] = 'HLA-GI'
            
            # a new column to merge with cashbook
            hla['Insurer_Cashbook'] = 'hong'
            
            # merge with cashbook to find the cashbook and payment type
            hla = super().merge_cb(hla)
            
        return hla


    # create a function for HSBC

    def run_hsbc_1(self, pattern):
        
        file_pattern = pattern + '*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            hsbc_1 = []

        else:

            hsbc_1 = pd.read_excel(matching_files[0], skipfooter=1)
            
            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['POLNUM', 'COMM_LCEAMT', 'EFFECTDATE']

            rename_col = ['Policy no.', 'Comm.Recd (with GST)', 'Pol Date']

            label = dict(zip(columns, rename_col))

            hsbc_1 = hsbc_1.rename(columns=label)

            # reformat the date
            hsbc_1['Pol Date'] = pd.to_datetime(hsbc_1['Pol Date'], format='%Y%m%d')

            # create 'Insurer' column
            hsbc_1['Insurer'] = 'HSBC-GI'

            # create an Insured Name column using the given name and surname
            hsbc_1['Insured Name'] = hsbc_1['LSURNAME'] + ' ' + hsbc_1['LGIVNAME']
            
            # merge with working files to get TFAR
            hsbc_1 = pd.merge(hsbc_1, self.previous_wk[['Policy No', 'TFAR']], how='left', left_on='Policy no.', right_on='Policy No')

            # insert comment
            hsbc_1['JY_comment'] = np.where(hsbc_1['TFAR'].isna(), 'Unable to find the TFAR in the working file', np.NaN)
            
            # a new column to merge with cashbook
            hsbc_1['Insurer_Cashbook'] = 'hsbc'
            
            # merge with cashbook to find the cashbook and payment type
            hsbc_1 = super().merge_cb(hsbc_1)
            
        return hsbc_1

        
    # create a function for HSBC

    def run_hsbc_2(self, pattern):
        
        file_pattern = '20_02*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            hsbc_2 = []

        else:

            hsbc_2 = pd.read_excel(matching_files[0], sheet_name='Detailed Breakdown (Earned)', header=9, skipfooter=2)

            hsbc_2['TFAR'] = pd.read_excel(matching_files[0], sheet_name='Summary').at[7, 'Unnamed: 2']

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['Policy No.', 'Commission Amount', 'Client Name', 'Policy GST Amount', 'Effective Date']

            rename_col = ['Policy no.', 'Comm.Recd (with GST)', 'Insured Name', 'GST amt', 'Pol Date']

            label = dict(zip(columns, rename_col))

            hsbc_2 = hsbc_2.rename(columns=label)

            # create 'Insurer' column
            hsbc_2['Insurer'] = 'HSBC-GI'

            # reformat date
            hsbc_2['Pol Date'] = pd.to_datetime(hsbc_2['Pol Date'])
            
            # merge with working files to get TFAR
            hsbc_2 = pd.merge(hsbc_2, self.previous_wk[['Policy No', 'TFAR']], how='left', left_on='Policy no.', right_on='Policy No').rename(columns={'TFAR_x':'TFAR'})

            # get TFAR from TFAR_y
            hsbc_2['TFAR'] = np.where(hsbc_2['TFAR'].isna(), hsbc_2['TFAR_y'], hsbc_2['TFAR'])

            # insert comment
            hsbc_2['JY_comment'] = np.where(hsbc_2['TFAR'].isna(), 'Unable to find the TFAR in the working file', np.NaN)
            
            # a new column to merge with cashbook
            hsbc_2['Insurer_Cashbook'] = 'hsbc'
            
            # merge with cashbook to find the cashbook and payment type
            hsbc_2 = super().merge_cb(hsbc_2)
            
        return hsbc_2


    # create a function for Income

    def run_income_1(self, pattern):
        file_pattern = pattern + '*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            income = []

        else:
            income = pd.read_excel(matching_files[0], header=1)

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['PolicyNo', 'FAR Name', 'Insured/ProposerName', 'Comm. (Round)']

            rename_col = ['Policy no.', 'TFAR', 'Insured Name', 'Comm.Recd (with GST)']

            label = dict(zip(columns, rename_col))

            income = income.rename(columns=label)
            
            # reformat date
            income['Pol Date'] = pd.to_datetime(income['Pol Date'])

            # create 'Insurer' column
            income['Insurer'] = 'INCOME-GI'
            
            # merge the allied with the P0.Working file to get the TFAR based on policy number
            income = pd.merge(income, self.previous_wk[['Policy No', 'TFAR', 'Insured ']].rename(columns={'Policy No':'Policy no.', 'Insured ':'Insured Name'}), how='left', on='Policy no.', suffixes=['', '_y'])

            # fill missing values with the corresponding values from Insured Name_y and TFAR_y
            income['Insured Name'].fillna(income['Insured Name_y'], inplace=True)
            income['TFAR'].fillna(income['TFAR_y'], inplace=True)

            # remove rows without TFAR
            income = income.dropna(subset=['TFAR'])

            # insert comment
            income['JY_comment'] = np.where(income['TFAR'].isna(), 'Unable to find TFAR in the working file', np.NaN)
            
            
            # a new column to merge with cashbook
            income['Insurer_Cashbook'] = 'income'
            
            # merge with cashbook to find the cashbook and payment type
            income = super().merge_cb(income)
            
        return income

    # create a function for india int

    def run_india_1(self, pattern):
        file_pattern = pattern + '*.xls*'

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            india = []

        else:
            sheet_dict = pd.read_excel(matching_files[0], sheet_name=None)
            
            # find the first sheet name
            first_sheet_name = next(iter(sheet_dict))
            
            # remove the first sheet
            del sheet_dict[first_sheet_name]
            
            # concatenate all the other sheets
            india = pd.concat(sheet_dict.values(), ignore_index=True)
            
            # remove policy number that is NaN
            india = india.dropna(subset=['Policy No.'])
            
            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['Policy No.', 'Customer', 'Total Comm', 'Gst on Commission']

            rename_col = ['Policy no.', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']

            label = dict(zip(columns, rename_col))

            india = india.rename(columns=label)
            
            # reformat date
            #india['Pol Date'] = pd.to_datetime(india['Pol Date'])
            
            # create 'Insurer' column
            india['Insurer'] = 'INDIA-GI'
            
            # merge the allied with the P0.Working file to get the TFAR based on policy number
            india = pd.merge(india, self.previous_wk[['Policy No', 'TFAR', 'Insured ']].rename(columns={'Policy No':'Policy no.', 'Insured ':'Insured Name'}), how='left', on='Policy no.', suffixes=['', '_y'])

            # fill missing values with the corresponding values from Insured Name_y and TFAR_y
            india['Insured Name'].fillna(india['Insured Name_y'], inplace=True)

            # insert comment
            india['JY_comment'] = np.where(india['TFAR'].isna(), 'Unable to find TFAR in the working file', np.NaN)
            
            # a new column to merge with cashbook
            india['Insurer_Cashbook'] = 'india'
            
            # merge with cashbook to find the cashbook and payment type
            india = super().merge_cb(india)
            
        return india

    # create a function for liberty

    def run_liberty_1(self, pattern):
        file_pattern = pattern + '*.csv*' 
        advisor_file_pattern = 'W7*.xls*'

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))
        advisor_matching_files = glob.glob(os.path.join(self.folder_path, advisor_file_pattern))

        if matching_files == []:
            liberty = []

        else:

            liberty_adviser = pd.read_excel(advisor_matching_files[0])
            liberty = pd.read_csv(matching_files[0])

            # merge with liberty adviser code list for the adviser details
            liberty = pd.merge(liberty, liberty_adviser, how='left', left_on='Sub Agent Code', right_on='CODE')

            # remove rows without policy number
            liberty = liberty.dropna(subset=['Policy/Renewal/Endorsement'])

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['Policy/Renewal/Endorsement', 'NAME OF ADVISER', 'Name of Insured', 'Total Commission Paid', 'Commission GST', 'Transaction Date']

            rename_col = ['Policy no.', 'TFAR', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt', 'Pol Date']

            label = dict(zip(columns, rename_col))

            liberty = liberty.rename(columns=label)

            # reformat date
            liberty['Pol Date'] = pd.to_datetime(liberty['Pol Date']) 

            # create 'Insurer' column
            liberty['Insurer'] = 'LIBERTY-GI'
            
            # a new column to merge with cashbook
            liberty['Insurer_Cashbook'] = 'liberty'
            
            # merge with cashbook to find the cashbook and payment type
            liberty = super().merge_cb(liberty)
            
        return liberty


    # create a function for msig

    def run_msig_1(self, pattern):
        file_pattern = '25*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            msig = []

        else:

            msig = pd.read_excel(matching_files[0], header=32)
            msig_soa = pd.read_excel(folder_path + 'MSIG SOA.xlsx', header=7)
            msig_rha = pd.read_excel(folder_path + 'MSIG-RHA231107.xlsx', sheet_name='Commission Statement - FA Firm', header=17)

            # drop rows where all column values are NaN
            msig = msig.dropna(how='all')

            # drop columns where all row values are NaN
            msig = msig.dropna(how='all', axis=1)

            # filter without 'Settlement Date' and with 'Unnamed: 27' - commission amount
            msig = msig[(msig['Settlement Date'].isna()) & (msig['Unnamed: 27'].notna())]

            # merge commission statement and SOA
            msig = pd.merge(msig, msig_soa[['Name of FA Rep', 'Policy Number']], how='left', left_on='Policy No\n', right_on='Policy Number')
            msig = pd.merge(msig, msig_rha[['Name of the Individual FA', 'Policy No']], how='left', left_on='Policy No\n', right_on='Policy No')

            # assign the "Unnamed: 18" as the "Insured Name"
            msig['Insured Name'] = msig['Unnamed: 18']

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['Policy No\n', 'Name of FA Rep', 'Unnamed: 37', 'Unnamed: 33', 'Effective Date']

            rename_col = ['Policy no.', 'TFAR', 'Comm.Recd (with GST)', 'GST amt', 'Pol Date']

            label = dict(zip(columns, rename_col))

            msig = msig.rename(columns=label)

            # reformat date
            msig['Pol Date'] = pd.to_datetime(msig['Pol Date'])


            # create 'Insurer' column
            msig['Insurer'] = 'MSIG-GI'
            
            # a new column to merge with cashbook
            msig['Insurer_Cashbook'] = 'msig'
            
            # merge with cashbook to find the cashbook and payment type
            msig = super().merge_cb(msig)
            
            
        return msig


    # create a function for qbe

    def run_qbe_1(self, pattern):
        file_pattern = pattern + '*.xls*'
        advisor_file_pattern = 'W8*.xls*'

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))
        advisor_matching_files = glob.glob(os.path.join(self.folder_path, advisor_file_pattern))

        if matching_files == []:
            qbe = []

        else:

            qbe_adviser = pd.read_excel(advisor_matching_files[0])
            qbe = pd.read_excel(matching_files[0], skipfooter=1)

            # remove whitespace in the agent code
            qbe_adviser['P400_USER'] = qbe_adviser['P400_USER'].str.strip()

            # merge with QBE adviser code list to get adviser details
            qbe = pd.merge(qbe, qbe_adviser, how='left', left_on='REP_NAME', right_on='P400_USER')

            # create a new column "Agent Name"
            #qbe['Agent Name'] = qbe['LASTNAME'] + ' ' + qbe['FIRSTNAME']
            
            # sum the commission and gst amount
            qbe['Comm.Recd (with GST)'] = qbe['COMMISSION_SGD'] + qbe['GST_ON_COMMISSION_SGD']

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['POLICY_NUMBER', 'CLIENT_NAME', 'GST_ON_COMMISSION_SGD', 'EFFECTDATE']

            rename_col = ['Policy no.', 'Insured Name', 'GST amt', 'Pol Date']

            label = dict(zip(columns, rename_col))

            qbe = qbe.rename(columns=label)

            # reformat date
            qbe['Pol Date'] = pd.to_datetime(qbe['Pol Date'], format='%Y%m%d')

            # create 'Insurer' column
            qbe['Insurer'] = 'QBE-GI'
            
            # a new column to merge with cashbook
            qbe['Insurer_Cashbook'] = 'qbe'
            
            # merge with cashbook to find the cashbook and payment type
            qbe = super().merge_cb(qbe)
            
            
        return qbe

    # create a function for singlife

    def run_singlife_1(self, pattern):
        file_pattern = pattern + '*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            singlife = []

        else:

            singlife = pd.read_excel(matching_files[0])

            # remove rows without ACCNUM
            singlife = singlife.dropna(subset=['ACCNUM'])

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['POLNUM', 'SRVAGNAME', 'PARTICULAR', 'Total', 'GST on Commission', 'TRAN_DATE']

            rename_col = ['Policy no.', 'TFAR', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt', 'Pol Date']

            label = dict(zip(columns, rename_col))

            singlife = singlife.rename(columns=label)

            # reformat date
            singlife['Pol Date'] = pd.to_datetime(singlife['Pol Date'])

            # create 'Insurer' column
            singlife['Insurer'] = 'SINGLIFE-GI'
            
            # remove row without Policy no.
            singlife = singlife.dropna(subset=['Policy no.'])
            
            # a new column to merge with cashbook
            singlife['Insurer_Cashbook'] = 'singapore'
            
            # merge with cashbook to find the cashbook and payment type
            singlife = super().merge_cb(singlife)
            
        return singlife


    # create a function for sompo

    def run_sompo_1(self, pattern):
        file_pattern = pattern + '*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            sompo = []

        else:

            sompo = pd.read_excel(matching_files[0], header=9)

            # remove columns with all the row values are NaN
            sompo = sompo.dropna(how='all', axis=1)

            # remove rows without policy number
            sompo = sompo.dropna(subset=['Policy No.'])

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['Policy No.', 'Producer Name', 'Insured Name & Vehicle No.', 'Commisison', 'GST Comm.']

            rename_col = ['Policy no.', 'TFAR', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']

            label = dict(zip(columns, rename_col))

            sompo = sompo.rename(columns=label)

            # create 'Insurer' column
            sompo['Insurer'] = 'SOMPO-GI'
            
            # remove rows where the commission amount is not float
            sompo['Comm.Recd (with GST)'] = pd.to_numeric(sompo['Comm.Recd (with GST)'], errors='coerce')

            sompo = sompo.dropna(subset=['Comm.Recd (with GST)'])
            
            # a new column to merge with cashbook
            sompo['Insurer_Cashbook'] = 'sompo'
            
            # merge with cashbook to find the cashbook and payment type
            sompo = super().merge_cb(sompo)
            
        return sompo


    # create a function for NHI

    def run_nhi_1(self, pattern):
        file_pattern = pattern + '*.docx' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        if matching_files == []:
            nhi = []
            
        else:
            
            doc = docx.Document(matching_files[0])
            
            table = doc.tables[2]

            table_data = []
            for row in table.rows:
                row_data = [cell.text for cell in row.cells]
                table_data.append(row_data)

            # make it as a df
            nhi = pd.DataFrame(table_data)

            # get the column as the first row values
            nhi.columns = nhi.iloc[0]

            # remove the first row 
            nhi = nhi.iloc[1:]

            # rename the columns as per the working file ['TFAR', 'Policy no.', 'Pol Date', 'Insured Name', 'Comm.Recd (with GST)', 'GST amt']
            columns = ['POLICY NO', 'SUB AGENT', 'INSURED NAME', 'COMM AMT (INC VAT)', 'POLICY START DATE']

            rename_col = ['Policy no.', 'TFAR', 'Insured Name', 'Comm.Recd (with GST)', 'Pol Date']

            label = dict(zip(columns, rename_col))

            nhi = nhi.rename(columns=label)
            
            # sum commission and gst amount
            nhi['Comm.Recd (with GST)'] = nhi['COMM AMT'].astype(float) + nhi['TAX AMT'].astype(float)
            #nhi['Comm.Recd (with GST)'] = nhi['Comm.Recd (with GST)'].astype(float)
            
            # create 'Insurer' column
            nhi['Insurer'] = 'NHI-GI'

            # create GST amount column
            nhi['GST amt'] = nhi['PREM PAID'].str.replace(',', '').astype(float) * nhi['GST %'].str.replace('%', '').astype(float) * 0.01
            
            # a new column to merge with cashbook
            nhi['Insurer_Cashbook'] = 'now'
            
            # merge with cashbook to find the cashbook and payment type
            nhi = super().merge_cb(nhi)

        return nhi