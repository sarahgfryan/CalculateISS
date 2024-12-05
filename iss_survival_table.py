import os
import re
import pandas as pd
import openpyxl
from datetime import datetime

try:
    # Load the ISS table from excel
    iss_table = pd.read_excel("ISS.xlsx")
    print("ISS table loaded successfully.")

    # Remove irrelevant ISS columns
    iss_table = iss_table.drop(columns=['ldh', 'alb', 'b2m', 't(4;14)', 't(14;16)', 'gain1q', 'del17p'])

    # Change numbers for ISS, R-ISS, R2-ISS into Roman numeral strings, if no value present change to NaN
 #   iss_table['ISS'] = iss_table['ISS'].apply(lambda x: 'I' if x == 1 else ('II' if x == 2 else ('III' if x == 3 else None)))
 #   iss_table['R-ISS'] = iss_table['R-ISS'].apply(lambda x: 'I' if x == 1 else ('II' if x == 2 else ('III' if x == 3 else None)))
 #   iss_table['R2-ISS'] = iss_table['R2-ISS'].apply(lambda x: 'I' if x == 1 else ('II' if x == 2 else ('III' if x == 3 else ('IV' if x == 4 else None))))

    # Rename R-ISS and R2-ISS to RISS and R2ISS respectively
    iss_table.rename(columns={'R-ISS': 'RISS', 'R2-ISS': 'R2ISS'}, inplace=True)
    
    # Load the patient survival data from survival tsv
    survival_data = pd.read_csv("MMRF_CoMMpass_IA22_STAND_ALONE_SURVIVAL.tsv",delimiter="\t",low_memory=False)
    print("Survival data loaded successfully")

    # Pull out desired columns and rename them
    surv_col = ['PUBLIC_ID','censos','ttcos','censpfs','ttcpfs']
    survival_data = survival_data[surv_col]
    survival_data.columns = ['Patient_ID','death','os_time','progression','pfs_time']

    # Merge the survival data with the ISS table
    merged_data = pd.merge(iss_table, survival_data, on='Patient_ID', how='left')

    # Save the merged data to a new excel file
    merged_data.to_excel("Survival_ISS.xlsx", index=False)
    print("Survival and ISS merged data saved successfully.")

except Exception as e:
    print(f"An error occurred: {e}")