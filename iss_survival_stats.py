import os
import re
import pandas as pd
import openpyxl
from datetime import datetime

try:
    # Load the ISS table from excel
    iss_table = pd.read_excel("ISS.xlsx")
    print("ISS table loaded successfully.")

    # Load the patient survival data from survival tsv
    survival_data = pd.read_csv("MMRF_CoMMpass_IA22_STAND_ALONE_SURVIVAL.tsv",delimiter="\t",low_memory=False)
    print("Survival data loaded successfully")

    # Filter the data to get just the columns of interest and remove _ from the patient IDs
    survival_data = survival_data[['PUBLIC_ID','deathdy','linesdy2']]
    survival_data.columns = ['Patient_ID','deathdy','progressiondy']
    survival_data['Patient_ID'] = survival_data['Patient_ID'].apply(lambda x: re.sub(r'(MMRF)_(\d+)', r'\1\2', x))

    # Merge the survival data with the ISS table
    merged_data = pd.merge(iss_table, survival_data, on='Patient_ID', how='left')

    # Function to calculate survival statistics for ISS, R-ISS or R2-ISS
    def survival_stats(iss, category, data):
        iss_grouped = data.groupby(iss)
        iss_mean = iss_grouped[category].mean()
        iss_stdev = iss_grouped[category].std()
        iss_median = iss_grouped[category].median()
        iss_iqr = iss_grouped[category].quantile(0.75) - iss_grouped[category].quantile(0.25)
        return pd.DataFrame({'Mean': iss_mean, 'StDev': iss_stdev, 'Median': iss_median, 'IQR': iss_iqr})
    
    # Calculate survival statistics for ISS
    iss_os_stats = survival_stats('ISS', 'deathdy', merged_data)
    print("\nSurvival statistics for ISS:")
    print(iss_os_stats)

    # Calculate survival statistics for R-ISS
    r_iss_os_stats = survival_stats('R-ISS', 'deathdy', merged_data)
    print("\nSurvival statistics for R-ISS:")
    print(r_iss_os_stats)

    # Calculate survival statistics for R2-ISS
    r2_iss_os_stats = survival_stats('R2-ISS', 'deathdy', merged_data)
    print("\nSurvival statistics for R2-ISS:")
    print(r2_iss_os_stats)

    # Calculate progression statistics for ISS
    iss_pfs_stats = survival_stats('ISS', 'progressiondy', merged_data)
    print("\nProgression statistics for ISS:")
    print(iss_pfs_stats)

    # Calculate progression statistics for R-ISS
    r_iss_pfs_stats = survival_stats('R-ISS', 'progressiondy', merged_data)
    print("\nProgression statistics for R-ISS:")
    print(r_iss_pfs_stats)

    # Calculate progression statistics for R2-ISS
    r2_iss_pfs_stats = survival_stats('R2-ISS', 'progressiondy', merged_data)
    print("\nProgression statistics for R2-ISS:")
    print(r2_iss_pfs_stats)

    # Save the survival statistics to an excel file, with separate sheets for ISS, R-ISS and R2-ISS
    with pd.ExcelWriter('Survival_Statistics.xlsx') as writer:
        iss_os_stats.to_excel(writer, sheet_name='ISS OS')
        r_iss_os_stats.to_excel(writer, sheet_name='R-ISS OS')
        r2_iss_os_stats.to_excel(writer, sheet_name='R2-ISS OS')
        iss_pfs_stats.to_excel(writer, sheet_name='ISS PFS')
        r_iss_pfs_stats.to_excel(writer, sheet_name='R-ISS PFS')
        r2_iss_pfs_stats.to_excel(writer, sheet_name='R2-ISS PFS')

except Exception as e:
    print(f"An error occurred: {e}")