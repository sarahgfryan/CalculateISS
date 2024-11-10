import os
import re
import pandas as pd
import openpyxl
from datetime import datetime

try:
    # Load the patient visit data
    patient_visit_data = pd.read_csv("MMRF_CoMMpass_IA22_PER_PATIENT_VISIT.tsv",delimiter="\t",low_memory=False)
    print("Visit data loaded successfully")

    # Filter the data to get just the baseline visit data
    baseline_data = patient_visit_data[patient_visit_data["VJ_INTERVAL"] == "Baseline"]

    # Filter the data to get just the columns of interest to calculate ISS, R-ISS and R2-ISS
    lab_columns = ['PUBLIC_ID', 'D_LAB_serum_beta2_microglobulin','D_LAB_chem_albumin','D_LAB_chem_ldh']
    baseline_data = baseline_data[lab_columns]
    baseline_data.columns = ['Patient_ID', 'b2m', 'alb', 'ldh']

    # Drop duplicates for baseline visits
    baseline_data = baseline_data.drop_duplicates(subset='Patient_ID', keep='first')

    # Load the cytogenetic data
    cytogenetic_data = pd.read_csv("Cytogenetics_Baseline.txt", delimiter="\t")
    print("Cytogenetic data loaded successfully")

    # Filter the cytogenetic data to get just the columns of interest
    cytogenetic_data = cytogenetic_data[['Patient_ID','t(4;14) - WHSC1','t(14;16) - MAF','Gain1q21','Del17p13']]
    # Rename columns to merge with the baseline data
    cytogenetic_data.columns = ['Patient_ID', 't(4;14)', 't(14;16)', 'gain1q', 'del17p']
    # Add the underscore in Patient_ID to match the baseline data
    cytogenetic_data['Patient_ID'] = cytogenetic_data['Patient_ID'].apply(lambda x: re.sub(r'(MMRF)(\d+)', r'\1_\2', x))

    # Load the Compass provided ISS data
    iss_data = pd.read_csv("Compass_ISS.txt", delimiter="\t")
    print("Compass ISS data loaded successfully")
    iss_data['ISS_Compass'] = None
    iss_data.columns = ['Patient_ID', 'ISS_str',"ISS_Compass"]

    # Function to convert ISS string to number
    def iss_str_to_num(row):
        if row['ISS_str'] == 'Stage I':
            return 1
        elif row['ISS_str'] == 'Stage II':
            return 2
        elif row['ISS_str'] == 'Stage III':
            return 3
        else:
            return float('nan')
    
    # Apply the function to convert ISS string to number
    iss_data['ISS_Compass'] = iss_data.apply(iss_str_to_num, axis=1)
    del iss_data['ISS_str']

    # Merge labs, cytogenetics and ISS data
    merged_data = pd.merge(baseline_data, cytogenetic_data, on='Patient_ID', how='left')
    merged_data = pd.merge(merged_data, iss_data, on='Patient_ID', how='left')

    # Sort by Patient_ID
    merged_data = merged_data.sort_values(by='Patient_ID')

    # Report the number of patients with missing data
    # Count missing values in each specific lab column
    missing_ldh = merged_data['ldh'].isna().sum()
    missing_albumin = merged_data['alb'].isna().sum()
    missing_beta2_microglobulin = merged_data['b2m'].isna().sum()

    # Count missing values in cytogenetic columns
    cytogenetic_columns = ['t(4;14)', 't(14;16)', 'gain1q', 'del17p']
    missing_cytogenetic_data = merged_data[cytogenetic_columns].isna().any(axis=1).sum()

    # Count entries with missing lab values (1 or more missing)
    missing_lab_values = merged_data[['ldh', 'alb', 'b2m']].isna().any(axis=1).sum()

    # Count entries with missing lab or cytogenetic data
    missing_data = merged_data[['ldh', 'alb', 'b2m', 't(4;14)', 't(14;16)', 'gain1q', 'del17p']].isna().any(axis=1).sum()

    # Output results
    print("\nData missingness report: ")
    print(f"Number of entries missing LDH: {missing_ldh}")
    print(f"Number of entries missing albumin: {missing_albumin}")
    print(f"Number of entries missing beta2 microglobulin: {missing_beta2_microglobulin}")
    print(f"Number of entries missing one or more lab values: {missing_lab_values}")
    print(f"Number of entries missing cytogenetic data: {missing_cytogenetic_data}")
    print(f"Number of entries missing lab or cytogenetic data: {missing_data}")

    print("Calculating ISS, R-ISS and R2-ISS...\n")
    # Function to calculate ISS from labs
    def calculate_iss(row):
        if pd.isna(row['b2m']):
            return float('nan')
        elif row['b2m'] >= 5.5:
            return 3
        elif row['b2m'] >= 3.5 and row['b2m'] < 5.5:
            return 2
        elif  pd.isna(row['alb']):
            return float('nan')
        elif row['b2m'] < 3.5 and row['alb'] >= 35:
            return 1
        elif row['b2m'] < 3.5 and row['alb'] < 35:
            return 2
        else:
            return float('nan')

    # Apply the function to calculate ISS
    merged_data['ISS_calc'] = merged_data.apply(calculate_iss, axis=1)
    print(f"\n{merged_data['ISS_calc'].count()} ISS calculations added successfully.\n")

    # Check for discrepancies. First filter out rows where both ISS_calc and ISS_Compass are NaN
    filtered_data = merged_data.dropna(subset=['ISS_calc', 'ISS_Compass'], how='any')
    discrepant_rows = filtered_data['ISS_calc'] != filtered_data['ISS_Compass']
    print(f"ISS discrepancies detected for Patient IDs: {filtered_data[discrepant_rows]['Patient_ID'].tolist()}")
    # Resolve discrepancies by using Compass provided ISS
    filtered_data.loc[discrepant_rows, 'ISS_calc'] = filtered_data.loc[discrepant_rows, 'ISS_Compass']
    print(f"Discrepancies resolved by using Compass provided ISS for following row(s):\n{filtered_data[discrepant_rows]}\n")

    # Push the ISS_Compass and ISS_calc together, taking the first non-NaN value
    merged_data['ISS'] = merged_data['ISS_calc'].combine_first(merged_data['ISS_Compass'])
    # Drop the ISS_calc and ISS_Compass columns
    merged_data = merged_data.drop(columns=['ISS_calc', 'ISS_Compass'])

    # Function to calculate R-ISS
    def calculate_r_iss(row):
        if pd.isna(row['ldh']) or pd.isna(row['del17p']) or pd.isna(row['t(4;14)']) or pd.isna(row['t(14;16)']) or pd.isna(row['ISS']):
            return float('nan')
        elif row['ldh'] <= 2.8 and row['ISS'] == 1 and row['del17p'] == "Not Detected" and row['t(4;14)'] == "Not Detected" and row['t(14;16)'] == "Not Detected":
            return 1
        elif row['ISS'] == 3 and (row['ldh'] > 2.8 or row['del17p'] == "Detected" or row['t(4;14)'] == "Detected" or row['t(14;16)'] == "Detected"):
            return 3
        else:
            return 2

    # Apply function to calculate R-ISS
    merged_data['R-ISS'] = merged_data.apply(calculate_r_iss, axis=1)
    print(f"{merged_data['R-ISS'].count()} R-ISS calculations added successfully.")

    # Function to calculate R2-ISS points
    def calculate_r2_iss_points(row):
        if pd.isna(row['del17p']) or pd.isna(row['ldh']) or pd.isna(row['t(4;14)']) or pd.isna(row['gain1q']) or pd.isna(row['ISS']):
            return float('nan')
        
        points = 0
        if row['ISS'] == 2:
            points += 1
        elif row['ISS'] == 3:
            points += 1.5
        if row['del17p'] == "Detected":
            points += 1
        if row['ldh'] > 2.8:
            points += 1
        if row['t(4;14)'] == "Detected":
            points += 1
        if row['gain1q'] == "Detected":
            points += 0.5

        return points

    # Function to calculate R2-ISS risk category
    def calculate_r2_iss_risk(points):
        if pd.isna(points):
            return float('nan')
        elif points == 0:
            return 1
        elif 0.5 <= points <= 1:
            return 2
        elif 1.5 <= points <= 2.5:
            return 3
        elif 3 <= points <= 5:
            return 4
        else:
            return float('nan')

    # Apply functions to calculate R2-ISS points and risk category
    merged_data['R2-ISS Points'] = merged_data.apply(calculate_r2_iss_points, axis=1)
    merged_data['R2-ISS'] = merged_data['R2-ISS Points'].apply(calculate_r2_iss_risk)
    del merged_data['R2-ISS Points']
    print(f"{merged_data['R2-ISS'].count()} R2-ISS calculations added successfully.")

    # Store data in a excel file
    merged_data.to_excel("ISS.xlsx", index=False)
    print("\nData saved successfully to ISS.xlsx")
    
except Exception as e:
    print(f"An error occurred: {e}")