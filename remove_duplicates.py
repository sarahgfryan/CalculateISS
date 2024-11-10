import pandas as pd

# Load the data
patient_visit_data = pd.read_csv("MMRF_CoMMpass_IA22_PER_PATIENT_VISIT.tsv", delimiter="\t", low_memory=False)

# Identify duplicated rows based on 'PUBLIC_ID' and 'VJ_INTERVAL'
duplicates = patient_visit_data[patient_visit_data.duplicated(subset=['PUBLIC_ID', 'VJ_INTERVAL'], keep=False)]

# Save the duplicated entries to a new TSV file
duplicates.to_csv("Duplicated_Entries.tsv", sep="\t", index=False)

print(f"Duplicated entries saved to 'Duplicated_Entries.tsv' with {len(duplicates)} rows.")