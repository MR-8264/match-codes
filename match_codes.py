import csv
import msoffcrypto
import pandas as pd
from io import BytesIO
import getpass
from datetime import datetime

# Get date for output file
current_date = datetime.now().strftime("%m%d%y")

# Get filename
filename  = input("Enter filename (including .xlsx): ")

# Ask for password
password = getpass.getpass("Enter password: ")

# Create Data Lists
patients = []
drug_codes = set()
matches = []

# Load patient data
try:
    with open(filename, "rb") as file:
        decrypted = BytesIO()
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)  # Use user-entered password
        office_file.decrypt(decrypted)

    # Load the decrypted file into a pandas DataFrame
    df = pd.read_excel(decrypted, engine="openpyxl")

    # Convert each row into a dictionary
    for _, row in df.iterrows():
        patients.append({
            "MRN": row.iloc[0],  # Use .iloc for positional access
            "Patient Last Name": row.iloc[2],
            "DOB": row.iloc[3],
            "Problem List with ICD 10": str(row.iloc[7]).strip().split(","),  
            "Encounter Date": row.iloc[10],
            "Primary Payer": row.iloc[11],
            "Provider": row.iloc[12]
    })

    # Load drug list data and collect diagnosis codes
    with open("Code_List.csv") as file2:
        reader2 = csv.reader(file2)
        for row in reader2:
            drug_codes.update([code.strip().upper() for code in row[1].replace(" ", "").split(",")])

    # Process each patient
    for patient in patients:
        mrn = patient["MRN"]
        name = patient["Patient Last Name"]
        dob = patient["DOB"]
        date = patient["Encounter Date"]
        payer = patient["Primary Payer"]
        provider = patient["Provider"]

        # Collect unique ICD10 codes for the patient that match the drug list
        patient_codes = set(code.strip().upper() for code in patient["Problem List with ICD 10"])
        matching_codes = patient_codes & drug_codes  # Find intersection of patient and drug codes

        # Add unique matches to the output
        for code in matching_codes:
            matches.append({
                "MRN": mrn,
                "Patient Last Name": name,
                "DOB": dob,
                "ICD10 Code": code,
                "Encounter Date": date,
                "Primary Payer": payer,
                "Provider": provider
            })

    # Export matches to a CSV file
    csv_filename = f"Two_Week_Appointment_Report_{current_date}.csv"
    csv_headers = ["MRN", "Patient Last Name", "DOB", "Encounter Date", "ICD10 Code", "Primary Payer", "Provider"]
    
    with open(csv_filename, mode='w', newline='') as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=csv_headers)
        writer.writeheader()  # Write header row
        writer.writerows(matches)  # Write data rows

    print(f"Data successfully exported to {csv_filename}")

except FileNotFoundError:
    print(f"\nError: The file '{filename}' was not found.")
except Exception as e:
    print(f"\nAn error occurred: {e}")


