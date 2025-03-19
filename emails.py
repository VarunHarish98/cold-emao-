import pandas as pd
import re

def company_to_domain(company):
    """
    Convert a company name into a simplified domain.
    This function lowercases the name, removes non-alphanumeric characters,
    and appends '.com'.
    """
    company = company.lower()
    company = re.sub(r'[^a-z0-9]', '', company)
    return company + ".com"

def generate_email(name, company, format_type="first.last"):
    """
    Generate an email address based on the provided format_type.
    
    Supported format_types:
    - "first.last": e.g., john.doe@company.com
    - "firstinitiallastname": e.g., jdoe@company.com
    - "first": e.g., john@company.com
    
    If the name has only one part, it will simply use that as the username.
    """
    parts = name.strip().split()
    if len(parts) == 0:
        return ""
    
    if format_type == "first.last":
        if len(parts) >= 2:
            first = parts[0].lower()
            last = parts[-1].lower()
            return f"{first}.{last}@{company_to_domain(company)}"
        else:
            return f"{parts[0].lower()}@{company_to_domain(company)}"
    elif format_type == "firstinitiallastname":
        if len(parts) >= 2:
            first_initial = parts[0][0].lower()
            last = parts[-1].lower()
            return f"{first_initial}{last}@{company_to_domain(company)}"
        else:
            return f"{parts[0].lower()}@{company_to_domain(company)}"
    elif format_type == "first":
        return f"{parts[0].lower()}@{company_to_domain(company)}"
    else:
        raise ValueError("Unsupported email format type.")

def main(input_excel, output_excel, format_type="first.last"):
    # Read the Excel file (ensure it contains columns named "Company" and "Name")
    df = pd.read_excel(input_excel)
    
    # Verify that the expected columns exist
    if not all(col in df.columns for col in ['Company', 'Name']):
        print("Error: The input Excel file must contain 'Company' and 'Name' columns.")
        return
    
    # Generate email addresses for each row based on the selected format
    df['Email'] = df.apply(
        lambda row: generate_email(row['Name'], row['Company'], format_type=format_type), axis=1
    )
    
    # Save the updated DataFrame to a new Excel file
    df.to_excel(output_excel, index=False)
    print(f"Email addresses generated and saved to {output_excel}")

if __name__ == '__main__':
    # Replace these with your actual file paths and desired format
    input_excel = "input.xlsx"   # Input file with 'Company' and 'Name' columns
    output_excel = "output.xlsx" # Output file where results will be saved
    # Choose format: "first.last", "firstinitiallastname", or "first"
    chosen_format = "firstinitiallastname"
    
    main(input_excel, output_excel, format_type=chosen_format)
