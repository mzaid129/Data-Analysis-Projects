import win32com.client as win32
import os
import re
import shutil
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF


# --- Step 0: Helper function to clean sales values ---
def clean_omzet_value(val):
    if pd.isna(val):
        return 0.0

    val = str(val).strip()

    # Preserve negative sign if present
    is_negative = '-' in val

    val = val.replace('-', '')       # Temporarily remove negative for cleaning
    val = val.replace(' ', '')       # Remove spaces
    val = re.sub(r'[^\d]', '', val)  # Keep only digits

    if not val.isdigit():
        return 0.0

    if len(val) < 3:
        val = val.zfill(3)  # Pad with zeros if needed

    try:
        cleaned = float(val[:-2] + '.' + val[-2:])
        return -cleaned if is_negative else cleaned
    except:
        return 0.0


# --- Step 1: Locate and copy the closest file ---
def find_closest_verrichtingen_file(folder=r'your folder path'):
    today = datetime.today()
    pattern = r"your file name pattern"
    closest_file = None
    closest_diff = None

    for filename in os.listdir(folder):
        match = re.match(pattern, filename)
        if match:
            file_date_str = match.group(1)
            try:
                file_date = datetime.strptime(file_date_str, "%d-%m-%Y")
                date_diff = abs((file_date - today).days)
                if closest_diff is None or date_diff < closest_diff:
                    closest_diff = date_diff
                    closest_file = filename
            except ValueError:
                continue

    if closest_file:
        print(f"📁 Closest file found: {closest_file}")
        file_path = os.path.join(folder, closest_file)
        destination_folder = os.path.dirname(os.path.abspath(__file__))
        destination_path = os.path.join(destination_folder, os.path.basename(file_path))
        shutil.copy(file_path, destination_path)
        print(f"✅ Copied file to: {destination_path}")
        return destination_path
    else:
        print("❌ No matching file found.")
        exit()

input_file = find_closest_verrichtingen_file()



# --- Step 2: Load and clean the CSV ---
def load_clean_csv(path):
    for skip in range(5):
        df = pd.read_csv(path, encoding='latin1', sep=',', skiprows=skip, engine='python')
        df.columns = df.columns.str.strip()
        if any('Datum' in col for col in df.columns):
            break
    else:
        print("❌ Couldn't find any column resembling 'your desired column name'.")
        exit()

    # Normalize column names
    
    df.columns = df.columns.str.strip()
    df.rename(columns={col: 'column' for col in df.columns if 'column' in col}, inplace=True)



    # Clean 'sales'
    if 'sales' in df.columns:
        df['sales'] = df['sales'].astype(str).str.replace(r'\s+', '', regex=True)
        df['sales'] = df['sales'].str.replace('.', '', regex=False)
        df['sales'] = df['sales'].str.replace(',', '.', regex=False)
        df['sales'] = df['sales'].apply(clean_omzet_value)
        df['sales'] = pd.to_numeric(df['sales'], errors='coerce').fillna(0)
    else:
        print("❌ 'sales' column not found.")
        exit()

    

    # Parse 'Date' as datetime64[ns] — do not convert to .dt.date here!
    # Force convert 'Date' to datetime
    df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')
    print("📅 Parsed Date column:")
    print(df['Date'].head())

    # Drop rows where date couldn't be parsed
    df = df[df['Date'].notna()]

    return df


df = load_clean_csv(input_file)

# --- Step 3: Determine report date (yesterday or Saturday if Monday) ---
today = datetime.today()
if today.weekday() == 0:  # Monday
    report_date = (today - timedelta(days=2)).date()  # Saturday
else:
    report_date = (today - timedelta(days=1)).date()  # Yesterday

# ✅ Correct comparison using .dt.date
# Set the report date as a datetime object
report_date = pd.to_datetime(report_date)

# Filter
df_selected = df[df['Date'].dt.date == report_date.date()]


if df_selected.empty:
    print(f"❌ No data found for report date")
    exit()

print(f"📄 Generating report for: <given date>")

# --- Step 4: PDF Report Generation ---
output_dir = 'sales_reports_pdf'
os.makedirs(output_dir, exist_ok=True)

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'sales report', ln=True, align='C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', align='C')

    def add_table(self, df):
        self.set_font('Arial', '', 10)
        columns = df.columns.tolist()
        col_widths = {
            columns[0]: 25,
            columns[1]: 128,
            columns[2]: 30
        }
        for col in columns:
            self.cell(col_widths[col], 8, str(col), border=1, align='C')
        self.ln()

        for _, row in df.iterrows():
            y_start = self.get_y()
            x_start = self.get_x()
            omschrijving = str(row[columns[1]]) if pd.notnull(row[columns[1]]) else ''
            lines = self.multi_cell(col_widths[columns[1]], 6, omschrijving, border=0, split_only=True)
            row_height = 6 * len(lines)
            if self.get_y() + row_height > self.page_break_trigger:
                self.add_page()
                for col in columns:
                    self.cell(col_widths[col], 8, str(col), border=1, align='C')
                self.ln()
                y_start = self.get_y()

            self.set_y(y_start)
            for col in columns:
                x = self.get_x()
                val = str(row[col]) if pd.notnull(row[col]) else ''
                if col == columns[1]:
                    self.multi_cell(col_widths[col], 6, val, border=1)
                    self.set_xy(x + col_widths[col], y_start)
                else:
                    self.set_xy(x, y_start)
                    self.cell(col_widths[col], row_height, val, border=1, align='L')
            self.ln(row_height)

columns_to_exclude = [
    'columns you want to exclude',
]

if 'column' not in df.columns:
    print("❌ Column not found.")
    exit()

for dentist, group in df_selected.groupby('Dentist Name'):
    pdf = PDF()
    pdf.add_page()
    pdf.image("logo.png", x=9, y=-8, w=50)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, f'Dentist Name: {dentist}', ln=True)
    pdf.ln(3)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 8, f'Date of report', ln=True)

    trimmed_group = group.drop(columns=columns_to_exclude, errors='ignore')
    pdf.add_table(trimmed_group)
    pdf.ln(2)

    total_patients = group['Patient: code'].dropna().unique()
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 11, f"unique data", ln=True)
    
    total_sales = group['sales'].sum()
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 10, f"total sales", ln=True)

    safe_name = f"{dentist.strip()} sales report for {report_date.strftime('%d-%m-%y')}.pdf".replace("/", "-")
    pdf.output(os.path.join(output_dir, safe_name))
    print(f"✅ PDF saved for {dentist}")


# Load the dentist email lookup file
try:
    email_df = pd.read_csv('Dentists Email.csv', encoding='latin1')
except Exception as e:
    print(f"❌ Failed to load Dentists Email.csv: {e}")
    exit()

# Normalize names
email_df['Dentist Name'] = email_df['Dentist Name'].str.strip()
email_lookup = dict(zip(email_df['Dentist Name'], email_df['Email Address']))

# Initialize email server
outlook = win32.Dispatch('email server')

# Loop through the generated reports
for filename in os.listdir(output_dir):
    if filename.endswith('.pdf'):
        dentist_name = filename.split(' sales report for ')[0].strip()
        report_path = os.path.join(output_dir, filename)

        if dentist_name in email_lookup:
            email_address = email_lookup[dentist_name]
            if pd.isna(email_address) or not str(email_address).strip():
                print(f"⚠️ Email not found for dentist: {dentist_name}")
                continue

            print(f"📧 Email found for {dentist_name}: {email_address}. Sending...")

            try:
                mail = outlook.CreateItem(0)
                mail.To = email_address
                mail.Subject = f"Sales Report - {dentist_name} ({report_date.strftime('%d-%m-%Y')})"
                mail.Body = f"your email body"
                if os.path.exists(report_path):
                    mail.Attachments.Add(os.path.abspath(report_path))
                else:
                    print(f"❌ Attachment file not found: {report_path}")
                    continue
                mail.Send()
                print(f"✅ Email sent successfully to {dentist_name}")
            except Exception as e:
                print(f"❌ Failed to send email to {dentist_name}: {e}")
        else:
            print(f"⚠️ No email entry found for dentist: {dentist_name}")


