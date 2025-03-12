# Dash App to generate survey pdf for each employee from excel file

import os
import pandas as pd
from fpdf import FPDF
import dash
import dash_bootstrap_components as dbc
from dash import dcc, html, Input, Output
import zipfile

# Load Excel file
excel_file = "employee_data.xlsx"
df = pd.read_excel(excel_file)

# Output directory for PDFs
output_folder = "Employee_PDFs"
os.makedirs(output_folder, exist_ok=True)


# Custom PDF class to format the table and headings
class CustomPDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 16)
        self.cell(200, 10, "Employee Report", ln=True, align="C")
        self.ln(10)

    def colored_table(self, employee):
        columns = ["Field", "Details"]
        col_widths = [70, 100]

        self.set_fill_color(0, 102, 204)  # Blue Background
        self.set_text_color(255, 255, 255)  # White Text
        self.set_font("Arial", "B", 12)

        for col in columns:
            self.cell(col_widths[columns.index(col)], 10, col, border=1, align="C", fill=True)
        self.ln()

        self.set_text_color(0, 0, 0)
        self.set_font("Arial", "", 12)

        details = [
            ("Supervisor Number", employee["Supervisor Number"]),
            ("Name", employee["Name"]),
            ("Department", employee["Department"]),
            ("Location", employee["Location"]),
            ("Tenure", f"{employee['Tenure']} years"),
            ("Headcount", str(employee["Headcount"])),
            ("Survey Results", str(employee["Survey Results"]))
        ]

        for idx, (label, value) in enumerate(details):
            self.set_fill_color(240, 240, 240) if idx % 2 == 0 else self.set_fill_color(255, 255, 255)
            self.cell(col_widths[0], 10, label, border=1, fill=True)
            self.cell(col_widths[1], 10, str(value), border=1, fill=True)
            self.ln()


# Function to generate PDFs for all employees
def generate_all_pdfs():
    for _, row in df.iterrows():
        pdf = CustomPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.colored_table(row)
        pdf_filename = os.path.join(output_folder, f"{row['Name']}.pdf")
        pdf.output(pdf_filename)


# Function to create a ZIP file of all PDFs
def create_zip():
    zip_filename = "Employee_Reports.zip"
    with zipfile.ZipFile(zip_filename, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in os.listdir(output_folder):
            if file.endswith(".pdf"):
                zipf.write(os.path.join(output_folder, file), file)
    return zip_filename


# Generate PDFs on app start
generate_all_pdfs()

# Initialize Dash App with Bootstrap Theme
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.MINTY])

# Dash Layout with Bootstrap Styling
app.layout = dbc.Container(
    [
        dbc.Row(
            dbc.Col(
                html.H1("Employee PDF Generator", className="text-center mt-4 mb-3"),
                width=12
            ), style={"margin-top": "200px"}
        ),

        dbc.Row(
            dbc.Col(
                dbc.Button("Download All PDFs", id="download-btn", color="primary",
                           className="d-grid gap-2 col-4 mx-auto"),
                width=25
            ),style={"margin-top": "70px"}
        ),

        dcc.Download(id="zip-download"),
    ],

)


# Callback to download ZIP file
@app.callback(
    Output("zip-download", "data"),
    Input("download-btn", "n_clicks"),
    prevent_initial_call=True
)
def download_zip(n_clicks):
    zip_path = create_zip()
    return dcc.send_file(zip_path)


# Run the Dash app
if __name__ == "__main__":
    app.run_server(debug=True)
