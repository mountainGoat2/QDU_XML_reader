import os
import pandas as pd
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox

def extract_data(xml_file, sigma):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Namespace dictionary
    ns = {'n42': 'http://physics.nist.gov/N42/2011/N42',
          'mis-n42': 'http://www.microstepmis.com/N42/756/Extension'}

    # Find Nuclide tags for Alpha, Beta, and Rn-222
    alpha_data = root.find(".//n42:Nuclide[n42:NuclideName='Alpha']", ns)
    beta_data = root.find(".//n42:Nuclide[n42:NuclideName='Beta']", ns)
    rn_data = root.find(".//n42:Nuclide[n42:NuclideName='Rn-222']", ns)

    # Extract Alpha data
    alpha_activity = None
    alpha_error = None
    if alpha_data is not None:
        alpha_activity_elem = alpha_data.find(".//n42:NuclideActivityValue", ns)
        alpha_error_elem = alpha_data.find(".//n42:NuclideIDConfidenceUncertaintyValue", ns)
        if alpha_activity_elem is not None:
            alpha_activity = float(alpha_activity_elem.text)
        if alpha_error_elem is not None:
            alpha_error = float(alpha_error_elem.text)

    # Extract Beta data
    beta_activity = None
    beta_error = None
    if beta_data is not None:
        beta_activity_elem = beta_data.find(".//n42:NuclideActivityValue", ns)
        beta_error_elem = beta_data.find(".//n42:NuclideIDConfidenceUncertaintyValue", ns)
        if beta_activity_elem is not None:
            beta_activity = float(beta_activity_elem.text)
        if beta_error_elem is not None:
            beta_error = float(beta_error_elem.text)

    # Extract Rn-222 data
    rn_activity = None
    rn_error = None
    rn_concent = None
    rn_con_error = None
    if rn_data is not None:
        rn_activity_elem = rn_data.find(".//n42:NuclideActivityValue", ns)
        rn_error_elem = rn_data.find(".//n42:NuclideIDConfidenceUncertaintyValue", ns)
        rn_concent_elem = rn_data.find(".//mis-n42:NuclideConcentration", ns)
        rn_con_error_elem = rn_data.find(".//mis-n42:NuclideConcentrationError", ns)
        if rn_activity_elem is not None:
            rn_activity = float(rn_activity_elem.text)
        if rn_error_elem is not None:
            rn_error = float(rn_error_elem.text)
        if rn_concent_elem is not None:
            rn_concent = float(rn_concent_elem.text)
        if rn_con_error_elem is not None:
            rn_con_error = float(rn_con_error_elem.text)

    # Extract Record DateTime
    record_datetime = None
    measured_flow = None
    pressure_dp = None
    record_datetime_elem = root.find(".//mis-n42:MeasurementRecordDateTime", ns)
    measured_flow_elem = root.find(".//mis-n42:MeasurementFlow", ns)
    pressure_dp_elem = root.find(".//mis-n42:MeasurementDeltaPressure", ns)
    if record_datetime_elem is not None:
        record_datetime = record_datetime_elem.text
    if measured_flow_elem is not None:
        measured_flow = measured_flow_elem.text
    if pressure_dp_elem is not None:
        pressure_dp = pressure_dp_elem.text

    # Determine the values to write in Excel
    if alpha_activity == 0 and alpha_error == 0:
        alpha_activity_value = "<MDA"
    elif alpha_activity is None or alpha_error is None or alpha_activity < sigma * alpha_error:
        alpha_activity_value = "<MDA"
    else:
        alpha_activity_value = f"{alpha_activity:.4f} ± {alpha_error:.4f}"

    if beta_activity == 0 and beta_error == 0:
        beta_activity_value = "<MDA"
    elif beta_activity is None or beta_error is None or beta_activity < sigma * beta_error:
        beta_activity_value = "<MDA"
    else:
        beta_activity_value = f"{beta_activity:.4f} ± {beta_error:.4f}"

    if rn_activity == 0 and rn_error == 0:
        rn_activity_value = "<MDA"
    elif rn_activity is None or rn_error is None or rn_activity < sigma * rn_error:
        rn_activity_value = "<MDA"
    else:
        rn_activity_value = f"{rn_activity:.4f} ± {rn_error:.4f}"

    if rn_concent == 0 and rn_con_error == 0:
        rn_conc_value = f"0.0000 ± 0.0000"
    elif rn_concent is None or rn_con_error is None or rn_concent < sigma * rn_con_error:
        rn_conc_value = "<MDA"
    else:
        rn_conc_value = f"{rn_concent:.4f} ± {rn_con_error:.4f}"

    return alpha_activity_value, beta_activity_value, rn_activity_value, rn_conc_value, record_datetime, measured_flow, pressure_dp


def select_directory():
    selected_directory = filedialog.askdirectory()
    if selected_directory:
        directory_entry.delete(0, tk.END)
        directory_entry.insert(tk.END, selected_directory)

def select_excel_location():
    excel_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if excel_file_path:
        excel_entry.delete(0, tk.END)
        excel_entry.insert(tk.END, excel_file_path)

def run_extraction():
    directory_path = directory_entry.get()
    excel_file_path = excel_entry.get()
    if directory_path:
        try:
            sigma_value = float(sigma_entry.get())
            dfs = []
            for subdir, dirs, files in os.walk(directory_path):
                for file in files:
                    if file.endswith('.xml'):
                        xml_file_path = os.path.join(subdir, file)
                        alpha_activity, beta_activity, rn_activity, rn_concentration, record_datetime, flow_rate, dp_mbar = extract_data(xml_file_path, sigma_value)
                        df = pd.DataFrame({'Alpha Activity (Bq)': [alpha_activity],
                                           'Beta Activity (Bq)': [beta_activity],
                                           'Rn-222 Activity (Bq)': [rn_activity],
                                           'Rn-222 Concentration (Bq/m^3)': [rn_concentration],
                                           'Record DateTime': [record_datetime],
                                           'Flow rate (m^3/h)': [flow_rate],
                                           'dp (mbar)': [dp_mbar]})
                        dfs.append(df)

            final_df = pd.concat(dfs, ignore_index=True)

            if excel_file_path:
                # Set column width dynamically based on the length of column headers
                column_widths = {col: max(final_df[col].astype(str).map(len).max(), len(col)) for col in final_df.columns}
                writer = pd.ExcelWriter(excel_file_path, engine='xlsxwriter')
                final_df.to_excel(writer, index=False, sheet_name='Sheet1')
                worksheet = writer.sheets['Sheet1']
                for i, col in enumerate(final_df.columns):
                    worksheet.set_column(i, i, column_widths[col] + 2)  # Add a little extra width
                writer.close()
                messagebox.showinfo("Extraction Complete", "Data has been extracted and written to Excel successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
    else:
        messagebox.showwarning("Warning", "Please select a directory first!")

# Create the main window
root = tk.Tk()
root.title("XML Data Extraction")

# Create and place widgets
directory_label = tk.Label(root, text="Select Directory:")
directory_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

directory_entry = tk.Entry(root, width=50)
directory_entry.grid(row=0, column=1, padx=5, pady=5)

select_button = tk.Button(root, text="Select Directory", command=select_directory)
select_button.grid(row=0, column=2, padx=5, pady=5)

sigma_label = tk.Label(root, text="Enter Sigma Value (Default is 2):")
sigma_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)

sigma_entry = tk.Entry(root, width=10)
sigma_entry.insert(tk.END, "2")
sigma_entry.grid(row=1, column=1, padx=5, pady=5)

excel_label = tk.Label(root, text="Select Excel File Location:")
excel_label.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)

excel_entry = tk.Entry(root, width=50)
excel_entry.grid(row=2, column=1, padx=5, pady=5)

excel_button = tk.Button(root, text="Select Excel Location", command=select_excel_location)
excel_button.grid(row=2, column=2, padx=5, pady=5)

run_button = tk.Button(root, text="Run Extraction", command=run_extraction)
run_button.grid(row=3, columnspan=3, padx=5, pady=5)

# Start the GUI event loop
root.mainloop()
