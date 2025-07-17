import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from fpdf import FPDF
from email.message import EmailMessage
import smtplib
from dotenv import load_dotenv

load_dotenv(dotenv_path="C:/Users/KIIT/Desktop/SAIL/.env")
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

master_file = ""
changes_file = ""
last_change_report = None
new_joinees_report = None
key_column_name = "SAIL_PERNO"
current_view = None

def get_distinct_colors(n):
    colors = [
        "#FFCCCC", "#E6FFCC", "#CCFFE5", "#CCE5FF", "#E5CCFF",
        "#FFCCF2", "#F2F2F2", "#D9F2E6", "#FFF2CC", "#CCE8FF"
    ]
    return [colors[i % len(colors)] for i in range(n)]

def read_file(file_path):
    dtypes = {'BANK_ACNO': str, 'PRAN_NO': str, 'SAIL_PERNO': str}
    if file_path.endswith('.csv'):
        return pd.read_csv(file_path, dtype=dtypes)
    elif file_path.endswith(('.xlsx', '.xls')):
        return pd.read_excel(file_path, dtype=dtypes)
    else:
        raise ValueError("Select a CSV or Excel file.")

def compare_files(master_file, changes_file, key_column):
    master_df = read_file(master_file)
    changes_df = read_file(changes_file)

    if key_column not in master_df.columns or key_column not in changes_df.columns:
        raise KeyError(f"Column '{key_column}' not found in both files!")

    master_df.set_index(key_column, inplace=True)
    changes_df.set_index(key_column, inplace=True)

    common_ids = master_df.index.intersection(changes_df.index)
    common_columns = master_df.columns.intersection(changes_df.columns)

    exclude_cols = {'YYYYMM', 'SEPR_YYYYMM'}
    compare_columns = [col for col in common_columns if col not in exclude_cols]

    changes_list = []

    for idx in common_ids:
        for col in compare_columns:
            old_value = master_df.loc[idx, col]
            new_value = changes_df.loc[idx, col]

            if pd.isnull(old_value) and pd.isnull(new_value):
                continue
            if old_value != new_value:
                if isinstance(old_value, float) and pd.notnull(old_value) and pd.notnull(new_value):
                    try:
                        decimal_places = len(str(old_value).split(".")[1])
                        new_value = round(float(new_value), decimal_places)
                    except:
                        pass
                changes_list.append({
                    key_column: idx,
                    'Column Changed': col,
                    'Old Value': old_value if pd.notnull(old_value) else '-',
                    'New Value': new_value if pd.notnull(new_value) else '-'
                })

    return pd.DataFrame(changes_list)

def upload_master_file():
    global master_file
    master_file = filedialog.askopenfilename(title="Select Master File", filetypes=[("CSV/Excel Files", "*.csv *.xlsx *.xls")])
    master_label.config(text=f"Master File: {master_file.split('/')[-1] if master_file else 'None'}")

def upload_changes_file():
    global changes_file
    changes_file = filedialog.askopenfilename(title="Select Changes File", filetypes=[("CSV/Excel Files", "*.csv *.xlsx *.xls")])
    changes_label.config(text=f"Changes File: {changes_file.split('/')[-1] if changes_file else 'None'}")

def clear_treeview():
    for row in tree.get_children():
        tree.delete(row)

def display_dataframe(df, key_col):
    clear_treeview()
    df = df.fillna('-')
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, anchor=tk.CENTER, width=200, stretch=True)
    colors = get_distinct_colors(len(df[key_col].unique()))
    color_map = dict(zip(df[key_col].unique(), colors))
    for _, row in df.iterrows():
        tag = f"color_{row[key_col]}"
        if not tree.tag_has(tag):
            tree.tag_configure(tag, background=color_map[row[key_col]])
        tree.insert("", tk.END, values=list(row), tags=(tag,))

def run_comparison():
    global last_change_report, current_view
    if not master_file or not changes_file:
        messagebox.showwarning("Missing Files", "Upload both files.")
        return
    try:
        df = compare_files(master_file, changes_file, key_column_name)
        last_change_report = df
        current_view = 'changes'
        if df.empty:
            change_count_label.config(text="No changes found.")
            messagebox.showinfo("Info", "No differences detected.")
        else:
            display_dataframe(df, key_column_name)
            change_count_label.config(text=f"Total unique employees changed: {df[key_column_name].nunique()}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def save_report():
    if last_change_report is None or last_change_report.empty:
        messagebox.showinfo("No Data", "No changes to save.")
        return
    path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if path:
        try:
            last_change_report.to_excel(path, index=False)
            messagebox.showinfo("Success", f"Report saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

def search_by_key():
    value = search_entry.get().strip()
    if not value:
        return
    df = last_change_report if current_view == 'changes' else new_joinees_report
    if df is None:
        messagebox.showinfo("No Data", "Run comparison or show new joinees.")
        return
    filtered = df[df[key_column_name].astype(str) == value]
    if filtered.empty:
        messagebox.showinfo("No Match", f"No data for {value}")
    else:
        display_dataframe(filtered, key_column_name)

def find_new_joinees():
    global new_joinees_report, current_view
    if not master_file or not changes_file:
        messagebox.showwarning("Missing Files", "Upload both files.")
        return
    try:
        m_df = read_file(master_file)
        c_df = read_file(changes_file)
        new_ids = set(c_df['SAIL_PERNO']) - set(m_df['SAIL_PERNO'])
        df = c_df[c_df['SAIL_PERNO'].isin(new_ids)][['SAIL_PERNO', 'DOJ_SAIL', 'DOB', 'PAN', 'IFSC_CD', 'BANK_ACNO']]
        new_joinees_report = df
        current_view = 'new_joinees'
        display_dataframe(df, 'SAIL_PERNO')
        change_count_label.config(text=f"Total new joinees: {len(df)}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def save_new_joinees_pdf():
    if new_joinees_report is None or new_joinees_report.empty:
        messagebox.showinfo("No Data", "No new joinees to export.")
        return
    path = filedialog.asksaveasfilename(defaultextension=".pdf")
    if not path:
        return
    try:
        from fpdf.enums import XPos, YPos  
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font("helvetica", size=12)
        pdf.cell(200, 10, text="New Joinees Report", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        pdf.ln(10)

        df = new_joinees_report.fillna('-')

        for col in df.columns:
            pdf.cell(40, 10, col, border=1)
        pdf.ln()

        for _, row in df.iterrows():
            for col in df.columns:
                pdf.cell(40, 10, str(row[col])[:15], border=1)
            pdf.ln()

        pdf.output(path)
        messagebox.showinfo("Success", f"PDF saved:\n{path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def send_email_report():
    try:
        sender_email = EMAIL_SENDER
        app_password = EMAIL_PASSWORD

        receiver_email = simpledialog.askstring("Receiver Email", "Enter recipient email address:")

        if not sender_email or not app_password:
            messagebox.showerror("Missing Credentials", "Email sender or password is missing in .env file.")
            return

        if not receiver_email:
            messagebox.showwarning("No Email", "No recipient email entered.")
            return

        msg = EmailMessage()
        msg.set_content("Please find the attached comparison report.")
        msg['Subject'] = 'Comparison Report'
        msg['From'] = sender_email
        msg['To'] = receiver_email

        if last_change_report is not None:
            temp_path = "temp_report.xlsx"
            last_change_report.to_excel(temp_path, index=False)

            with open(temp_path, 'rb') as f:
                file_data = f.read()
                file_name = "Comparison_Report.xlsx"
                msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, app_password)
            smtp.send_message(msg)

        messagebox.showinfo("Email Sent", f"Report sent to {receiver_email} successfully!")

    except Exception as e:
        messagebox.showerror("Email Error", str(e))

root = tk.Tk()
root.title("File Comparison Tool")
root.state('zoomed')

style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview", font=('Arial', 12), rowheight=28)
style.configure("Treeview.Heading", font=('Arial', 12, 'bold'))

title_label = tk.Label(root, text="FILE COMPARISON TOOL", font=('Arial', 20, 'bold'), fg='blue')
title_label.pack(pady=(20, 10))

frame_top = tk.Frame(root)
frame_top.pack(pady=(5, 15), padx=20)

frame_master = tk.Frame(frame_top)
frame_master.pack(side=tk.LEFT, padx=30)
btn_master = tk.Button(frame_master, text="Upload Master File", command=upload_master_file, font=('Arial', 11), width=20)
btn_master.pack()
master_label = tk.Label(frame_master, text="No Master File", font=('Arial', 10))
master_label.pack(pady=(5, 0))

frame_changes = tk.Frame(frame_top)
frame_changes.pack(side=tk.LEFT, padx=30)
btn_changes = tk.Button(frame_changes, text="Upload Changes File", command=upload_changes_file, font=('Arial', 11), width=20)
btn_changes.pack()
changes_label = tk.Label(frame_changes, text="No Changes File", font=('Arial', 10))
changes_label.pack(pady=(5, 0))

frame_buttons = tk.Frame(root)
frame_buttons.pack(pady=(5, 20))
tk.Button(frame_buttons, text="Compare Files", command=run_comparison, font=('Arial', 12), bg='lightblue', width=18).pack(side=tk.LEFT, padx=10)
tk.Button(frame_buttons, text="Save Compared File", command=save_report, font=('Arial', 12), bg='lightgreen', width=20).pack(side=tk.LEFT, padx=10)
tk.Button(frame_buttons, text="Show New Joinees", command=find_new_joinees, font=('Arial', 12), bg='lightpink', width=18).pack(side=tk.LEFT, padx=10)
tk.Button(frame_buttons, text="Save New Joinees PDF", command=save_new_joinees_pdf, font=('Arial', 12), bg='orange', width=22).pack(side=tk.LEFT, padx=10)
tk.Button(frame_buttons, text="Email Report", command=send_email_report, font=('Arial', 12), bg='lightgray', width=18).pack(side=tk.LEFT, padx=10)

frame_search = tk.Frame(root)
frame_search.pack(pady=(0, 20))
tk.Label(frame_search, text="Search by Key:", font=('Arial', 12)).pack(side=tk.LEFT, padx=(0, 5))
search_entry = tk.Entry(frame_search, font=('Arial', 12), width=25)
search_entry.pack(side=tk.LEFT, padx=(0, 10))
tk.Button(frame_search, text="Search", command=search_by_key, font=('Arial', 12), bg='lightyellow', width=12).pack(side=tk.LEFT)
root.bind('<Return>', lambda event: search_by_key())

frame_table = tk.Frame(root)
frame_table.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))
tree_scroll_y = tk.Scrollbar(frame_table)
tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
tree_scroll_x = tk.Scrollbar(frame_table, orient=tk.HORIZONTAL)
tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
tree = ttk.Treeview(frame_table, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
tree.pack(fill=tk.BOTH, expand=True)
tree_scroll_y.config(command=tree.yview)
tree_scroll_x.config(command=tree.xview)

change_count_label = tk.Label(root, text="", font=('Arial', 12), fg='green')
change_count_label.pack(pady=(5, 15))

root.mainloop()