import tkinter as tk
from tkinter import ttk, filedialog
import os
import subprocess
import ast
import openpyxl
import csv

# Create the main window
root = tk.Tk()
root.title('Python Script Runner')

def get_script_required_parameters(script_path):
    with open(script_path, 'r') as script_file:
        script_content = script_file.read()
    
    # Parse the script content into an Abstract Syntax Tree (AST)
    parsed_ast = ast.parse(script_content)
    
    # Find the REQUIRED_PARAMS variable and extract its value
    for node in ast.walk(parsed_ast):
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == 'REQUIRED_PARAMS':
                    if isinstance(node.value, ast.List):
                        required_params = []
                        for param in node.value.elts:
                            if isinstance(param, ast.List) and len(param.elts) == 2:
                                data_type = param.elts[0].id  # Get the type name
                                name = param.elts[1].s  # Use .s to get the string value
                                required_params.append((data_type, name))
                        return required_params

    return []


def run_script():
    selected_item = script_list.selection()
    if not selected_item:
        return

    script_name = script_list.item(selected_item, 'values')[0]
    script_path = script_list.item(selected_item, 'values')[1]
    required_params = get_script_required_parameters(script_path)

    if not required_params:
        submit_params({}, [], script_path, None)
        return

    # Create a popup window for entering parameter values
    popup = tk.Toplevel(root)
    popup.title('Enter Script Parameters')

    # Create input fields for each parameter using grid layout
    input_fields = {}
    for i, param_info in enumerate(required_params):
        param_type, param_name = param_info
        tk.Label(popup, text=f'Enter {param_name}:').grid(row=i, column=0, sticky=tk.W)
        
        if param_type == 'int':
            input_fields[param_name] = tk.Entry(popup)
            input_fields[param_name].grid(row=i, column=1, sticky=tk.W)
        elif param_type == 'float':
            input_fields[param_name] = tk.Entry(popup)
            input_fields[param_name].grid(row=i, column=1, sticky=tk.W)
        elif param_type == 'str':
            if param_name == 'input_file':
                input_fields[param_name] = tk.Entry(popup)
                input_fields[param_name].insert(0, 'Enter a file path (.csv or .xlsx)')
                input_fields[param_name].grid(row=i, column=1, sticky=tk.W)
                browse_button = tk.Button(popup, text=f'Browse for {param_name}', command=lambda name=param_name: browse_file(input_fields[param_name], popup))
                browse_button.grid(row=i, column=2)
            else:
                input_fields[param_name] = tk.Entry(popup)
                input_fields[param_name].grid(row=i, column=1, sticky=tk.W)
        elif param_type == 'bool':
            input_fields[param_name] = tk.BooleanVar()
            tk.Checkbutton(popup, text=f'{param_name}', variable=input_fields[param_name]).grid(row=i, column=1, sticky=tk.W)
        else:
            tk.Label(popup, text=f'Invalid parameter type: {param_type}. Please contact technical support.', fg='red').grid(row=i, column=1, sticky=tk.W)

    # Create the Submit button
    submit_button = tk.Button(popup, text='Submit', command=lambda: submit_params(input_fields, required_params, script_path, popup))
    submit_button.grid(row=len(required_params) + 2, column=0, columnspan=3)  # Place the button below the input fields and table


def browse_file(input_field, popup):
    file_path = filedialog.askopenfilename(
        title='Select a file',
        filetypes=[('Excel Files', '*.xlsx *.xls'), ('CSV Files', '*.csv')],
    )
    if file_path:
        input_field.delete(0, tk.END)
        input_field.insert(0, file_path)
        display_csv_or_excel_preview(file_path, popup)

def display_csv_or_excel_preview(file_path, popup):
    # Get the widget at the specified grid coordinate
    previous_table_widget = popup.grid_slaves(row=len(required_params) + 1, column=0)

    # Check if a widget exists at the specified coordinate
    if previous_table_widget:
        previous_table_widget[0].grid_forget()

    _, file_extension = os.path.splitext(file_path)

    if file_extension.lower() in ('.xlsx', '.xls'):
        # Read an Excel file
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True)
            sheet = wb.active
            data = []

            for row in sheet.iter_rows(max_row=10, values_only=True):
                data.append(row)

            if data:
                # Create a new frame for the table
                preview_frame = ttk.Frame(popup)
                preview_frame.grid(row=len(required_params) + 1, column=0, columnspan=3)

                # Create a Treeview widget to display the data
                tree = ttk.Treeview(preview_frame, columns=[str(i) for i in range(len(data[0]))], show="headings")
                tree.grid(row=0, column=0)

                # Add columns to the Treeview
                for i, header in enumerate(data[0]):
                    tree.heading(str(i), text=header)
                    tree.column(str(i), width=100)  # Adjust the width as needed

                # Add rows to the Treeview
                for row in data[1:]:
                    tree.insert("", "end", values=row)

        except Exception as e:
            print(f"Error reading Excel file: {str(e)}")

    elif file_extension.lower() == '.csv':
        # Read a CSV file
        try:
            with open(file_path, 'r', newline='') as csv_file:
                reader = csv.reader(csv_file)
                header = next(reader)
                data = [next(reader) for _ in range(10)]

                if data:
                    # Create a new frame for the table
                    preview_frame = ttk.Frame(popup)
                    preview_frame.grid(row=len(required_params) + 1, column=0, columnspan=3)

                    # Create a Treeview widget to display the data
                    tree = ttk.Treeview(preview_frame, columns=[str(i) for i in range(len(header))], show="headings")
                    tree.grid(row=0, column=0)

                    # Add columns to the Treeview
                    for i, col in enumerate(header):
                        tree.heading(str(i), text=col)
                        tree.column(str(i), width=100)  # Adjust the width as needed

                    # Add rows to the Treeview
                    for row in data:
                        tree.insert("", "end", values=row)

        except Exception as e:
            print(f"Error reading CSV file: {str(e)}")

def submit_params(input_fields, required_params, script_path, popup):
    # Collect the parameter values from input fields
    params = []
    for param_info in required_params:
        param_type, param_name = param_info
        if param_name in input_fields:
            param_value = input_fields[param_name].get()
            
            if param_type == 'int':
                param_value = int(param_value)
            elif param_type == 'float':
                param_value = float(param_value)
            elif param_type == 'bool':
                param_value = bool(param_value)
            
            params.append(param_value)

    # Close the parameter input window
    if popup:
        popup.destroy()

    # Check if all required parameters are provided
    if len(required_params) > 0 and len(params) == len(required_params):
        cmd = ['python', script_path] + [str(param) for param in params]
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True, text=True)

        # Read and print the output
        while True:
            line = process.stdout.readline()
            if not line:
                break
            print(line.strip())

        # Wait for the process to complete
        process.wait()
        
        root.destroy()
    elif len(required_params) == 0:
        cmd = ['python', script_path]
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True, text=True)

        # Read and print the output
        while True:
            line = process.stdout.readline()
            if not line:
                break
            print(line.strip())

        # Wait for the process to complete
        process.wait()
        
        root.destroy()
    else:
        print('Not all required parameters provided.')


def update_button_state(event):
    selected_items = script_list.selection()
    run_button.config(state=tk.NORMAL if selected_items else tk.DISABLED)


if __name__ == '__main__':
    # Create a treeview to display the script files
    script_list = ttk.Treeview(root, columns=('Script Name', 'Script Path', 'Required Parameters'))
    script_list.heading('#1', text='Script Name')
    script_list.heading('#2', text='Script Path')
    script_list.heading('#3', text='Required Parameters')
    script_list.column('#0', width=0, stretch=tk.NO)  # Remove the first column
    script_list.column('#1', width=100)
    script_list.column('#2', width=300)
    script_list.column('#3', width=200)

    script_list.grid(row=0, column=0, padx=10, pady=10, columnspan=2, sticky='nsew')  # Make the treeview expand with window

    # Get a list of Python script files in the current directory
    scripts_directory = os.path.dirname(os.path.abspath(__file__))
    current_script_name = os.path.basename(__file__)

    for script_file in os.listdir(scripts_directory):
        if script_file.endswith('.py') and script_file != current_script_name:
            script_name = os.path.basename(script_file)
            script_path = os.path.join(scripts_directory, script_file)
            required_params = ', '.join([f'{param[1]}' for param in get_script_required_parameters(script_path)])
            script_list.insert('', 'end', values=(script_name, script_path, required_params))

    # Create a Run Script button
    run_button = ttk.Button(root, text='Run Script', command=run_script, state=tk.DISABLED)
    run_button.grid(row=1, column=0, padx=10, pady=10)

    # Bind selection event to update button state
    script_list.bind('<<TreeviewSelect>>', update_button_state)

    # Calculate and set the window size based on the script_list size
    root.update()  # Needed to update widget sizes
    width = script_list.column('#1')['width'] + script_list.column('#2')['width'] + script_list.column('#3')['width'] + 20
    height = script_list.winfo_height() + run_button.winfo_height() + 50
    root.geometry(f'{width}x{height}')

    root.mainloop()