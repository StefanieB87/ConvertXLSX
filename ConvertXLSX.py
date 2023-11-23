import tkinter as tk
from tkinter import filedialog
import openpyxl
import pandas as pd

def load_and_process_excel_file():
     file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

     if file_path:
            try:
                # Open the Excel file
                workbook = openpyxl.load_workbook(file_path)

                # Assuming there is only one sheet, you can access it like this
                sheet = workbook.active

                # Convert the sheet data to a Pandas DataFrame
                df = pd.DataFrame(sheet.values, columns=[cell.value for cell in sheet[1]])

                # Save the DataFrame to a temporary CSV file
                temp_csv_path = "temp_output.csv"
                df.to_csv(temp_csv_path, index=False)

                # Execute the provided code on the temporary CSV file
                output_text = execute_code_on_csv(temp_csv_path)

                # Display the output in a text widget
                output_text_widget.config(state=tk.NORMAL)
                output_text_widget.delete(1.0, tk.END)
                output_text_widget.insert(tk.END, output_text)
                output_text_widget.config(state=tk.DISABLED)

            except Exception as e:
                print(f"Error loading and processing the Excel file: {e}")

def execute_code_on_csv(csv_path):
        try:
            # Read the CSV file using the provided code
            xlsin = pd.read_csv(csv_path)

            colnames = xlsin.columns.values
            records = xlsin.to_records(index=False)

            output_text = '{|class="wikitable sortable" border="1"\n'
            for c in colnames:
                output_text += '! %s\n' % c

            for r in records:
                output_text += '|-\n'
                output_text += '|%s\n' % '||'.join([str(x) for x in r])

            output_text += '|}'

            return output_text

        except Exception as e:
            print(f"Error executing code on CSV file: {e}")
            return ''

def save_output_text():
        # Allow the user to choose a location to save the text file
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])

        if file_path:
            # Get the text from the output text widget
            output_text = output_text_widget.get(1.0, tk.END)

            # Save the text to the chosen file
            with open(file_path, "w") as file:
                file.write(output_text)

# Create the main window
root = tk.Tk()
root.title("Excel File Processor")

# Create a button to trigger the file dialog and processing
process_button = tk.Button(root, text="Choose and Process Excel File", command=load_and_process_excel_file)
process_button.pack(pady=20)

# Create a text widget to display the processed output
output_text_widget = tk.Text(root, wrap=tk.WORD, height=10, width=50, state=tk.DISABLED)
output_text_widget.pack(pady=20)

# Create a button to save the processed output as a text file
save_button = tk.Button(root, text="Save Output as Text File", command=save_output_text)
save_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()