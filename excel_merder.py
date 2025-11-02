import os
import glob
import pandas as pd

def merge_excel_files(input_folder, output_file):
    """
    Merges the 'expressreport' worksheet from multiple Excel files into a single
    file and formats the header row with a specific background color and centered text.
    """
    excel_files = glob.glob(os.path.join(input_folder, '*.xlsx'))
    print(f"Found {len(excel_files)} Excel files in '{input_folder}'")

    if not excel_files:
        print("No Excel files found in the folder.")
        return

    df_list = []

    for file in excel_files:
        try:
            filename = os.path.basename(file)
            print(f"Reading 'ExpressReport' from: {filename}")
            df = pd.read_excel(file, sheet_name='expressreport', header=0)
            print(f"  Successfully read {len(df)} rows.")
            df_list.append(df)
        except ValueError:
            print(f"  Warning: Worksheet 'expressreport' not found in {filename}. Skipping this file.")
        except Exception as e:
            print(f"  Error reading {filename}: {str(e)}")

    if not df_list:
        print("\nNo data was read. Ensure your files contain a worksheet named 'expressreport'.")
        return

    try:
        print("\nStarting merge process...")
        merged_df = pd.concat(df_list, ignore_index=True)
        print(f"Successfully merged {len(merged_df)} total rows from all files.")

        print(f"Attempting to save and format the file as '{os.path.basename(output_file)}'...")

        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, sheet_name='expressreport', index=False, header=False, startrow=1)

            workbook = writer.book
            worksheet = writer.sheets['expressreport']

            # Updated header format to include center alignment
            header_format = workbook.add_format({
                'bold': True,
                'align': 'center',      # Center horizontally
                'valign': 'vcenter',    # Center vertically
                'fg_color': '#6F6F6F',
                'font_color': 'white',
                'border': 1
            })

            for col_num, value in enumerate(merged_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                column_len = max(
                    len(str(value)),
                    merged_df[value].astype(str).str.len().max()
                )
                worksheet.set_column(col_num, col_num, column_len + 4) # Added a bit more padding for centered text

        print("\nChecking if file was created...")
        if os.path.exists(output_file):
            size_mb = os.path.getsize(output_file) / (1024 * 1024)
            print(f"Success! Output file created: {output_file}")
            print(f"File size: {size_mb:.2f} MB")
        else:
            print("Error: File was not created")

    except PermissionError:
        print(f"\nError: Permission denied. Please make sure '{os.path.basename(output_file)}' is not open.")
    except MemoryError:
        print("\nError: Not enough memory to complete the operation.")
    except Exception as e:
        print(f"\nAn error occurred while writing the merged file: {str(e)}")

if __name__ == "__main__":
    # --- New ASCII Art Banner ---
    ascii_art = r"""
███████╗██╗  ██╗ ██████╗███████╗██╗     ███╗   ███╗███████╗██████╗  ██████╗ ███████╗██████╗
██╔════╝╚██╗██╔╝██╔════╝██╔════╝██║     ████╗ ████║██╔════╝██╔══██╗██╔═══██╗██╔════╝██╔══██╗
█████╗   ╚███╔╝ ██║     █████╗  ██║     ██╔████╔██║█████╗  ██████╔╝██║   ██║█████╗  ██████╔╝
██╔══╝   ██╔██╗ ██║     ██╔══╝  ██║     ██║╚██╔╝██║██╔══╝  ██╔══██╗██║   ██║██╔══╝  ██╔══██╗
███████╗██╔╝ ██╗╚██████╗███████╗███████╗██║ ╚═╝ ██║███████╗██║  ██║╚██████╔╝███████╗██║  ██║
╚══════╝╚═╝  ╚═╝ ╚═════╝╚══════╝╚══════╝╚═╝     ╚═╝╚══════╝╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═╝  ╚═╝
    """
    print(ascii_art)
    # --- End of Banner ---

    current_directory = os.getcwd()
    input_folder = current_directory
    output_filename = "merge File.xlsx"
    output_file = os.path.join(current_directory, output_filename)

    print(f"Script running in: {current_directory}")
    print("-" * 60)

    merge_excel_files(input_folder, output_file)