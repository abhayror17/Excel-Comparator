"""
Enhanced Excel File Comparison Tool

Key Improvements:
1. Uses proper file names instead of generic "file1/file2" references
2. Replaces composite key display with individual identifier columns (Channel Name, Program Date, Clip Start Time)
   for better Excel filtering and readability
3. Maintains internal composite key logic for comparison while presenting user-friendly individual columns

The tool compares Excel files using three unique identifiers and provides detailed
difference reports with proper file naming and filterable identifier columns.
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os
from tqdm import tqdm
import warnings
warnings.filterwarnings('ignore')

# Constants for unique identifiers
UNIQUE_IDENTIFIERS = ['Channel Name', 'Program Date', 'Clip Start Time']

def compare_excel_files(file1_path, file2_path, output_path=None):
    """
    Compares two Excel files and identifies differences between them.
    
    Args:
        file1_path (str): Path to the first Excel file
        file2_path (str): Path to the second Excel file
        output_path (str): Path for output Excel file with differences (optional)
    
    Returns:
        dict: Comparison results containing differences and statistics
    """
    
    print("="*80)
    print("🔍 EXCEL FILE COMPARISON TOOL")
    print("="*80)
    
    # File information
    file1_name = os.path.basename(file1_path)
    file2_name = os.path.basename(file2_path)
    
    print(f"\n📁 Comparing Files:")
    print(f"   File 1: {file1_name}")
    print(f"   File 2: {file2_name}")
    print(f"   Using proper file names for clear identification")
    
    # Check if files exist
    if not os.path.exists(file1_path):
        print(f"❌ Error: {file1_name} not found!")
        return None
    
    if not os.path.exists(file2_path):
        print(f"❌ Error: {file2_name} not found!")
        return None
    
    # Read Excel files and their sheets
    try:
        print(f"\n📊 Reading Excel files...")
        
        # Get sheet names for both files
        excel1_sheets = pd.ExcelFile(file1_path).sheet_names
        excel2_sheets = pd.ExcelFile(file2_path).sheet_names
        
        print(f"   {file1_name} sheets: {excel1_sheets}")
        print(f"   {file2_name} sheets: {excel2_sheets}")
        
        # Find common sheets
        common_sheets = list(set(excel1_sheets) & set(excel2_sheets))
        
        if not common_sheets:
            print("❌ No common sheets found between the files!")
            return None
        
        print(f"   Common sheets: {common_sheets}")
        
    except Exception as e:
        print(f"❌ Error reading Excel files: {str(e)}")
        return None
    
    comparison_results = {}
    
    # Compare each common sheet
    for sheet_name in common_sheets:
        print(f"\n🔍 Comparing sheet: '{sheet_name}'")
        print("-" * 60)
        
        try:
            # Read data from both files
            df1 = pd.read_excel(file1_path, sheet_name=sheet_name)
            df2 = pd.read_excel(file2_path, sheet_name=sheet_name)
            
            print(f"   {file1_name}: {len(df1)} rows, {len(df1.columns)} columns")
            print(f"   {file2_name}: {len(df2)} rows, {len(df2.columns)} columns")
            
            # Basic statistics
            stats = {
                'file1_rows': len(df1),
                'file1_columns': len(df1.columns),
                'file2_rows': len(df2),
                'file2_columns': len(df2.columns),
                'common_columns': list(set(df1.columns) & set(df2.columns)),
                'file1_unique_columns': list(set(df1.columns) - set(df2.columns)),
                'file2_unique_columns': list(set(df2.columns) - set(df1.columns))
            }
            
            print(f"   Common columns: {len(stats['common_columns'])}")
            print(f"   {file1_name} unique columns: {len(stats['file1_unique_columns'])}")
            print(f"   {file2_name} unique columns: {len(stats['file2_unique_columns'])}")
            
            # Find differences
            differences = analyze_differences(df1, df2, file1_name, file2_name, sheet_name)
            
            comparison_results[sheet_name] = {
                'statistics': stats,
                'differences': differences
            }
            
            # Display summary of differences
            display_difference_summary(differences, sheet_name, file1_name, file2_name)
            
        except Exception as e:
            print(f"❌ Error comparing sheet '{sheet_name}': {str(e)}")
            continue
    
    # Generate summary report
    print(f"\n📋 OVERALL COMPARISON SUMMARY")
    print("=" * 80)
    
    generate_summary_report(comparison_results, file1_name, file2_name)
    
    # Save differences to Excel if output path provided
    if output_path:
        save_differences_to_excel(comparison_results, file1_name, file2_name, output_path)
        print(f"\n💾 Differences saved to: {output_path}")
    
    return comparison_results

def create_composite_key(row, identifier_columns):
    """Create a composite key from identifier columns."""
    key_parts = []
    for col in identifier_columns:
        if col in row:
            value = row[col]
            if pd.isna(value):
                value = 'NULL'
            else:
                value = str(value).strip()
            key_parts.append(value)
        else:
            key_parts.append('MISSING')
    return '|'.join(key_parts)

def extract_identifier_values(row, identifier_columns):
    """Extract individual identifier values from a row."""
    identifier_values = {}
    for col in identifier_columns:
        if col in row:
            value = row[col]
            if pd.isna(value):
                value = 'NULL'
            else:
                value = str(value).strip()
            identifier_values[col] = value
        else:
            identifier_values[col] = 'MISSING'
    return identifier_values

def analyze_differences(df1, df2, file1_name, file2_name, sheet_name):
    """
    Enhanced difference analyzer using 3 unique identifiers: Channel Name, Program Date, Clip Start Time
    """
    print(f"\n🔍 Analyzing differences in sheet '{sheet_name}' using unique identifiers...")
    
    differences = {
        'rows_only_in_file1': [],
        'rows_only_in_file2': [],
        'identical_rows': [],
        'modified_rows': [],
        'column_differences': {},
        'summary': {}
    }
    
    # Check if unique identifiers exist in both files
    available_identifiers = []
    missing_identifiers = []
    
    for identifier in UNIQUE_IDENTIFIERS:
        if identifier in df1.columns and identifier in df2.columns:
            available_identifiers.append(identifier)
        else:
            missing_identifiers.append(identifier)
    
    print(f"   Available identifiers: {available_identifiers}")
    if missing_identifiers:
        print(f"   ⚠️ Missing identifiers: {missing_identifiers}")
    
    if len(available_identifiers) < 3:
        print(f"   ❌ Warning: Only {len(available_identifiers)} unique identifiers found. Comparison may not be accurate.")
    
    # Create composite keys for both dataframes
    print(f"   📝 Creating composite keys...")
    
    # Add composite key column to both dataframes
    df1_with_key = df1.copy()
    df2_with_key = df2.copy()
    
    df1_with_key['COMPOSITE_KEY'] = df1_with_key.apply(
        lambda row: create_composite_key(row, available_identifiers), axis=1
    )
    df2_with_key['COMPOSITE_KEY'] = df2_with_key.apply(
        lambda row: create_composite_key(row, available_identifiers), axis=1
    )
    
    # Create dictionaries for efficient lookup
    df1_dict = {row['COMPOSITE_KEY']: idx for idx, row in df1_with_key.iterrows()}
    df2_dict = {row['COMPOSITE_KEY']: idx for idx, row in df2_with_key.iterrows()}
    
    print(f"   File 1 unique records: {len(df1_dict)}")
    print(f"   File 2 unique records: {len(df2_dict)}")
    
    # Find common columns
    common_columns = list(set(df1.columns) & set(df2.columns))
    
    # Compare using composite keys
    all_keys = set(df1_dict.keys()) | set(df2_dict.keys())
    
    print(f"   🔍 Comparing {len(all_keys)} unique records...")
    
    for key in tqdm(all_keys, desc=f"    Progress", unit="record"):
        if key in df1_dict and key in df2_dict:
            # Records exist in both files - check for modifications
            row1_idx = df1_dict[key]
            row2_idx = df2_dict[key]
            
            row1 = df1.iloc[row1_idx]
            row2 = df2.iloc[row2_idx]
            
            modified_cols = {}
            for col in common_columns:
                val1 = row1[col]
                val2 = row2[col]
                
                # Handle NaN values
                if pd.isna(val1) and pd.isna(val2):
                    continue
                elif pd.isna(val1) or pd.isna(val2) or str(val1) != str(val2):
                    modified_cols[col] = {
                        'file1_value': val1,
                        'file2_value': val2
                    }
            
            if modified_cols:
                differences['modified_rows'].append({
                    'composite_key': key,
                    'file1_index': row1_idx,
                    'file2_index': row2_idx,
                    'modified_columns': modified_cols,
                    'identifier_values': extract_identifier_values(row1, available_identifiers)
                })
            else:
                differences['identical_rows'].append({
                    'composite_key': key,
                    'file1_index': row1_idx,
                    'file2_index': row2_idx,
                    'identifier_values': extract_identifier_values(row1, available_identifiers)
                })
        
        elif key in df1_dict:
            # Record only in file 1
            row1_idx = df1_dict[key]
            row1 = df1.iloc[row1_idx]
            differences['rows_only_in_file1'].append({
                'composite_key': key,
                'file1_index': row1_idx,
                'data': row1.to_dict(),
                'identifier_values': extract_identifier_values(row1, available_identifiers)
            })
        else:
            # Record only in file 2
            row2_idx = df2_dict[key]
            row2 = df2.iloc[row2_idx]
            differences['rows_only_in_file2'].append({
                'composite_key': key,
                'file2_index': row2_idx,
                'data': row2.to_dict(),
                'identifier_values': extract_identifier_values(row2, available_identifiers)
            })
    
    # Column-level differences
    differences['column_differences'] = {
        'file1_unique_columns': list(set(df1.columns) - set(df2.columns)),
        'file2_unique_columns': list(set(df2.columns) - set(df1.columns))
    }
    
    # Generate summary
    differences['summary'] = {
        'total_rows_file1': len(df1),
        'total_rows_file2': len(df2),
        'rows_only_in_file1': len(differences['rows_only_in_file1']),
        'rows_only_in_file2': len(differences['rows_only_in_file2']),
        'identical_rows': len(differences['identical_rows']),
        'modified_rows': len(differences['modified_rows']),
        'common_columns': len(common_columns),
        'file1_unique_columns': len(differences['column_differences']['file1_unique_columns']),
        'file2_unique_columns': len(differences['column_differences']['file2_unique_columns']),
        'available_identifiers': available_identifiers,
        'missing_identifiers': missing_identifiers
    }
    
    print(f"   ✅ Analysis complete!")
    return differences

def display_difference_summary(differences, sheet_name, file1_name, file2_name):
    """
    Enhanced display of difference summary with detailed progress information.
    """
    summary = differences['summary']
    
    print(f"\n📊 Sheet '{sheet_name}' - Detailed Comparison Summary:")
    print("=" * 60)
    print(f"🔑 Unique Identifiers Used:")
    print(f"   Available: {summary.get('available_identifiers', [])}")
    print(f"   Missing: {summary.get('missing_identifiers', [])}")
    
    print(f"\n📈 ROW COMPARISON:")
    print(f"   • Total rows in {file1_name}: {summary['total_rows_file1']:,}")
    print(f"   • Total rows in {file2_name}: {summary['total_rows_file2']:,}")
    print(f"   • Identical rows: {summary['identical_rows']:,}")
    print(f"   • Modified rows: {summary['modified_rows']:,}")
    print(f"   • Rows only in {file1_name}: {summary['rows_only_in_file1']:,}")
    print(f"   • Rows only in {file2_name}: {summary['rows_only_in_file2']:,}")
    
    print(f"\n🏗️ STRUCTURE COMPARISON:")
    print(f"   • Common columns: {summary['common_columns']}")
    print(f"   • Unique columns in {file1_name}: {summary['file1_unique_columns']}")
    print(f"   • Unique columns in {file2_name}: {summary['file2_unique_columns']}")
    
    print(f"\n📊 ACCURACY METRICS:")
    total_records = summary['identical_rows'] + summary['modified_rows']
    if total_records > 0:
        match_percentage = (summary['identical_rows'] / total_records) * 100
        print(f"   • Match Rate: {match_percentage:.2f}%")
    
    match_rate = (summary['identical_rows'] + summary['modified_rows']) / max(summary['total_rows_file1'], summary['total_rows_file2']) * 100 if max(summary['total_rows_file1'], summary['total_rows_file2']) > 0 else 0
    print(f"   • Data Coverage: {match_rate:.2f}%")
    
    # Show detailed modified rows if any
    if differences['modified_rows']:
        print(f"\n🔧 DETAILED MODIFICATIONS:")
        for i, mod in enumerate(differences['modified_rows'][:5]):  # Show first 5
            identifiers = mod.get('identifier_values', {})
            ident_str = ' | '.join([f"{k}: {v}" for k, v in identifiers.items()])
            print(f"   {i+1}. {ident_str}")
            print(f"      {file1_name} Row {mod['file1_index']} → {file2_name} Row {mod['file2_index']}")
            print(f"      Changes: {list(mod['modified_columns'].keys())}")
            if i >= 4:  # Limit output
                remaining = len(differences['modified_rows']) - 5
                if remaining > 0:
                    print(f"      ... and {remaining} more modifications")
                break
    
    # Show some example missing rows if any
    if differences['rows_only_in_file1']:
        print(f"\n➕ SAMPLE ROWS ONLY IN {file1_name}:")
        for i, row in enumerate(differences['rows_only_in_file1'][:3]):
            identifiers = row.get('identifier_values', {})
            ident_str = ' | '.join([f"{k}: {v}" for k, v in identifiers.items()])
            print(f"   {i+1}. {ident_str}")
            print(f"      Row Index: {row['file1_index']}")
    
    if differences['rows_only_in_file2']:
        print(f"\n➖ SAMPLE ROWS ONLY IN {file2_name}:")
        for i, row in enumerate(differences['rows_only_in_file2'][:3]):
            identifiers = row.get('identifier_values', {})
            ident_str = ' | '.join([f"{k}: {v}" for k, v in identifiers.items()])
            print(f"   {i+1}. {ident_str}")
            print(f"      Row Index: {row['file2_index']}")
    
    print(f"\n{'✅ RESULTS READY FOR EXCEL EXPORT' if summary['modified_rows'] > 0 or summary['rows_only_in_file1'] > 0 or summary['rows_only_in_file2'] > 0 else '📋 NO DIFFERENCES FOUND'}")
    print("=" * 60)

def generate_summary_report(comparison_results, file1_name, file2_name):
    """
    Generates a comprehensive summary report of the comparison.
    """
    total_sheets = len(comparison_results)
    sheets_with_differences = 0
    total_modifications = 0
    total_missing_rows_file1 = 0
    total_missing_rows_file2 = 0
    
    for sheet_name, results in comparison_results.items():
        summary = results['differences']['summary']
        
        if (summary['modified_rows'] > 0 or 
            summary['rows_only_in_file1'] > 0 or 
            summary['rows_only_in_file2'] > 0):
            sheets_with_differences += 1
        
        total_modifications += summary['modified_rows']
        total_missing_rows_file1 += summary['rows_only_in_file1']
        total_missing_rows_file2 += summary['rows_only_in_file2']
    
    print(f"\n📈 OVERALL STATISTICS:")
    print(f"   • Total sheets compared: {total_sheets}")
    print(f"   • Sheets with differences: {sheets_with_differences}")
    print(f"   • Total modified rows: {total_modifications}")
    print(f"   • Total rows only in {file1_name}: {total_missing_rows_file1}")
    print(f"   • Total rows only in {file2_name}: {total_missing_rows_file2}")
    
    if total_modifications == 0 and total_missing_rows_file1 == 0 and total_missing_rows_file2 == 0:
        print(f"\n✅ RESULT: Files are IDENTICAL!")
    else:
        print(f"\n⚠️  RESULT: Files have DIFFERENCES!")

def save_differences_to_excel(comparison_results, file1_name, file2_name, output_path):
    """
    Creates comprehensive Excel reports with multiple workbooks for detailed comparison.
    """
    print(f"\n📋 Creating comprehensive Excel report...")
    
    # Create main comparison report
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Create formats
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#4CAF50',
            'font_color': 'white',
            'border': 1
        })
        
        warning_format = workbook.add_format({
            'fg_color': '#FFEB3B',
            'border': 1
        })
        
        error_format = workbook.add_format({
            'fg_color': '#F44336',
            'font_color': 'white',
            'border': 1
        })
        
        success_format = workbook.add_format({
            'fg_color': '#4CAF50',
            'font_color': 'white',
            'border': 1
        })
        
        # 1. SUMMARY SHEET
        print(f"   📊 Creating Summary sheet...")
        summary_data = []
        for sheet_name, results in comparison_results.items():
            summary = results['differences']['summary']
            summary_data.append([
                sheet_name,
                summary['total_rows_file1'],
                summary['total_rows_file2'],
                summary['identical_rows'],
                summary['modified_rows'],
                summary['rows_only_in_file1'],
                summary['rows_only_in_file2'],
                summary['common_columns'],
                len(summary.get('available_identifiers', [])),
                len(summary.get('missing_identifiers', []))
            ])
        
        summary_df = pd.DataFrame(summary_data, columns=[
            'Sheet Name',
            f'{file1_name} Rows',
            f'{file2_name} Rows',
            'Identical Rows',
            'Modified Rows',
            f'Only in {file1_name}',
            f'Only in {file2_name}',
            'Common Columns',
            'Available Identifiers',
            'Missing Identifiers'
        ])
        
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Format summary sheet
        worksheet = writer.sheets['Summary']
        for col_num, value in enumerate(summary_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # 2. MODIFICATIONS DETAIL SHEET
        print(f"   🔧 Creating Modifications detail sheet...")
        mod_all_data = []
        for sheet_name, results in comparison_results.items():
            if results['differences']['modified_rows']:
                for mod in results['differences']['modified_rows']:
                    for col, changes in mod['modified_columns'].items():
                        mod_all_data.append([
                            sheet_name,
                            '',  # composite_key removed - now using individual identifiers
                            mod['file1_index'],
                            mod['file2_index'],
                            col,
                            str(changes['file1_value']),
                            str(changes['file2_value']),
                            'CHANGED'
                        ])
        
        if mod_all_data:
            # Extract identifier values from the first modification to get column structure
            first_mod = comparison_results[list(comparison_results.keys())[0]]['differences']['modified_rows'][0]
            identifier_cols = list(first_mod['identifier_values'].keys())
            
            # Create data with individual identifier columns
            mod_data_with_identifiers = []
            for sheet_name, results in comparison_results.items():
                if results['differences']['modified_rows']:
                    for mod in results['differences']['modified_rows']:
                        for col, changes in mod['modified_columns'].items():
                            row_data = [
                                sheet_name,
                                mod['file1_index'],
                                mod['file2_index'],
                                col,
                                str(changes['file1_value']),
                                str(changes['file2_value']),
                                'CHANGED'
                            ]
                            # Add individual identifier columns
                            for ident_col in identifier_cols:
                                row_data.append(mod['identifier_values'].get(ident_col, ''))
                            
                            mod_data_with_identifiers.append(row_data)
            
            # Create columns list
            columns = ['Sheet Name', f'{file1_name} Row', f'{file2_name} Row',
                      'Column', f'{file1_name} Value', f'{file2_name} Value', 'Change Type'] + identifier_cols
            
            mod_all_df = pd.DataFrame(mod_data_with_identifiers, columns=columns)
            mod_all_df.to_excel(writer, sheet_name='All_Modifications', index=False)
            
            # Format modifications sheet
            worksheet = writer.sheets['All_Modifications']
            for col_num, value in enumerate(mod_all_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
        
        # 3. UNIQUE RECORDS SHEETS
        print(f"   ➕ Creating unique records sheets...")
        unique_file1_data = []
        unique_file2_data = []
        
        for sheet_name, results in comparison_results.items():
            # Records only in file 1 - composite_key removed, will use individual identifiers
            for row in results['differences']['rows_only_in_file1']:
                unique_file1_data.append([
                    sheet_name,
                    '',  # composite_key removed - now using individual identifiers
                    row['file1_index']
                ] + [str(row['data'].get(col, '')) for col in results['statistics']['common_columns']])
            
            # Records only in file 2 - composite_key removed, will use individual identifiers
            for row in results['differences']['rows_only_in_file2']:
                unique_file2_data.append([
                    sheet_name,
                    '',  # composite_key removed - now using individual identifiers
                    row['file2_index']
                ] + [str(row['data'].get(col, '')) for col in results['statistics']['common_columns']])
        
        # Create unique records sheets with individual identifier columns
        if unique_file1_data:
            # Get identifier columns from first available result
            first_result = list(comparison_results.values())[0]
            available_identifiers = first_result['differences']['summary']['available_identifiers']
            common_cols = first_result['statistics']['common_columns']
            
            # Create data with individual identifier columns
            unique_file1_data_with_identifiers = []
            for sheet_name, results in comparison_results.items():
                for row in results['differences']['rows_only_in_file1']:
                    row_data = [
                        sheet_name,
                        row['file1_index']
                    ]
                    # Add individual identifier columns
                    for ident_col in available_identifiers:
                        row_data.append(row['identifier_values'].get(ident_col, ''))
                    
                    # Add common column data
                    row_data.extend([str(row['data'].get(col, '')) for col in common_cols])
                    unique_file1_data_with_identifiers.append(row_data)
            
            # Create columns list
            columns = ['Sheet Name', f'{file1_name} Row'] + available_identifiers + common_cols
            
            unique_file1_df = pd.DataFrame(unique_file1_data_with_identifiers, columns=columns)
            unique_file1_df.to_excel(writer, sheet_name=f'Only_{file1_name}', index=False)
            
            # Format sheet
            worksheet = writer.sheets[f'Only_{file1_name}']
            for col_num, value in enumerate(unique_file1_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
        
        if unique_file2_data:
            # Get identifier columns from first available result
            first_result = list(comparison_results.values())[0]
            available_identifiers = first_result['differences']['summary']['available_identifiers']
            common_cols = first_result['statistics']['common_columns']
            
            # Create data with individual identifier columns
            unique_file2_data_with_identifiers = []
            for sheet_name, results in comparison_results.items():
                for row in results['differences']['rows_only_in_file2']:
                    row_data = [
                        sheet_name,
                        row['file2_index']
                    ]
                    # Add individual identifier columns
                    for ident_col in available_identifiers:
                        row_data.append(row['identifier_values'].get(ident_col, ''))
                    
                    # Add common column data
                    row_data.extend([str(row['data'].get(col, '')) for col in common_cols])
                    unique_file2_data_with_identifiers.append(row_data)
            
            # Create columns list
            columns = ['Sheet Name', f'{file2_name} Row'] + available_identifiers + common_cols
            
            unique_file2_df = pd.DataFrame(unique_file2_data_with_identifiers, columns=columns)
            unique_file2_df.to_excel(writer, sheet_name=f'Only_{file2_name}', index=False)
            
            # Format sheet
            worksheet = writer.sheets[f'Only_{file2_name}']
            for col_num, value in enumerate(unique_file2_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
        
        # 4. INDENTIFIER ANALYSIS SHEET
        print(f"   🔑 Creating identifier analysis sheet...")
        identifier_data = []
        for sheet_name, results in comparison_results.items():
            summary = results['differences']['summary']
            identifier_data.append([
                sheet_name,
                ', '.join(summary.get('available_identifiers', [])),
                ', '.join(summary.get('missing_identifiers', [])),
                summary['total_rows_file1'],
                summary['total_rows_file2'],
                len(set([r['composite_key'] for r in results['differences']['rows_only_in_file1']])),
                len(set([r['composite_key'] for r in results['differences']['rows_only_in_file2']]))
            ])
        
        identifier_df = pd.DataFrame(identifier_data, columns=[
            'Sheet Name',
            'Available Identifiers',
            'Missing Identifiers',
            f'{file1_name} Records',
            f'{file2_name} Records',
            f'Unique to {file1_name}',
            f'Unique to {file2_name}'
        ])
        identifier_df.to_excel(writer, sheet_name='Identifier_Analysis', index=False)
        
        # Format identifier analysis sheet
        worksheet = writer.sheets['Identifier_Analysis']
        for col_num, value in enumerate(identifier_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # 5. SHEET-SPECIFIC DETAILS
        print(f"   📄 Creating sheet-specific detail sheets...")
        for sheet_name, results in comparison_results.items():
            # Create detailed sheet for each original sheet
            sheet_data = []
            
            # Add all modifications for this sheet
            if results['differences']['modified_rows']:
                for mod in results['differences']['modified_rows']:
                    for col, changes in mod['modified_columns'].items():
                        sheet_data.append([
                            'MODIFIED',
                            '',  # composite_key removed - now using individual identifiers
                            mod['file1_index'],
                            mod['file2_index'],
                            col,
                            str(changes['file1_value']),
                            str(changes['file2_value']),
                            'Data changed between files'
                        ])
            
            # Add records only in file 1
            for row in results['differences']['rows_only_in_file1']:
                sheet_data.append([
                    'ONLY_IN_FILE1',
                    '',  # composite_key removed - now using individual identifiers
                    row['file1_index'],
                    '',
                    'ALL_COLUMNS',
                    'N/A',
                    str(row['data']),
                    'Record exists only in File 1'
                ])
            
            # Add records only in file 2
            for row in results['differences']['rows_only_in_file2']:
                sheet_data.append([
                    'ONLY_IN_FILE2',
                    '',  # composite_key removed - now using individual identifiers
                    '',
                    row['file2_index'],
                    'ALL_COLUMNS',
                    str(row['data']),
                    'N/A',
                    'Record exists only in File 2'
                ])
            
            if sheet_data:
                # Get identifier columns for this sheet
                available_identifiers = results['differences']['summary']['available_identifiers']
                
                # Create data with individual identifier columns
                sheet_data_with_identifiers = []
                for item in sheet_data:
                    change_type = item[0]
                    row_data = [
                        change_type,
                        item[2],  # file1_index
                        item[3],  # file2_index
                        item[4],  # column
                        item[5],  # file1_value
                        item[6],  # file2_value
                        item[7]   # description
                    ]
                    
                    # Add identifier columns based on the type of change
                    if change_type == 'MODIFIED':
                        # Find the corresponding modification to get identifiers
                        for mod in results['differences']['modified_rows']:
                            if mod['file1_index'] == item[2] and mod['file2_index'] == item[3]:
                                for ident_col in available_identifiers:
                                    row_data.append(mod['identifier_values'].get(ident_col, ''))
                                break
                    elif change_type == 'ONLY_IN_FILE1':
                        # Find the corresponding row to get identifiers
                        for row in results['differences']['rows_only_in_file1']:
                            if row['file1_index'] == item[2]:
                                for ident_col in available_identifiers:
                                    row_data.append(row['identifier_values'].get(ident_col, ''))
                                break
                    elif change_type == 'ONLY_IN_FILE2':
                        # Find the corresponding row to get identifiers
                        for row in results['differences']['rows_only_in_file2']:
                            if row['file2_index'] == item[3]:
                                for ident_col in available_identifiers:
                                    row_data.append(row['identifier_values'].get(ident_col, ''))
                                break
                    
                    sheet_data_with_identifiers.append(row_data)
                
                # Create columns list
                columns = ['Change Type', f'{file1_name} Row', f'{file2_name} Row',
                          'Column', f'{file1_name} Value', f'{file2_name} Value', 'Description'] + available_identifiers
                
                sheet_df = pd.DataFrame(sheet_data_with_identifiers, columns=columns)
                
                # Limit sheet name for Excel compatibility
                safe_sheet_name = f"Details_{sheet_name}"[:31] if len(f"Details_{sheet_name}") > 31 else f"Details_{sheet_name}"
                sheet_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                
                # Format sheet
                worksheet = writer.sheets[safe_sheet_name]
                for col_num, value in enumerate(sheet_df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
        
        print(f"   ✅ Comprehensive Excel report created successfully!")
        print(f"   📁 File: {output_path}")
        print(f"   📋 Sheets created:")
        print(f"      • Summary - Overall comparison statistics")
        print(f"      • All_Modifications - All changed data across sheets")
        print(f"      • Identifier_Analysis - Unique identifier usage analysis")
        if unique_file1_data:
            print(f"      • Only_{file1_name} - Records unique to first file")
        if unique_file2_data:
            print(f"      • Only_{file2_name} - Records unique to second file")
        for sheet_name in comparison_results.keys():
            safe_name = f"Details_{sheet_name}"[:31]
            print(f"      • {safe_name} - Detailed changes for {sheet_name}")

if __name__ == "__main__":
    print("🚀 Starting Excel File Comparison...")
    
    # Define file paths
    file1_path = "Client-Final.xlsx"
    file2_path = "IT-FINAL.xlsx"
    output_path = "Excel_Comparison_Report.xlsx"
    
    # Run comparison
    results = compare_excel_files(file1_path, file2_path, output_path)
    
    if results:
        print(f"\n🎉 Comparison completed successfully!")
        print(f"📋 Check '{output_path}' for detailed differences.")
    else:
        print(f"\n❌ Comparison failed. Please check the error messages above.")