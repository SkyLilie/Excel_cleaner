import pandas as pd
import os


def remove_empty_rows(df: pd.DataFrame) -> pd.DataFrame:

    print("\n--- Cleaning sheet ---")
    original_row_count = len(df)
    cleaned_df = df.dropna(how='all')
    rows_dropped = original_row_count - len(cleaned_df)

    if rows_dropped > 0:
        print(f"  - Found and removed {rows_dropped} completely empty row(s).")
    else:
        print("  - No completely empty rows found.")

    return cleaned_df.reset_index(drop=True)

def process_column_a_duplicates(df: pd.DataFrame, col_A: str, col_B: str) -> pd.DataFrame:

    print(f"\n--- Processing Duplicates in {col_A} ---")
    duplicates_in_a = df[df.duplicated(subset=[col_A], keep=False)].copy()
    grouped = duplicates_in_a.groupby(col_A)
    indices_to_drop = []

    if grouped.ngroups == 0:
        print(f"No duplicates found in {col_A}.")
        return df

    print(f"Found {grouped.ngroups} group(s) of duplicates in {col_A}.")

    for name, group in grouped:
        print(f"\nProcessing duplicates for value: '{name}' in {col_A}")

        if group[col_B].nunique() == 1:
            print(f"  - Values in {col_B} are identical. Keeping one and deleting others.")
            indices_to_drop.extend(group.index[1:])
        else:
            print(f"  - CONFLICT: Values in {col_B} are different. Please choose which row to KEEP.")
            print(group.to_string())

            valid_indices = group.index.tolist()
            chosen_index = None

            while chosen_index not in valid_indices:
                try:
                    choice = int(input(f"  Enter the index of the row to KEEP {valid_indices}: "))
                    if choice in valid_indices:
                        chosen_index = choice
                    else:
                        print("  Invalid index. Please choose one of the available indices.")
                except ValueError:
                    print("  Invalid input. Please enter a number.")

            for idx in valid_indices:
                if idx != chosen_index:
                    indices_to_drop.append(idx)
            print(f"  - Keeping index {chosen_index} and marking others for deletion.")

    if indices_to_drop:
        print(f"\nDropping {len(indices_to_drop)} redundant row(s) based on Column A analysis.")
        df = df.drop(indices_to_drop)

    print("--- Finished Function 1 ---")
    return df.reset_index(drop=True)

def process_column_b_duplicates(df: pd.DataFrame, col_B: str) -> pd.DataFrame:

    print(f"\n--- Processing Duplicates in {col_B} ---")

    while True:
        duplicates_in_b = df[df.duplicated(subset=[col_B], keep=False)]

        if duplicates_in_b.empty:
            print(f" No duplicates found in {col_B}.")
            break

        first_dup_value = duplicates_in_b[col_B].iloc[0]
        group = df[df[col_B] == first_dup_value]

        print(f"\nFound duplicate value '{first_dup_value}' in {col_B} at the following indices:")
        print(group.to_string())

        user_choice = ''
        while user_choice not in ['yes', 'no']:
            user_choice = input("Do you want to provide a new value for one of these? (yes/no): ").lower().strip()

        if user_choice == 'yes':
            valid_indices = group.index.tolist()
            chosen_index = None

            while chosen_index not in valid_indices:
                try:
                    choice = int(input(f"  Enter the index of the row to CHANGE {valid_indices}: "))
                    if choice in valid_indices:
                        chosen_index = choice
                    else:
                        print("  Invalid index. Please choose one of the available indices.")
                except ValueError:
                    print("  Invalid input. Please enter a number.")

            new_value = input(f"  Enter the new value for Column B at index {chosen_index}: ")
            df.loc[chosen_index, col_B] = new_value
            print(f"  - Updated index {chosen_index} with new value '{new_value}'.")
        else:
            print("  - Skipping this group. The duplicates will remain.")
            print(f"  - NOTE: To prevent an infinite loop, we will stop checking {col_B}.")
            print(f"    Re-run the script if you want to process other duplicates in {col_B}.")
            break

    return df

def main():

    print("Excel Sheet Cleaner Utility")
    print("=" * 50)

    while True:
        input_file_path = input("Enter the path to your Excel file: ")
        if os.path.exists(input_file_path):
            break
        print(f"Error: File not found at '{input_file_path}'. Please check the path and try again.")

    try:
        xls = pd.ExcelFile(input_file_path)
        sheet_names = xls.sheet_names
        if len(sheet_names) == 1:
            sheet_name = sheet_names[0]
            print(f"Automatically selected sheet: '{sheet_name}'")
        else:
            print("Available sheets:", sheet_names)
            sheet_name = input("Which sheet would you like to process? ")
            if sheet_name not in sheet_names:
                print("Error: Sheet not found.")
                return
    except Exception as e:
        print(f"Could not read sheets from Excel file. Error: {e}")
        return

    output_file = 'cleaned_' + os.path.basename(input_file_path)

    try:
        df = pd.read_excel(input_file_path, sheet_name=sheet_name)
        col_A = input("Enter the name of first column of your Excel file: ").strip()
        col_B = input("Enter the name of second column of your Excel file: ").strip()
        if col_A not in df.columns or col_B not in df.columns:
            print(f"\nError: The Excel sheet must contain {col_A} and {col_B}.")
            print("Please rename your columns and re-run the script.")
            return

    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return

    print("\nOriginal Data:")
    print(df)


#Clean up completely empty rows first.
    df_after_empty_removal = remove_empty_rows(df.copy())

#Process Column A duplicates on the result of the previous step.
    df_after_a = process_column_a_duplicates(df_after_empty_removal.copy(), col_A, col_B)

#Process Column B duplicates on the result of the first function.
    final_df = process_column_b_duplicates(df_after_a.copy(), col_B)

    try:
        final_df.to_excel(output_file, index=False)
        print(f"\n✅ Success! The cleaned data has been saved to '{output_file}'")
    except Exception as e:
        print(f"\n❌ Error: Could not save the file. Reason: {e}")

if __name__ == "__main__":
    main()