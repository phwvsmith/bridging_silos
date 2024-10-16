import random
from typing import List, Set
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font
from datetime import datetime

def read_csv_files() -> (pd.DataFrame, pd.DataFrame):
    """
    Reads two CSV files from the current directory into pandas DataFrames.
    One file contains 'emails' in its name, and the other contains 'history'.

    Returns:
        tuple of pandas.DataFrame: Returns two DataFrames, one for emails and one for history.
    """

    files = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith('.csv')]
    email_file = next((f for f in files if 'emails' in f), None)
    history_file = next((f for f in files if 'history' in f), None)
    if not email_file or not history_file:
        raise FileNotFoundError("Required files not found in the current directory.")

    emails_df = pd.read_csv(email_file)
    history_df = pd.read_csv(history_file)
    print("fnmdjklsn")
    return emails_df, history_df



def rename_columns(emails: pd.DataFrame, history: pd.DataFrame) -> (pd.DataFrame, pd.DataFrame):
    """
    Renames the single column in the 'emails' and 'history' dataframes, if each dataframe has exactly one column.
    The column in 'emails' is renamed to 'colleague_emails' and in 'history' to 'combination_history_1','combination_history_2'.

    Args:
        emails (pd.DataFrame): DataFrame with email data.
        history (pd.DataFrame): DataFrame with history data.

    Returns:
        tuple of pd.DataFrame: The modified 'emails' and 'history' DataFrames.

    Raises:
        ValueError: If either 'emails' or 'history' does not have exactly one column.
    """

    if len(emails.columns) != 1 or len(history.columns) != 2:
        raise ValueError("One of the DataFrames does not have exactly one column.")

    emails.columns = ['colleague_emails']
    history.columns = ['combination_history_1','combination_history_2']

    return emails, history



def dataframes_to_lists(emails_df: pd.DataFrame, history_df: pd.DataFrame) -> (list, list):
    """
    Converts the single column of 'emails' dataframe into a list and
    each row of the 'history' dataframe into a set, then into a list of sets.

    Args:
        emails_df (pd.DataFrame): DataFrame containing the emails data.
        history_df (pd.DataFrame): DataFrame containing the history data.

    Returns:
        tuple of list, list of sets: A list for emails and a list of sets for history.
    """
    emails_list = emails_df.iloc[:, 0].tolist()
    history_list = [set(row) for row in history_df.itertuples(index=False, name=None)]

    return emails_list, history_list


def all_combinations(emails_list: list, history_list: list) -> list:
    """
    Creates all possible new combinations from the emails list as sets, 
    ensuring that the set does not already exist in the history list.

    Args:
        emails_list (list): List containing email addresses.
        history_list (list): List containing previously formed combinations as sets.

    Returns:
        list: A list of all new combination sets that are not present in the history list.
    """
    new_combinations = []
    print("fndjksnfjds")
    for i in range(len(emails_list)):
        for j in range(i + 1, len(emails_list)):
            email1, email2 = emails_list[i], emails_list[j]
            combination = {email1, email2}

            # Add combination if it's not in history
            if combination not in history_list:
                new_combinations.append(combination)

    if not new_combinations:
        print("No new unique combinations exist")
        return []

    return new_combinations

"""
chatgpt promt for next function:
Could you please create a function called create_new_combinations, that
takes two lists: emails_list which is a list of strings and history_list
which is a list of sets with two strings in each set. It reacts an empty
list called repeat_check. It creates another empty list called new_combinations.
It randomly selects two different string from emails_list and puts them
within a new temporary set set. It then checks if that set already exist
within history_list and if its not in history_list then it check that each
item within the set is not in repeat_check, and if not it adds the set
to new_combinations, each item to repeat_check and removes each item from
emails_list and repeats the process from randomly selecting strings
until emails_list is empty

Second prompt due to being stuck in while loop: 
At the moment create_new_combinations gets stuck in a while, can you suggest a fix

Third prompt to add formatted excel output:
Update the function so that it outputs new_combinations as a excel file
with the file name as batch and time and date
"""

def create_new_combinations(emails_list: List[str], history_list: List[Set[str]]) -> List[Set[str]]:
    """
    Generate new combinations of email pairs from a given list of emails.

    This function takes a list of email addresses and a history list containing sets of email pairs.
    It creates unique pairs of emails that do not exist in the history list, and ensures that each
    email is only used once in the new combinations. The process continues until there are no more
    emails left to pair or it's impossible to form more pairs.

    Parameters:
    emails_list (List[str]): A list of email addresses as strings.
    history_list (List[Set[str]]): A list of sets, with each set containing two email addresses.

    Returns:
    List[Set[str]]: A list of sets, with each set containing a unique pair of email addresses.

    Note:
    - The function uses random selection, so the output may vary each time it's called.
    - If the number of emails is odd, one email might remain unused.
    """
    repeat_check = []
    new_combinations = []

    # Track attempts to avoid infinite loop
    attempts = 0
    max_attempts = len(emails_list) * 3  # Arbitrary number, adjust as needed

    while emails_list and attempts < max_attempts:
        # if len(emails_list) < 2:
        #     break

        email1, email2 = random.sample(emails_list, 2)
        new_set = {email1, email2}

        if new_set not in history_list and email1 not in repeat_check and email2 not in repeat_check:
            new_combinations.append(new_set)
            repeat_check.extend(new_set)
            emails_list.remove(email1)
            emails_list.remove(email2)
            attempts = 0  # Reset attempts after a successful pairing
        else:
            attempts += 1  # Increment attempts if pairing wasn't successful


    # Convert the list of sets to a DataFrame
    df = pd.DataFrame(new_combinations, columns=['Email 1', 'Email 2'])

    # Generate a filename with the current date and time
    filename = f"batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    # Save to Excel
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    # Open the Excel file with openpyxl to apply styles
    workbook = load_workbook(filename)
    sheet = workbook.active

    # Apply red font to 'Email 1' column (assuming it's the first column)
    for row in range(2, sheet.max_row + 1):  # start from 2 to skip the header
        cell = sheet.cell(row, 1)
        cell.font = Font(color="FF0000")  # Red color

    # Save the changes
    workbook.save(filename)

    print(f"Saved to {filename}")


    return new_combinations

def update_history_and_save_csv(new_combinations: list, history_list: list):
    """
    Updates the history list with new combinations, converts it to a DataFrame, 
    and writes it to a CSV file, overwriting the original 'history' CSV file.

    Args:
        new_combinations (list): List of new combinations to add to history.
        history_list (list): The existing history list.

    Returns:
        None
    """
    # Update history list if there are new combinations
    if new_combinations:
        history_list.extend(new_combinations)

    # Convert the updated history list to a DataFrame
    history_df = pd.DataFrame(history_list, columns=['combination_history_1','combination_history_2'])

    # Find the original CSV file with 'history' in its name
    history_file = next((f for f in os.listdir('.') if os.path.isfile(f) and 'history' in f and f.endswith('.csv')), None)
    if not history_file:
        raise FileNotFoundError("History CSV file not found.")

    # Write the DataFrame to CSV, overwriting the original file
    history_df.to_csv(history_file, index=False)


# Example usage
# Assuming emails_df and history_df are the DataFrames read from the CSV files and possibly renamed
emails_df, history_df = read_csv_files()  # or use your existing DataFrames
emails_df, history_df = rename_columns(emails_df, history_df)  # Optional if columns are already renamed
emails_list, history_list = dataframes_to_lists(emails_df, history_df)

new_combinations = create_new_combinations(emails_list, history_list)

print(new_combinations)

print(emails_list)  # Prints the list of emails
print(history_list)  # Prints the list of history items


update_history_and_save_csv(new_combinations, history_list)
