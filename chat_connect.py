import pandas as pd
import os

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
    print("fnmdjklsn")
    if not email_file or not history_file:
        raise FileNotFoundError("Required files not found in the current directory.")

    emails_df = pd.read_csv(email_file)
    history_df = pd.read_csv(history_file)

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


def create_new_combinations(emails_list: list, history_list: list) -> list:
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
