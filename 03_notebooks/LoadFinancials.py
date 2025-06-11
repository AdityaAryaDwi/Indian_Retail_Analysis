import pandas as pd
import os

def comp_indv_statements(raw_files,statements):
    """
    Load and process financial statements for multiple companies from Excel files.

    For each file in the list of raw_files, this function extracts and cleans
    the requested financial statements, organizing them in a structured dictionary.

    Args:
        raw_files (list): Filenames of Excel files containing raw financial data.
        statements (list): List of financial statement titles to extract (e.g., 'PROFIT & LOSS').

    Returns:
        dict: Dictionary with company names as keys and their financial statements as nested dictionaries.
    """
    financials={}
    for company in raw_files:
        comp_initials=company.split(".")[0]
        print(f"Financial Statements for {comp_initials} loading...")
        financials[comp_initials]=load_financials(company,statements)
        print(f"Financials Statements for {comp_initials} loaded.")
    return financials

def load_financials(file_name,target_financials):
    """
    Open a company's Excel file and load its financial data sheet.

    Reads the raw 'Data Sheet' from the file and passes it along to be filtered
    for specific statements.

    Args:
        file_name (str): The name of the Excel file to process.
        target_financials (list or str): Financial statements to extract.

    Returns:
        dict or pd.DataFrame: Extracted statements in a dictionary or single DataFrame.
    """
    file_path=os.path.join("..","01_data_raw","screener_data",file_name)
    xls=pd.ExcelFile(file_path)
    target_sheet="Data Sheet"
    raw_df=pd.read_excel(file_path,target_sheet,header=None)
    return fetch_statements(raw_df,target_financials)

def fetch_statements(raw_df,financials):
    """
    Retrieve one or multiple financial statements from raw data.

    Decides whether to fetch a single statement or multiple and returns
    them accordingly.

    Args:
        raw_df (pd.DataFrame): Raw data read from Excel.
        financials (list or str): Statement title(s) to extract.

    Returns:
        dict or pd.DataFrame: Extracted DataFrame(s) for the requested statement(s).
    """
    if isinstance(financials,list):
        dfs={}
        for title in financials:
            df=load_statement(raw_df,title)
            dfs[title]=df
        return dfs
    else:
        return load_statement(raw_df,financials)

def load_statement(raw_df,title):
    """
    Extract and clean a specific financial statement from raw Excel data.

    Searches the sheet for the given title, extracts the relevant block,
    formats headers and fiscal columns, and returns a tidy DataFrame.

    Args:
        raw_df (pd.DataFrame): Full raw data as read from Excel.
        title (str): Title of the financial statement to extract.

    Returns:
        pd.DataFrame or None: Cleaned DataFrame, or None if extraction failed.
    """
    title_idx=raw_df[raw_df[0]==title.upper()].index[0]
    header_idx=title_idx+1
    columns=raw_df.iloc[header_idx]
    data_block=raw_df.iloc[header_idx+1:]
    final_row_idx=data_block[data_block.isnull().all(axis=1)].index[0]
    df=data_block.loc[:final_row_idx-1]
    df.columns=columns
    df.columns.name=None
    df=df.dropna(axis=1,how="all")
    df.columns=["Particulars"]+[pd.to_datetime(col).strftime("%b-%y") for col in df.columns[1:]]
    df.reset_index(drop=True,inplace=True)
    return df

def common_fiscal_years(financials,statements):
    """
    Determine which fiscal years are common to all companies for a given statement.

    Ensures consistency when combining financial data by identifying overlapping
    periods across companies.

    Args:
        financials (dict): Dictionary of financial statements per company.
        statements (list): List of statements, using the first as reference.

    Returns:
        list: Columns to use for master tables, including shared fiscal years.
    """
    initial_cols=set(financials[list(financials.keys())[0]].get(statements[0]).columns)
    for company in list(financials.keys())[1:]:
        current_cols=set(financials[company].get(statements[0]).columns)
        initial_cols.intersection_update(current_cols)
    initial_cols.remove("Particulars")
    return ["Particulars","Company"]+sorted(list(initial_cols))

def master_table(financials,statements):
    """
    Generate long-format master DataFrames for each statement type.

    Combines all companies' data, filters to common fiscal years, and reshapes
    into a tidy format for analysis.

    Args:
        financials (dict): Dictionary of all companies and their statements.
        statements (list): List of statement types to process.

    Returns:
        dict: Dictionary with statement titles as keys and tidy DataFrames as values.
    """
    final_cols=common_fiscal_years(financials,statements)
    statements_dict={}
    for i in range(len(statements)):
        dfs_list=[]
        for company in financials.keys():
            df=financials[company].get(statements[i])
            df["Company"]=company
            dfs_list.append(df)
        master_df=pd.concat(dfs_list,ignore_index=True)
        master_df=master_df[final_cols]
        master_df=master_df.melt(
            id_vars=["Company","Particulars"],
            var_name="Fiscal Year End",
            value_name="Value"
        )
        statements_dict[statements[i]]=master_df
    return statements_dict