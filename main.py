import pandas as pd
import streamlit as st
from io import BytesIO
import xlsxwriter


def load_data(file_path, sheet_name):
    """Load the Excel file and return the DataFrame after initial processing."""
    df = pd.read_excel(file_path, header=5, usecols="B,D:AS", sheet_name=sheet_name)
    
    project = df.columns[0]
    df.set_index(df.columns[0], inplace=True)
    df = df[df.index.notnull()]
    df.index = df.index.astype(str).str.strip()
    
    return df, project

def filter_before_equipa(df):
    """Filter the DataFrame to remove rows after 'Equipa'."""
    if 'Equipa' in df.index:
        index_equipa = df.index.get_loc('Equipa')
        df = df.iloc[:index_equipa]
    else:
        raise ValueError("ERROR: 'Equipa' not found in the index.")
    return df

def extract_task_dfs(df):
    """Extract DataFrames for each task and store them in a dictionary."""
    task_indices = df.index[df.index.str.contains('Task')].tolist()
    
    task_dfs = {}
    for i, task_index in enumerate(task_indices):
        start_idx = task_index
        end_idx = task_indices[i+1] if i+1 < len(task_indices) else None
        
        task_df = df.loc[start_idx:end_idx]
        if i+1 < len(task_indices):
            task_df = task_df.iloc[1:-1]
        else:
            task_df = task_df.iloc[1:]
        
        task_df = task_df.reset_index()
        task_df.rename(columns={task_df.columns[0]: 'Person'}, inplace=True)
        task_df['Person'] = task_df['Person'].astype(str)
        task_df = task_df[~task_df['Person'].str.contains('WP', na=False)]
        
        task_dfs[task_index] = task_df
    
    return task_dfs

def prepare_final_df(task_dfs, project, description):
    """Prepare the final DataFrame combining all tasks' data."""
    all_data = []
    
    for task_name, df in task_dfs.items():
        for index, row in df.iterrows():
            for col in df.columns[1:]:  # Skip the 'Person' column
                month = pd.to_datetime(col).strftime('%Y-%m')
                effort = row[col]
                if pd.notna(effort):
                    task = task_name.split(" - ")[0]
                    wp = int(task.split()[1].split(".")[0])
                    
                    row_data = {
                        "Project": project,
                        "WP": wp,
                        "Task": task,
                        #"Description": description,
                        "Person": row['Person'],
                        "Effort": effort,
                        "Month": month
                    }
                    
                    all_data.append(row_data)
    
    final_df = pd.DataFrame(all_data)
    return final_df

def main(file_path, sheet_name, description, output_file):
    """Main function to orchestrate the process."""
    df, project = load_data(file_path, sheet_name)
    df = filter_before_equipa(df)
    task_dfs = extract_task_dfs(df)
    final_df = prepare_final_df(task_dfs, project, description)
    
    final_df.to_excel(output_file, index=False)
    print(final_df)
    
description = "Description of the task"  # Replace with actual description if needed
sheet_name = 'Planning To Be Updated'
flnme = "tasks_output.xlsx"

file_path = st.file_uploader("Upload Excel File from Template")

run_button=st.button("Run")

if run_button:
    main(file_path, sheet_name, description, output_file)


buffer = BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    final_df.to_excel(writer, sheet_name='Integrated_Data')

st.download_button(label="Download Integrated Data Excel workbook", data=buffer.getvalue(), file_name=flnme, mime="application/vnd.ms-excel")

