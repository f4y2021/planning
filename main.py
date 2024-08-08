import pandas as pd
import streamlit as st
from io import BytesIO
import xlsxwriter
file_path = st.file_uploader("Upload Excel File from Template")

run_button=st.button("Run")

if run_button:
    # Load the data with the correct header row, assuming headers are on row 5 (index 4)
    df = pd.read_excel(file_path, header=4, usecols="B,D:AB",sheet_name='Planning To Be Updated')
    st.dataframe(df)
    # Drop columns A and C (0-indexed, so we drop columns at positions 0 and 2)
    #df.drop(df.columns[[0, 2]], axis=1, inplace=True)
    
    # Set column B as the index
    df.set_index(df.columns[0], inplace=True)
    
    df = df[df.index.notnull()]
    
    task_indices = df.index[df.index.str.contains('Task')].tolist()
    
    wp_indices = df.index[df.index.str.contains('WP')].tolist()
    
    
    # Dictionary to store DataFrames for each task
    task_dfs = {}
    
    # Iterate over the task indices and slice the DataFrame
    for i, task_index in enumerate(task_indices):
        # Define the start and end indices for slicing
        start_idx = task_index
        end_idx = task_indices[i+1] if i+1 < len(task_indices) else None
        
        # Slice the DataFrame for the current task
        task_df = df.loc[start_idx:end_idx]
        # Removes first and last line, where the task would still be
        task_df = task_df.iloc[1:-1]
        
        #Resets the index, so that the Person name becomes a column instead of the index
        task_df = task_df.reset_index()
        
        task_df.rename(columns={task_df.columns[0]: 'Person'}, inplace=True)
        # Converts the person name into string (for the contains used next)
        task_df[task_df.columns[0]] = task_df[task_df.columns[0]].astype(str)
        
        # Removes rows where the WP still would be
        task_df = task_df[~task_df[task_df.columns[0]].str.contains('WP', na=False)]
        
        # Concatenates the dataframes for each task into a dictionary
        task_dfs[task_index] = task_df


    
    # Define additional fixed variables
    project = "Roketsan"
    subproject = "Tanks"
    task = None  # Will be set dynamically from task name
    subtask = None  # Set manually or dynamically if applicable / IGNORE FOR NOW
    description = "Description of the task"  # This can be specific or the same for all / NEEDS ADTIONAL DATA TO BE USED
    
    # Create a list to store all rows for the final DataFrame
    all_data = []
    
    # Loop through each task and DataFrame in the dictionary
    for task_name, df in task_dfs.items():
        for index, row in df.iterrows():
            for col in df.columns[1:]:  # Skip the 'Person' column
                month = pd.to_datetime(col).strftime('%Y-%m')
                effort = row[col]
                if pd.notna(effort):  # Consider only non-NaN values
                    # Extract task from the dictionary key
                    task = task_name.split(" - ")[0]
                    wp = int(task.split()[1].split(".")[0])
                    hours = effort * 160  # 160 hours/month
                    
                    # Prepare the row data
                    row_data = {
                        "Project": project,
                        #"SubProject": subproject,
                        "WP": wp,
                        "Task": task,
                        #"Subtask": subtask,
                        "Description": description,
                        "Person": row['Person'],
                        "Effort": effort,
                        "Hours": hours,
                        #Time": col,
                        "Month": month
                    }
                    
                    # Append the row data to the list
                    all_data.append(row_data)
    
    # Create a final DataFrame
    final_df = pd.DataFrame(all_data)
    
    st.dataframe(final_df)
    
    flnme = "Integrated_Data.xlsx"
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Integrated_Data')
    
    st.download_button(label="Download Integrated Data Excel workbook", data=buffer.getvalue(), file_name=flnme, mime="application/vnd.ms-excel")

