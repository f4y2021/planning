import pandas as pd
import streamlit as st
from io import BytesIO
import xlsxwriter
import matplotlib.pyplot as plt
import seaborn as sns


st.set_page_config(
    page_title="Integrated Planning",
    page_icon="📈",
    layout="wide",
)



def load_data(file_obj, sheet_name):
    """Load the Excel file and return the DataFrame after initial processing."""
    df = pd.read_excel(file_obj, header=5, usecols="B,D:AS", sheet_name=sheet_name)
    
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
                month = pd.to_datetime(col, errors='coerce').to_period('M') if pd.notna(col) else None
                effort = row[col]
                if pd.notna(effort) and month:
                    task = task_name.split(" - ")[0]
                    wp = int(task.split()[1].split(".")[0])
                    hours = effort * 165  # 165 hours/month
                    row_data = {
                        "Project": project,
                        "WP": wp,
                        "Task": task,
                        "Person": row['Person'],
                        "Effort": effort,
                        "Hours": hours,
                        "Month": month
                    }
                    
                    all_data.append(row_data)
    
    final_df = pd.DataFrame(all_data)
    return final_df

def plot_pie_chart(data, label):
    """Plot a pie chart for the given data."""
    fig, ax = plt.subplots()
    sns.set(font_scale=0.5)
    data.plot.pie(autopct='%1.1f%%', ax=ax, startangle=90)
    ax.set_ylabel('')
    ax.set_title(f'Total Hours Distribution by {label}')
    return fig

def plot_heatmap(data, label):
    """Plot a heatmap for the given data."""
    fig, ax = plt.subplots()
    sns.set(font_scale=0.3)
    sns.heatmap(data, annot=True, fmt='g', cmap='YlGnBu', ax=ax)
    ax.set_title(f'Hours Heatmap by {label}')
    return fig

def visualize_data(final_df):
    """Visualize the final DataFrame with graphs and tables."""

    st.subheader("Graphical Analysis")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Hours Distribution by Task and Month")
        effort_pivot = final_df.pivot_table(values='Hours', index='Month', columns='Task', aggfunc='sum')
        st.bar_chart(effort_pivot)

    with col2:
        st.subheader("Hours Distribution by Work Package (WP)")
        wp_pivot = final_df.pivot_table(values='Hours', index='Month', columns='WP', aggfunc='sum')
        st.bar_chart(wp_pivot)

    with col1:
        st.subheader("Total Hours Distribution by Task")
        total_effort_by_task = final_df.groupby('Task')['Hours'].sum()
        st.pyplot(plot_pie_chart(total_effort_by_task, 'Task'))
    with col2:
        st.subheader("Total Hours Distribution by WP")
        total_effort_by_wp = final_df.groupby('WP')['Hours'].sum()
        st.pyplot(plot_pie_chart(total_effort_by_wp, 'WP'))
    
    with col1:
        st.subheader("Hours Heatmap by Task and Month")
        heatmap_data_task = final_df.pivot_table(values='Hours', index='Task', columns='Month', aggfunc='sum')
        st.pyplot(plot_heatmap(heatmap_data_task, 'Task'))
    with col2:
        st.subheader("Hours Heatmap by WP and Month")
        heatmap_data_wp = final_df.pivot_table(values='Hours', index='WP', columns='Month', aggfunc='sum')
        st.pyplot(plot_heatmap(heatmap_data_wp, 'WP'))

def main(file_objs, sheet_name, description):
    """Main function to orchestrate the process."""
    all_dfs = []
    project = None

    for file_obj in file_objs:
        df, project = load_data(file_obj, sheet_name)
        df = filter_before_equipa(df)
        task_dfs = extract_task_dfs(df)
        final_df = prepare_final_df(task_dfs, project, description)
        all_dfs.append(final_df)

    # Concatenate all DataFrames
    concatenated_df = pd.concat(all_dfs, ignore_index=True)
    
    return concatenated_df

# Streamlit UI
#st.title("Task Data Processor")

colu1, colu2, colu3 = st.columns([1, 2, 1])
colu2.image('logo_400.png')

file_objs = st.file_uploader("Upload Excel Files from Template", type=["xlsx"], accept_multiple_files=True)

sheet_name = 'Planning To Be Updated'
description = "Description of the task"

if st.button("Run", type="primary") and file_objs:
    with st.spinner("Processing..."):
        try:
            final_df = main(file_objs, sheet_name, description)
            
            # Prepare the Excel file for download
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, sheet_name='Integrated_Data', index=False)
            
            st.success("Processing complete!")
            st.download_button(
                label="Download Integrated Data Excel workbook",
                data=buffer.getvalue(),
                file_name="tasks_output.xlsx",
                mime="application/vnd.ms-excel"
            )

            # Visualize the data
            visualize_data(final_df)

        except Exception as e:
            st.error(f"An error occurred: {e}")
else:
    st.warning("Please upload Excel files")

