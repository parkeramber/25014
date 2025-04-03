import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import xlsxwriter
import re
import os
from datetime import datetime

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Define the relative path to your TSV file
file_path = os.path.join(script_dir, '25014 - Camera for Aerospace Situational Awareness - Team Tasks4_1_25.tsv')

# Load the TSV file into a DataFrame
df = pd.read_csv(file_path, sep='\t')

# Fill missing or empty Assignees with 'None'
df['Assignees'] = df['Assignees'].fillna('None')  # Replace NaN values with 'None'

# Convert the 'Start date' and 'End date' columns to datetime format
df['Start date'] = pd.to_datetime(df['Start date'], format='%b %d, %Y')
df['End date'] = pd.to_datetime(df['End date'], format='%b %d, %Y')

# Sort by 'End date' in ascending order
df = df.sort_values(by='End date')

# Create a duration column for plotting
df['Duration'] = (df['End date'] - df['Start date']).dt.days

# Rename Assignees based on mapping
assignee_mapping = {
    'aidancler24': 'Aidan',
    'DiegoG185593': 'Diego',
    'JohnDT-MechE': 'John',
    'carlywingness': 'Carly',
    'm-colson': 'Matthew',
    'parkeramber': 'Amber'
}
df['Assignees'] = df['Assignees'].replace(assignee_mapping, regex=True)

# Check if "Sprint" exists, if not rename "Iteration" to "Sprint"
if 'Iteration' in df.columns:
    df.rename(columns={'Iteration': 'Sprint'}, inplace=True)

# Add a new 'Corrective Action' column, which will be initially empty
df['Corrective Action'] = ''  # Or fill with placeholder text like "TBD"

# Define colors for each sprint
sprint_colors = {
    'Sprint 1': 'skyblue',
    'Sprint 2': 'lightgreen',
    'Sprint 3': 'lightcoral',
    'Sprint 4': 'plum',
    'Sprint 5': 'gold',
    'Sprint 6': 'lightseagreen'
}

# Get the current date and create a folder with the current date
current_date = datetime.now().strftime('%Y-%m-%d')
folder_name = f"export_{current_date}"
os.makedirs(folder_name, exist_ok=True)  # Create the folder if it doesn't exist

# Define the Excel file path inside the newly created folder
output_filename = os.path.join(folder_name, 'gantt_chart_export_with_sprint_tables.xlsx')

# Export to Excel and include all data in multiple sheets, one per sprint
writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')

# Access the workbook and define date format
workbook = writer.book
date_format = workbook.add_format({'num_format': 'mm/dd/yyyy'})  # Date format for date columns

# Define the border format for the outline (outer border only)
outer_border_format = workbook.add_format({'border': 1})  # Outer border only

# Define the highlight format for the total hours row
highlight_format = workbook.add_format({'bg_color': '#FFFF00', 'border': 1, 'bold': True})

# Rename the "Title" column to "Task"
df.rename(columns={'Title': 'Task'}, inplace=True)

# Loop through each sprint and create a separate sheet for each sprint
sprints = df['Sprint'].unique()

for sprint in sprints:
    df_sprint = df[df['Sprint'] == sprint]  # Filter tasks by sprint

    # Write task data to a new sheet for this sprint
    sheet_name = f'Sprint {sprint}'
    df_sprint[['Task', 'Assignees', 'Start date', 'End date', 'Sprint', 'Status', 'Hours Completed', 'Corrective Action']].to_excel(writer, sheet_name=sheet_name, index=False)
    
    # Access the worksheet for this sprint
    worksheet = writer.sheets[sheet_name]
    
    # Set column widths and apply the date format
    worksheet.set_column('A:A', 40)  # Task title column width
    worksheet.set_column('B:B', 40)  # Assignees column width
    worksheet.set_column('C:C', 25, date_format)  # Start date column width and date format
    worksheet.set_column('D:D', 25, date_format)  # End date column width and date format
    worksheet.set_column('E:E', 12)  # Sprint column width
    worksheet.set_column('F:F', 12)  # Status column width
    worksheet.set_column('G:G', 15)  # Hours completed column width
    worksheet.set_column('H:H', 30)  # Corrective Action column width

    # Apply the outer border around the table (calculate last row and column)
    rows, cols = df_sprint.shape
    last_row = rows + 1  # Add 1 because row numbering starts at 1 in Excel
    last_col = cols - 1  # Adjust for 0-based index

    # Apply the outer border to the range of the table
    worksheet.conditional_format(f'A1:H{last_row}', {'type': 'no_blanks', 'format': outer_border_format})
    worksheet.conditional_format(f'A1:H{last_row}', {'type': 'blanks', 'format': outer_border_format})

    # Highlight rows where the "Status" is "Todo"
    worksheet.conditional_format(f'A2:H{last_row}', {
        'type': 'formula',
        'criteria': '=$F2="Todo"',
        'format': workbook.add_format({'bg_color': '#D0D8B3'})  # Light fill color for "Todo" rows
    })

    # Calculate the total of "Hours Completed" for the current sprint
    hours_total = df_sprint['Hours Completed'].sum()

    # Write the total hours completed to a new row below the table
    total_row = rows + 1  # This should be the row after the last task row
    worksheet.write(total_row, 5, 'Total Hours', highlight_format)  # Write "Total Hours" in the Status column with highlighting
    worksheet.write(total_row, 6, hours_total, highlight_format)  # Write the total hours in the "Hours Completed" column with highlighting

    # Count statuses for this sprint
    status_counts = df_sprint['Status'].value_counts()
    todo_count = status_counts.get('Todo', 0)
    in_progress_count = status_counts.get('In Progress', 0)
    done_count = status_counts.get('Done', 0)

    # Add status counts for this sprint
    status_start_row = total_row + 2  # Leave one empty row after "Total Hours"
    worksheet.write(status_start_row, 5, 'TODO Count', highlight_format)
    worksheet.write(status_start_row, 6, todo_count, highlight_format)

    worksheet.write(status_start_row + 1, 5, 'In Progress Count', highlight_format)
    worksheet.write(status_start_row + 1, 6, in_progress_count, highlight_format)

    worksheet.write(status_start_row + 2, 5, 'Done Count', highlight_format)
    worksheet.write(status_start_row + 2, 6, done_count, highlight_format)

    # Create the Gantt chart for the specific sprint
    fig, ax = plt.subplots(figsize=(25, 15))

    # Loop through each task in the sprint and plot the Gantt chart
    for i, task in enumerate(df_sprint['Task']):
        start = df_sprint['Start date'].iloc[i]
        duration = df_sprint['Duration'].iloc[i]
        color = sprint_colors.get(sprint, 'gray')  # Default to gray if sprint color not found
        ax.barh(task, duration, left=start, color=color, edgecolor='black')

    # Formatting the x-axis for dates
    ax.xaxis_date()
    ax.xaxis.set_major_locator(mdates.AutoDateLocator())  # Automatically place date ticks
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))  # Format dates as YYYY-MM-DD
    plt.xticks(rotation=45, ha='right')  # Rotate x-axis labels for clarity

    # Adjust the limits to match the start and end date of the tasks with a small buffer
    start_date = df_sprint['Start date'].min()
    end_date = df_sprint['End date'].max()

    if start_date == end_date:
        start_date -= pd.Timedelta(days=1)
        end_date += pd.Timedelta(days=1)

    plt.xlim(start_date, end_date)

    # Labels and title
    plt.xlabel('Date')
    plt.ylabel('Tasks')
    plt.title(f'Gantt Chart - {sprint}')
    
    # Save the Gantt chart for this sprint as an image
    gantt_image = os.path.join(folder_name, f'gantt_chart_{sprint}.png')
    plt.tight_layout()
    plt.savefig(gantt_image)
    plt.close()

    # Insert the Gantt chart image into the Excel file (insert below the data)
    worksheet.insert_image(f'I2', gantt_image)

# Close the Excel writer to finalize the Excel file
# === Add Overall Status Summary to the Last Sheet ===
status_counts_overall = df['Status'].value_counts()
todo_total = status_counts_overall.get('Todo', 0)
in_progress_total = status_counts_overall.get('In Progress', 0)
done_total = status_counts_overall.get('Done', 0)

df_summary = pd.DataFrame({
    'Status': ['TODO', 'In Progress', 'Done'],
    'Count': [todo_total, in_progress_total, done_total]
})

df_summary.to_excel(writer, sheet_name='Overall Status Summary', index=False)

writer.close()

print(f"Excel file and images saved in folder: {folder_name}")