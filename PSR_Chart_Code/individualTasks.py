import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import xlsxwriter
import os
from datetime import datetime

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Define the relative path to your TSV file
file_path = os.path.join(script_dir, '25014 - Camera for Aerospace Situational Awareness - Team Tasks 09-26-2024.tsv')

# Load the TSV file into a DataFrame
df = pd.read_csv(file_path, sep='\t')

# Rename the "Title" column to "Task" if it exists
if 'Title' in df.columns:
    df.rename(columns={'Title': 'Task'}, inplace=True)

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

# Debugging print to check unique sprints
print("Unique Sprints in data:", df['Sprint'].unique())

# Add a new 'Corrective Action' column, which will be initially empty
df['Corrective Action'] = ''  # Or fill with placeholder text like "TBD"

# Define colors for each assignee
assignee_colors = {
    'Amber': 'lightblue',
    'Diego': 'lightgreen',
    'Carly': 'lightcoral',
    'Matthew': 'plum',
    'Aidan': 'gold',
    'John': 'lightseagreen',
    'Team': 'gray'
}

# Get the current date and create a folder with the current date
current_date = datetime.now().strftime('%Y-%m-%d')
folder_name = f"export_{current_date}"
os.makedirs(folder_name, exist_ok=True)  # Create the folder if it doesn't exist

# Define the Excel file path inside the newly created folder
output_filename = os.path.join(folder_name, 'gantt_chart_export_by_assignee_and_sprint.xlsx')

# Export to Excel and include all data in multiple sheets, one per assignee and sprint
writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')

# Access the workbook and define date format
workbook = writer.book
date_format = workbook.add_format({'num_format': 'mm/dd/yyyy'})  # Date format for date columns

# Define the border format for the outline (outer border only)
outer_border_format = workbook.add_format({'border': 1})  # Outer border only

# Define a yellow background format with a border for total hours
highlight_format = workbook.add_format({'bg_color': '#FFFF00', 'border': 1, 'bold': True})

# Define a list of team members including "Team" for tasks not assigned to a specific person
team_members = ['Team', 'Amber', 'Diego', 'Carly', 'Matthew', 'Aidan', 'John']

# Fill missing or empty Assignees with 'Team' to group unassigned tasks
df['Assignees'] = df['Assignees'].replace('None', 'Team')

# Get the unique sprints
sprints = df['Sprint'].unique()

# Loop through each team member and sprint combination
for member in team_members:
    for sprint in sprints:
        # Filter tasks where the member's name is part of the 'Assignees' column (can have multiple people)
        df_member_sprint = df[df['Assignees'].str.contains(member) & (df['Sprint'] == sprint)]
        
        if df_member_sprint.empty:
            print(f"No tasks found for {member} in {sprint}")
            continue
        else:
            print(f"Processing tasks for {member} in {sprint}")

        # Write task data to a new sheet for this member and sprint
        sheet_name = f'{member}_Sprint_{sprint}'
        df_member_sprint[['Task', 'Assignees', 'Start date', 'End date', 'Sprint', 'Status', 'Hours Completed', 'Corrective Action']].to_excel(writer, sheet_name=sheet_name, index=False)

        # Access the worksheet for this member and sprint
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
        rows, cols = df_member_sprint.shape
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

        # Calculate the total of "Hours Completed" for the current member and sprint
        hours_total = df_member_sprint['Hours Completed'].sum()

        # Write the total hours completed to a new row below the table
        total_row = rows + 1  # This should be the row after the last task row
        worksheet.write(total_row, 5, 'Total Hours', highlight_format)  # Write "Total Hours" in the Status column with highlighting
        worksheet.write(total_row, 6, hours_total, highlight_format)  # Write the total hours in the "Hours Completed" column with highlighting

        # Create the Gantt chart for the specific member and sprint
        fig, ax = plt.subplots(figsize=(14, 10))

        # Loop through each task for the member and plot the Gantt chart
        for i, task in enumerate(df_member_sprint['Task']):
            start = df_member_sprint['Start date'].iloc[i]
            duration = df_member_sprint['Duration'].iloc[i]
            color = assignee_colors.get(member, 'gray')  # Default to gray if color not found
            ax.barh(task, duration, left=start, color=color, edgecolor='black')

        # Formatting the x-axis for dates
        ax.xaxis_date()
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())  # Automatically place date ticks
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))  # Format dates as YYYY-MM-DD
        plt.xticks(rotation=45, ha='right')  # Rotate x-axis labels for clarity

        # Adjust the limits to match the start and end date of the tasks with a small buffer
        start_date = df_member_sprint['Start date'].min()
        end_date = df_member_sprint['End date'].max()

        if start_date == end_date:
            start_date -= pd.Timedelta(days=1)
            end_date += pd.Timedelta(days=1)

        # Set the x-axis limits to match the start and end dates with the adjusted buffer
        plt.xlim(start_date, end_date)

        # Labels and title
        plt.xlabel('Date')
        plt.ylabel('Tasks')
        plt.title(f'Gantt Chart - {member} - Sprint {sprint}')

        # Save the Gantt chart for this member and sprint as an image
        gantt_image = os.path.join(folder_name, f'gantt_chart_{member}_Sprint_{sprint}.png')
        plt.tight_layout()
        plt.savefig(gantt_image)
        plt.close()

        # Insert the Gantt chart image into the Excel file (insert below the data)
        worksheet.insert_image(f'I2', gantt_image)

# Close the Excel writer to finalize the Excel file
writer.close()

print(f"Excel file and images saved in folder: {folder_name}")

