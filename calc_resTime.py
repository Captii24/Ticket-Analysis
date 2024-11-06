import pandas as pd

# Load the ticket data
df = pd.read_excel("Ticket_Log.xlsx")

# Convert dates to datetime for calculations
df['Date Submitted'] = pd.to_datetime(df['Date Submitted'], errors='coerce')
df['Date Resolved'] = pd.to_datetime(df['Date Resolved'], errors='coerce')

# Calculate resolution time in hours
df['Resolution Time (hours)'] = (df['Date Resolved'] - df['Date Submitted']).dt.total_seconds() / 3600

# Filter for resolved tickets only
resolved_df = df[df['Status'] == 'Resolved']

# Calculate average resolution time by Priority and Status
avg_resolution_time = resolved_df.groupby(['Priority', 'Status'])['Resolution Time (hours)'].mean().reset_index()
avg_resolution_time.to_excel("Average_Resolution_Time.xlsx", index=False)

print("Average resolution times have been saved to 'Average_Resolution_Time.xlsx'")
