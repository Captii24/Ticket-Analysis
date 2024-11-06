import pandas as pd
# Load data
df = pd.read_excel("Ticket_Log.xlsx")

# Convert date column to datetime
df['Date Submitted'] = pd.to_datetime(df['Date Submitted'], errors='coerce')

# Group by day, week, or month for trend analysis
daily_trends = df.groupby(df['Date Submitted'].dt.date).size().reset_index(name='Ticket Count')
daily_trends.to_excel("Daily_Ticket_Trends.xlsx", index=False)

print("Daily ticket trends have been saved to 'Daily_Ticket_Trends.xlsx'")
