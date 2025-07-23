#Before running include these libraries
#pip install pandas seaborn matplotlib plotly


#pip install pandas seaborn matplotlib plotly

import pandas as pd
import plotly.express as px
import numpy as np # For numerical operations, especially mean

# IMPORTANT: You must ensure this CSV file is accessible.
# If running locally, make sure the path is correct.
# If you've uploaded it, please specify the exact filename (e.g., 'your_data.csv').
try:
    df = pd.read_csv(r'C:\Users\4483\N9340 Scans\No_header\Mashed-20250702-110847.csv')
except FileNotFoundError:
    print("ERROR: CSV file not found! Please ensure the path is correct or upload the file.")
    # Exit or handle the error appropriately if running non-interactively
    exit()

# --- IMPORTANT: Inspect your DataFrame and identify columns ---
print("DataFrame Head (first 5 rows):")
print(df.head())
print("\nDataFrame Columns (and their order):")
print(df.columns.tolist()) # Shows columns as a list for easy viewing

# Assign column names:
# Assume the first column is 'Frequency'.
frequency_column_name = df.columns[0]

# All other columns are assumed to be amplitude measurements.
# We will exclude the frequency column when calculating the average.
amplitude_columns = df.columns[1:].tolist()

# --- Data Cleaning and Type Conversion ---
# Convert all amplitude columns to numeric, coercing any non-numeric values to NaN.
# Plotly will automatically skip NaN values, creating gaps in the lines.
for col in amplitude_columns:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# --- Calculate the Average Amplitude ---
# Calculate the mean across the amplitude columns for each row.
# `numeric_only=True` is a safeguard, though we've already coerced to numeric.
df['Average Amplitude'] = df[amplitude_columns].mean(axis=1, numeric_only=True)

# --- Prepare Columns for Plotly ---
# The 'y' parameter in px.line can accept a list of column names.
# We'll include all original amplitude columns plus the new 'Average Amplitude' column.
columns_to_plot = amplitude_columns + ['Average Amplitude']




# --- Create Interactive Plot with Plotly Express ---
fig = px.line(df,
              x=frequency_column_name,
              y=columns_to_plot, # Plot all specified amplitude columns and the average
              title="Amplitude (dBM) vs Frequency (MHz) for Multiple Runs and Average",
              labels={frequency_column_name: "Frequency (MHz)", "value": "Amplitude (dBM)"}, # Custom labels for hover & axis
              hover_name="variable" # Shows the column name in the hover tooltip
             )

# --- Set X-axis (Frequency) to Logarithmic Scale ---
fig.update_xaxes(type="log",
                 title="Frequency (MHz)",
                 showgrid=True, gridwidth=1,
                 tickformat = None) # Setting tickformat to None or "" often helps prevent scientific notation

# To ensure non-scientific labels on a log axis, you might need more specific tickmode
# if default 'None' isn't sufficient for your specific data range.
# For example, to show specific powers of 10:
# fig.update_xaxes(
#     tickmode='array',
#     tickvals=[10**i for i in range(int(np.log10(df[frequency_column_name].min())), int(np.log10(df[frequency_column_name].max())) + 1)]
# )
# --- Apply Dark Mode Theme ---
fig.update_layout(template="plotly_dark")


# --- Set Y-axis (Amplitude) Maximum to 0 dBm ---
# Plotly auto-determines the min unless specified.
# We set a floor of -100 dBm if the data minimum is higher than -100, just to keep the scale sensible.
y_min = df[columns_to_plot].min().min() # Get overall min across all plotted amplitude series
fig.update_yaxes(range=[y_min if y_min < 0 else -100, 0], # Ensure y-axis goes up to 0 dBm
                 title="Amplitude (dBM)",
                 showgrid=True, gridwidth=1)

# --- Customize Colors and Line Styles (Optional) ---
# Plotly automatically assigns distinct colors. If you want specific colors or line styles,
# you can use fig.update_traces or pass a color_discrete_map to px.line.
# Example: Make the 'Average Amplitude' line thicker or dashed
fig.for_each_trace(lambda trace: trace.update(line=dict(width=2)) if trace.name == 'Average Amplitude' else ())

# This will open the interactive plot in your default web browser if run as a Python script,
# or display it directly in a Jupyter Notebook/Lab environment.
fig.show()

print("\n--- Plotly Express Interactive Plot Generated ---")
print("Run this code in a Python environment (like a Jupyter Notebook or a .py file and open the generated HTML) to experience:")
print("1. Hovering over lines to see exact Frequency and Amplitude values.")
print("2. Clicking on legend items (column names) to hide/show individual lines.")
print("3. Pan, zoom, and other interactive controls.")
