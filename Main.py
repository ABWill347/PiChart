import pandas as pd

import matplotlib.pyplot as plt

import os



FILE_PATH = r""

GRAPH_FILE_PATH = r""

COMPLETED_COLUMN = 'Status of Update'

DATE_COLUMN = 'Date Approved / Effective Date'



df = pd.read_excel(FILE_PATH)



print("Columns in the DataFrame:", df.columns)



df_incomplete = df[df[COMPLETED_COLUMN] != 'Complete1']



df_excluded = df[df[COMPLETED_COLUMN] == 'Complete1']

count_excluded = df_excluded.shape[0]

print(f"\nExcluded Entries (Status is 'Complete1'): {count_excluded}")

print(df_excluded)



df_incomplete = df_incomplete.copy()

df_incomplete[DATE_COLUMN] = pd.to_datetime(df_incomplete[DATE_COLUMN], errors='coerce')



df_invalid_dates = df_incomplete[df_incomplete[DATE_COLUMN].isna()]

count_invalid_dates = df_invalid_dates.shape[0]



print(f"\nEntries with invalid or missing dates: {count_invalid_dates}")

print(df_invalid_dates)



df_incomplete_valid_dates = df_incomplete.dropna(subset=[DATE_COLUMN])




def filter_by_date_condition(df, date_column, cutoff_date, condition):
    if condition == 'before':

        return df[df[date_column] < cutoff_date]

    elif condition == 'on_or_before':

        return df[df[date_column] <= cutoff_date]

    elif condition == 'after':

        return df[df[date_column] > cutoff_date]

    elif condition == 'on_or_after':

        return df[df[date_column] >= cutoff_date]

    else:

        raise ValueError("Condition must be 'before', 'on_or_before', 'after', or 'on_or_after'.")


# Date cutoff for October 2021

cutoff_date = pd.Timestamp('2021-10-01')

condition = 'on_or_before'  # 'before', 'on_or_before', 'after', 'on_or_after'

# Filter data

df_before_cutoff = filter_by_date_condition(df_incomplete_valid_dates, DATE_COLUMN, cutoff_date, 'before')

df_on_or_after_cutoff = filter_by_date_condition(df_incomplete_valid_dates, DATE_COLUMN, cutoff_date, 'on_or_after')

# Count entries

count_before_cutoff = df_before_cutoff.shape[0]

count_on_or_after_cutoff = df_on_or_after_cutoff.shape[0]

# Calculate total number of entries included in the chart

count_included = count_before_cutoff + count_on_or_after_cutoff

# Calculate total entries in the original dataset

total_entries = df.shape[0]

# Print out the numbers that make up the graph and totals

print(f"\nTotal entries in the original dataset: {total_entries}")

print(f"Total entries excluded: {count_excluded}")

print(f"Total entries included in the chart: {count_included}")

print(f"Entries categorized as 'Before Oct 2021': {count_before_cutoff}")

print(f"Entries categorized as 'Oct 2021 to Present': {count_on_or_after_cutoff}")

# Data for the pie chart

labels = ['Oct 2021 to Present', 'Before Oct 2021']

sizes = [count_on_or_after_cutoff, count_before_cutoff]

colors = ['lightcoral', 'lightskyblue']

# Plotting the pie chart

fig, ax = plt.subplots(figsize=(14, 7))

# Position the pie chart

ax.set_position([0.1, 0.1, 0.35, 0.8])  # [left, bottom, width, height]

ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=140)

ax.axis('equal')  # Pie is drawn as a circle.

# Adding text

plt.text(

    1.2, 0.5,


    f"Total entries included in the chart: {count_included}\n"

    f"Entries categorized as 'Before Oct 2021': {count_before_cutoff}\n"

    f"Entries categorized as 'Oct 2021 to Present': {count_on_or_after_cutoff}",

    fontsize=12, bbox=dict(facecolor='white', alpha=0.5),

    verticalalignment='center'

)

plt.title('GSPs categorized by timeframe')


try:

    plt.savefig(GRAPH_FILE_PATH)

    print(f"\nGraph saved to {GRAPH_FILE_PATH}")

except Exception as e:

    print(f"Error saving graph: {e}")


plt.show()


def open_image(path):
    if os.name == 'nt':

        try:

            os.startfile(path)

        except Exception as e:

            print(f"Error opening image: {e}")



open_image(GRAPH_FILE_PATH)