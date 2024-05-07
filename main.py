import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel('Data-Analysis-output.xlsx', sheet_name='Table_1')
print(df)

# Extract the locations and their IDs
locations = df.columns[1:]
location_ids = df.iloc[0, 1:]

# Initialize lists to store transformed data
rec_locs = []
rec_loc_ids = []
del_locs = []
del_loc_ids = []
distances = []

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    rec_loc = row[0]  # Extract the "Rec Loc" value
    rec_loc_id = row[1]  # Extract the "Rec LocID" value
    del_loc_start = None  # Track the starting index for "Del Loc"
    for i, val in enumerate(row[2:]):
        if pd.notna(val):  # Skip empty cells
            if del_loc_start is None:
                del_loc_start = i + 2  # Shift index to account for starting at index 2
            else:
                # Append values to respective lists
                rec_locs.append(rec_loc)
                rec_loc_ids.append(rec_loc_id)
                # Swap "Del Loc" and "Del LocID" values
                del_loc_ids.append(locations[i + 1])  # Shift index to match column
                del_locs.append(location_ids[i + 1])  # Shift index to match column
                distances.append(val)


# Create a DataFrame from the transformed data
transformed_df = pd.DataFrame({
    'Rec Loc': rec_locs,
    'Rec LocID': rec_loc_ids,
    'Del Loc': del_locs,
    'Del LocID': del_loc_ids,
    'Distance': distances
})

transformed_df.dropna(subset=['Rec Loc', 'Rec LocID'], inplace=True)


print(transformed_df)

# Save the transformed DataFrame to a new Excel file
transformed_df.to_excel('transformed_output.xlsx', index=False)