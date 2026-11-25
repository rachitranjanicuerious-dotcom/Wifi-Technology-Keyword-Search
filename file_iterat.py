
# import os
# import pandas as pd

# # Specify the directory path
# directory_path = r"D:\vs code\Wifi-Gen-Invalidation\Sample Wi-Fi 7 WG Doc\Sample Wi-Fi 7 WG Doc"

# # Extract all file names
# file_names = []

# for file in os.listdir(directory_path):

#     file_path = os.path.join(directory_path, file)

#     if os.path.isfile(file_path):
#         file_names.append(file)

# # Create DataFrame
# df = pd.DataFrame(file_names, columns=["File Name"])

# # Save Excel file in the project directory
# output_path = r"D:\vs code\Wifi-Gen-Invalidation\File_Names.xlsx"

# df.to_excel(output_path, index=False)

# print(f"Total files found: {len(file_names)}")
# print(f"Excel file saved at: {output_path}")
