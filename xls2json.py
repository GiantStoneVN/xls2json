import pandas as pd
import json
import sys

# Check if file path is provided
if len(sys.argv) < 2:
    print("Usage: python script.py <input_excel_file>")
    sys.exit(1)

# Get the input Excel file path from command-line arguments
input_file = sys.argv[1]

# Load Excel file
try:
    df = pd.read_excel(input_file, sheet_name="W1")
except Exception as e:
    print(f"Error reading file: {e}")
    sys.exit(1)

# Convert to JSON with proper encoding
json_data = df.to_json(orient="records", force_ascii=False, indent=4)

# Generate output JSON file name
output_file = input_file.rsplit(".", 1)[0] + ".json"

# Save to a JSON file
with open(output_file, "w", encoding="utf-8") as f:
    f.write(json_data)

print(f"Excel converted to JSON successfully! Output saved as: {output_file}")
