import pandas as pd
import sys
import json

def excel_to_parquet(excel_path, sheet_name, parquet_path):
    try:
        # Read Excel file
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        
        # Convert to Parquet
        df.to_parquet(parquet_path, index=False)
        
        # Return success
        print(json.dumps({"success": True, "message": "Conversion successful"}))
        
    except Exception as e:
        # Return error
        print(json.dumps({"success": False, "message": str(e)}))

def read_parquet(parquet_path):
    try:
        # Read Parquet file
        df = pd.read_parquet(parquet_path)
        
        # Convert to JSON
        result = df.to_json(orient='records')
        print(result)
        
    except Exception as e:
        print(json.dumps({"success": False, "message": str(e)}))

if __name__ == "__main__":
    command = sys.argv[1]
    
    if command == "convert":
        excel_path = sys.argv[2]
        sheet_name = sys.argv[3]
        parquet_path = sys.argv[4]
        excel_to_parquet(excel_path, sheet_name, parquet_path)
    
    elif command == "read":
        parquet_path = sys.argv[2]
        read_parquet(parquet_path)
