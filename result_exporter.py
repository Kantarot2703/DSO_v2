import pandas as pd
import os

def export_result_to_excel(results, output_path="output/result.xlsx"):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # ถ้าเป็น list of dict → แปลงเป็น DataFrame
    if isinstance(results, list):
        df = pd.DataFrame(results)
    elif isinstance(results, pd.DataFrame):
        df = results
    else:
        raise ValueError("❌ Invalid result format. Must be list or DataFrame.")

    df.to_excel(output_path, index=False)
    print(f"✅ Exported result to: {output_path}")
