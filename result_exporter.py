import pandas as pd
import os

def export_result_to_excel(results, output_path="output/result.xlsx"):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df = pd.DataFrame(results)
    df.to_excel(output_path, index=False)
    print(f"✅ Exported result to: {output_path}")
