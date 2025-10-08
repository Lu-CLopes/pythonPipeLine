import pandas as pd
import os

def setup_directories():
    """Setup the necessary directories and paths"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(script_dir, "inputExcel.xlsx")
    output_dir = os.path.join(script_dir, "saida")
    os.makedirs(output_dir, exist_ok=True)
    return input_path, output_dir

def create_first_excel(df, output_dir):
    """Create excel with columns A-F"""
    df_first = df.iloc[:, :6]  # Columns A to F
    output_path = os.path.join(output_dir, "parte_A_F.xlsx")
    df_first.to_excel(output_path, index=False)
    return output_path

def create_second_excel(df, output_dir):
    """Create excel with columns G-I and add Price Scenario column"""
    df_second = df.iloc[:, 6:9].copy()  # Columns G to I
    
    # Get the column names for Sales Org and Dist Channel to concat 
    sales_org_col = None
    dist_channel_col = None
    
    # MUST to found columns 
    for col in df.columns:
        if 'sales' in col.lower() and 'org' in col.lower():
            sales_org_col = col
        if 'dist' in col.lower() and 'channel' in col.lower():
            dist_channel_col = col
    
    # Add Price Scenario column
    df_second['Price Scenario'] = df.apply(
        lambda row: f"LUIZA_TUTTI_{row[sales_org_col]}_{row[dist_channel_col]}", 
        axis=1
    )
    
    output_path = os.path.join(output_dir, "parte_G_I.xlsx")
    df_second.to_excel(output_path, index=False)
    return output_path

def main():
    try:
        # Setup paths
        input_path, output_dir = setup_directories()
        
        # Read input Excel
        df = pd.read_excel(input_path)
        
        # Generate excel files
        first_excel = create_first_excel(df, output_dir)
        
        second_excel = create_second_excel(df, output_dir)

    except FileNotFoundError:
        print(f"❌ Error: Input file not found at {input_path}")
    except ValueError as e:
        print(f"❌ Error: {str(e)}")


if __name__ == "__main__":
    main()