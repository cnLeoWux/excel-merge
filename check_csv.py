import pandas as pd

# Check the CSV file directly
try:
    print("Trying to read CSV with default settings...")
    df = pd.read_csv("ExcelForHandel/order.csv")
    print("Success!")
    print(df)
except Exception as e:
    print(f"Error: {e}")
    
    # Try with different parameters
    try:
        print("\nTrying to read CSV with sep='auto'...")
        df = pd.read_csv("ExcelForHandel/order.csv", sep=None, engine='python')
        print("Success with auto separator!")
        print(df)
    except Exception as e2:
        print(f"Error with auto separator: {e2}")
        
        try:
            print("\nTrying to read CSV with different separator...")
            # Common separators
            for sep in [',', ';', '\t', '|']:
                try:
                    df = pd.read_csv("ExcelForHandel/order.csv", sep=sep)
                    print(f"Success with separator '{sep}'!")
                    print(df)
                    break
                except:
                    continue
        except:
            print("All separator attempts failed")