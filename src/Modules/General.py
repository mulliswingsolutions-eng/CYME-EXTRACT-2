# src/general.py
INPUT_SXST = r"C:\Users\ebelliv1\OneDrive - University of New Brunswick\Smart Grid\Github\CYME-EXTRACT\Examples\Example-13bus.txt"

def get_general():
    general_info = {
        "Excel file version": "v2.0",  # Static
        "Name": "Distribution",        # Static
        "Frequency (Hz)": None,
        "Power Base (MVA)": None
    }
    
    with open(INPUT_SXST, "r", encoding="utf-8") as f:
        for line in f:
            if "Frequency" in line:
                try:
                    general_info["Frequency (Hz)"] = float(line.split("=")[1].strip())
                except:
                    pass
            elif "PowerBase" in line:
                try:
                    general_info["Power Base (MVA)"] = float(line.split("=")[1].strip())
                except:
                    pass

    return general_info

if __name__ == "__main__":
    general_data = get_general()
    for k, v in general_data.items():
        print(f"{k}\t{v}")
