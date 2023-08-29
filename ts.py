from openpyxl import Workbook

input_file = "input.txt"
data_Cell_S = []
data_Cell_N = []

with open(input_file, "r") as file:
    current_entry = {}
    check_1=check_2=check_3 = False
    for line in file:
        line = line.strip()
        if "Serving Cell Measurements Response" in line:
            check_1 = True
        elif "Neighbor Cell Measurements" in line:  
            check_2 = True
        elif "Cell Reselection" in line:      #for RANK  
            check_3 = True
        
        if check_1:
            if line.startswith("Time :"):
                current_entry["Time"] = line.split(" : ")[1].strip()
            elif "PCI :" in line:
                current_entry["PCI"] = line.split(":")[1].strip()
            elif "RSRP :" in line:
                current_entry["RSRP"] = line.split(":")[1].strip()
            elif "SrxLev :" in line:
                current_entry["SrxLev"] = line.split(":")[1].strip()
            elif "Rank :" in line:
                current_entry["Rank"] = line.split(":")[1].strip()
                data_Cell_S.append(current_entry)
                current_entry = {}
            elif "///" in line:
                check_1 = False
        if check_2:
            if "Time :" in line:
                current_entry["Time"] = line.split(" : ")[1].strip()
            elif "Number of Cells" in line:
                current_entry["Number of Cells"] = line.split(":")[1].strip()
            elif "PCI :" in line:
                current_entry["PCI"] = line.split(":")[1].strip()
            elif "RSRP :" in line:
                current_entry["RSRP"] = line.split(":")[1].strip()
            elif "RSRP :" in line:
                current_entry["RSRP"] = line.split(":")[1].strip()
            elif "///" in line:
                check_2 = False
                data_Cell_N.append(current_entry)
                current_entry = {}

            
        
# Print collected data_Cell_S (for verification)
for entry in data_Cell_S:
    print(entry)

out_file = "out_S_cell.xlsx"

wb = Workbook()
ws = wb.active

# Write headers
ws.append(["Time", "PCI", "RSRP", "SrxLev", "Rank"])

# Write data_Cell_S to the worksheet
for entry in data_Cell_S:
    ws.append([
        entry.get("Time", ""),
        entry.get("PCI", ""),
        entry.get("RSRP", ""),
        entry.get("SrxLev", ""),
        entry.get("Rank", "")
    ])

wb.save(out_file)
print("data_Cell_S has been written to", out_file)

for entry in data_Cell_N:
    print(entry)

out_file = "out_N_cell.xlsx"
wb = Workbook()
ws = wb.active
ws.append(["Time","Number of Cells","PCI","RSRP","RSRQ"])
for entry in data_Cell_N:
    ws.append([
        entry.get("Time",""),
        entry.get("Number of Cells",""),
        entry.get("PCI",""),
        entry.get("RSRP",""),
        entry.get("RSRQ","")
    ])
wb.save(out_file)
print("Write data to N_cell OK")