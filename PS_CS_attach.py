from openpyxl import Workbook

input_file = "input_PS_CS.txt"
data = []

with open(input_file, "r") as file:
    current_entry = {}
    check_1=check_2=check_3=check_4=check_5=check_6  = False
    for line in file:
        line = line.strip()
        if "***" in line:
            if check_1 | check_2|check_3|check_4|check_5|check_6:
                data.append(current_entry)
                current_entry = {}
            check_1=check_2=check_3=check_4=check_5=check_6  = False
            
        if "Attach Request" in line:
            check_1 = True
        elif "Layer 3 Message type: Tracking Area Update Request" in line:  
            check_2 = True
        elif "Paging-NB" in line:
            check_3 = True
        elif "Layer 3 Message type: Control Plane Service Request" in line:  
            check_4 = True
        elif "RRC Connection Release" in line:  
            check_5= True
        elif "RRC Connection Setup-NB" in line:  
            check_6 = True
        
        # elif "Idle Mode"  in line:  
        #     current_entry["MODE"] = "Idle Mode"
        # elif "Dedicated Mode"  in line:  
        #     current_entry["MODE"] = "Dedicated Mode"
        if check_1:
            current_entry["Type_Name"] = "Attach Request"
            if line.startswith("Time :"):
                current_entry["Time"] = line.split(" : ")[1].strip() 
            elif "EPS Attach Type Value :" in line:
                current_entry["EPS Attach Type Value"] = line.split(":")[1].strip()   
                # data.append(current_entry)
                # current_entry = {}

        if check_2:
            current_entry["Type_Name"] = "Tracking Area Update Request"
            if "Time :" in line:
                current_entry["Time"] = line.split(" : ")[1].strip()
            elif "EPS update type Value :" in line:
                current_entry["EPS update type Value"] = line.split(":")[1].strip()
                # data.append(current_entry)
                # current_entry = {}

        if check_3:
            current_entry["Type_Name"] = "Paging-NB"
            if "Time :" in line:
                current_entry["Time"] = line.split(" : ")[1].strip()
                # data.append(current_entry)
                # current_entry = {}
        if check_4:
            current_entry["Type_Name"] = "Control Plane Service Request"
            if "Time :" in line:
                current_entry["Time"] = line.split(" : ")[1].strip()
                # data.append(current_entry)
                # current_entry = {}
        if check_5:
            if "Time :" in line:
                current_entry["Time"] = line.split(" : ")[1].strip()   
                current_entry["MODE"] = "Idle Mode"
                # data.append(current_entry)
                # current_entry = {}
        if check_6:
            if "Time :" in line:
                current_entry["Time"] = line.split(" : ")[1].strip()   
                current_entry["MODE"] = "Dedicated Mode"
                # data.append(current_entry)
                # current_entry = {}
            

for entry in data:
    print(entry)

out_file = "out_PS_CS.xlsx"

wb = Workbook()
ws = wb.active

# Write headers
ws.append(["Time", "Mode", "Type Name", "EPS Attach Type Value", "EPS update type Value"])

# Write data_Cell_S to the worksheet
for entry in data:
    ws.append([
        entry.get("Time", ""),
        entry.get("MODE", ""),
        entry.get("Type_Name", ""),
        entry.get("EPS Attach Type Value", ""),
        entry.get("EPS update type Value", "")
    ])

wb.save(out_file)
print("data_Cell_S has been written to", out_file)

