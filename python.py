from openpyxl import Workbook

# Path to the text file
file_path = "input.txt"

# Initialize variables to store data
data_list = []
current_data = {}

# Open the file for reading
with open(file_path, "r") as file:
    inside_measurement = False

    for line in file:
        line.strip()
        if "Serving Cell Measurements Response" in line:
            inside_measurement = True
            current_data = {}
        elif inside_measurement:
            if "Time :" in line:
                current_data["Time"] = line.split(":")
            elif "PCI :" in line:
                current_data["PCI"] = line.split(":")
            elif "RSRP :" in line:
                current_data["RSRP"] = line.strip().split(" : ")
            elif "SrxLev :" in line:
                current_data["SrxLev"] = line.strip().split(" : ")
            elif "Rank :" in line:
                current_data["Rank"] = line.strip().split(" : ")
        elif line == "///":
            inside_measurement = False
            if current_data:
                data_list.append(current_data)
                current_data = {}

# Create an Excel workbook and worksheet
output_file = "output.xlsx"
wb = Workbook()
ws = wb.active

# Write data to the worksheet
ws.append(["Time", "PCI", "RSRP", "SrxLev", "Rank"])
for data in data_list:
    ws.append([data["Time"], data.get("PCI", ""), data.get("RSRP", ""), data.get("SrxLev", ""), data.get("Rank", "")])

# Save the Excel workbook
wb.save(output_file)
