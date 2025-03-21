import csv
from collections import defaultdict

def migrate_ps_script_to_python():
    inputFilePath = r"C:\Users\User\Desktop\JEAV\Weekly Payments Reconcile (daily)\Weekly Payables\Raw File.csv"
    outputFilePath = r"C:\Users\User\Desktop\JEAV\Weekly Payments Reconcile (daily)\Weekly Payables\Combined_Invoice_Report.csv"

    carrierGroups = {
        "DHL": "DHL",
        "OTHER": ["FEDEX", "UPS", "FREIGHT"]
    }

    with open(inputFilePath, mode='r', newline='', encoding='utf-8') as infile:
        reader = csv.DictReader(infile)
        rawData = [row for row in reader]

    groupedData = defaultdict(list)
    for row in rawData:
        invoiceNumber = row["Invoice Number"]
        groupedData[invoiceNumber].append(row)

    invoiceReport = []

    for invoiceNumber, invoiceItems in groupedData.items():
        hasDHL = False
        hasOther = False
        totalAll = 0.0
        totalDHL = 0.0
        totalOther = 0.0
        cm_invoice_number = ""  # Default blank value
        assignedCategory = ""  # New variable to track assignment

        for item in invoiceItems:
            carrier = (item["Carrier"].strip().upper() if item.get("Carrier") else "")
            
            try:
                amount = float(item["Customer Total"])
            except (ValueError, TypeError):
                amount = 0.0

            if carrier == carrierGroups["DHL"]:
                hasDHL = True
                totalDHL += amount
            elif carrier in carrierGroups["OTHER"]:
                hasOther = True
                totalOther += amount

            totalAll += amount

            # Check for CM Invoice # condition
            if item.get("Customer #") == "10003217":
                cm_invoice_number = item.get("Airbill Number", "")

                # If no valid carrier is assigned, default to "SHIP"
                if carrier not in carrierGroups["OTHER"] and carrier != carrierGroups["DHL"]:
                    assignedCategory = "SHIP"

        # Assign categories based on conditions
        if assignedCategory == "SHIP":
            percentage = 1.0
            invoiceReport.append({
                "Invoice Number": invoiceNumber,
                "Column B": "SHIP",
                "Customer Total": round(totalAll, 2),
                "Percentage of Total (%)": f"{percentage:.10f}",
                "CM Invoice #": cm_invoice_number
            })
        elif hasDHL and not hasOther:
            percentage = 1.0
            invoiceReport.append({
                "Invoice Number": invoiceNumber,
                "Column B": "ILS",
                "Customer Total": round(totalAll, 2),
                "Percentage of Total (%)": f"{percentage:.10f}",
                "CM Invoice #": cm_invoice_number
            })
        elif not hasDHL and hasOther:
            percentage = 1.0
            invoiceReport.append({
                "Invoice Number": invoiceNumber,
                "Column B": "SHIP",
                "Customer Total": round(totalAll, 2),
                "Percentage of Total (%)": f"{percentage:.10f}",
                "CM Invoice #": cm_invoice_number
            })
        elif hasDHL and hasOther:
            percentageDHL = round((totalDHL / totalAll), 10) if totalAll != 0 else 0
            invoiceReport.append({
                "Invoice Number": invoiceNumber,
                "Column B": "DHL",
                "Customer Total": round(totalDHL, 2),
                "Percentage of Total (%)": f"{percentageDHL:.10f}",
                "CM Invoice #": cm_invoice_number
            })
            percentageOther = round((totalOther / totalAll), 10) if totalAll != 0 else 0
            invoiceReport.append({
                "Invoice Number": invoiceNumber,
                "Column B": "FEDEX",
                "Customer Total": round(totalOther, 2),
                "Percentage of Total (%)": f"{percentageOther:.10f}",
                "CM Invoice #": cm_invoice_number
            })
        else:
            percentage = 1.0
            invoiceReport.append({
                "Invoice Number": invoiceNumber,
                "Column B": "",
                "Customer Total": round(totalAll, 2),
                "Percentage of Total (%)": f"{percentage:.10f}",
                "CM Invoice #": cm_invoice_number
            })

    fieldnames = ["Invoice Number", "Column B", "Customer Total", "Percentage of Total (%)", "CM Invoice #"]
    with open(outputFilePath, mode='w', newline='', encoding='utf-8') as outfile:
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in invoiceReport:
            writer.writerow(row)

    print(f"Combined report generated at: {outputFilePath}")

if __name__ == "__main__":
    migrate_ps_script_to_python()
