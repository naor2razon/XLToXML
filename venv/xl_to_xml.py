from openpyxl import load_workbook
from yattag import Doc, indent
import math
import numpy as geek

wb = load_workbook("something.xlsx")
ws = wb.worksheets[0]

# Create Yattag doc, tag and text objects
doc, tag, text = Doc().tagtext()


xml_schema = '<Properties><Simple name="Capacity" value="256" /></Properties>'

with tag('Collection',name = "Root",type ="somthing"):
    doc.asis(xml_schema)
    with tag('Items'):
        with tag('Complex'):
            with tag('Properties'):
                for row in ws.iter_rows(min_row=3, max_row=4, min_col=2, max_col=10):
                    row = [cell.value for cell in row]
                    with tag('Complex', name="Begin"):
                        with tag('Properties'):
                            # changing the value each time
                            doc.stag('Simple', name="Latitude", value=row[3])
                            doc.stag('Simple', name="Longitude", value=row[4])
                    with tag('Complex', name="End"):
                        with tag('Properties'):
                            # changing the value each time
                            doc.stag('Simple', name="Latitude", value=row[5])
                            doc.stag('Simple', name="Longitude", value=row[6])
                    #changing the values
                    doc.stag('Simple', name="Lable", value=row[0])
                    doc.stag('Simple', name="PoligonStartDistance", value="Validate")
                    doc.stag('Simple', name="PoligonEndDistance", value=(math.sqrt((row[6] - row[4])**2 + (row[5] - row[3])**2)))
                    doc.stag('Simple', name="CableLength", value=(row[8]-row[7]))
                    doc.stag('Simple', name="OverplusAtBeginning", value="0")


result = indent(
    doc.getvalue(),
    indentation='    ',
    indent_text=True
)

print(result)


#doc.stag('input', type = "submit", value = "Validate")