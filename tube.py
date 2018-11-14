# want to extract tubing sufrace using shape design workbench and open catia files autonomously
import win32com.client

CATIA = win32com.client.Dispatch("CATIA.Application")

print("aaahhhh")
CATIA.Documents.Open("./Tube.CATPart")
print("hello")

# doc = CATIA.ActiveDocument.Product
# doc.Products.Count
# doc.Products.Item
# doc.Products.PartNumber
# CATIA.ActiveDocument.Name => "SealerStop.CATPart"
# CATIA.ActiveDocument.Part.Name => "Tube Practice"
# CATIA.StartWorkbench("GPSCfg") => Analysis Workbench
# print(CATIA.GetWorkBenchId()) => name of current active workbench

# To access the Active Document, you must have a document
# open in Catia on your machine
ad = CATIA.ActiveDocument






#CATIA.ActiveDocument.Analysis.Import(CATIA.ActiveDocument)
#CATIA.ActiveDocument.Import(CATIA.ActiveDocument)

'''

ad = CATIA.ActiveDocument

sel = ad.Selection
sel.Search("CATDrwSearch.DrwDimension,all")

for i in range(1, sel.Count2+1):
    aDim = sel.Item2(i).Value
    aDimValue = aDim.GetValue
    aDimValue.SetFormatPrecision(1,0.001)

sel.Clear
'''
