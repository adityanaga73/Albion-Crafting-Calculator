import csv, os
from collections import OrderedDict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

base=os.path.dirname(__file__)
with open(os.path.join(base,"data","recipes_flat.csv"),encoding="utf-8-sig",newline="") as f: recipes=list(csv.DictReader(f))
products=OrderedDict(); materials=OrderedDict()
for r in recipes:
 products.setdefault(r["Product_ID"],r["Product_Name"]); materials.setdefault(r["Material_ID"],r["Material_Name"])
wb=Workbook(); dash=wb.active; dash.title="Dashboard"
calc=wb.create_sheet("Crafting Calculator"); db=wb.create_sheet("Recipe Database"); items=wb.create_sheet("Item Database"); prices=wb.create_sheet("Market Prices"); settings=wb.create_sheet("Settings"); guide=wb.create_sheet("Read Me")
def top(ws,t,n):
 ws.cell(1,1,t); ws.cell(1,1).font=Font(size=15,bold=True,color="FFFFFF"); ws.cell(1,1).fill=PatternFill("solid",fgColor="17365D"); ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=n)
def hdr(ws,r,vals):
 for j,v in enumerate(vals,1):
  c=ws.cell(r,j,v); c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="1F4E78"); c.alignment=Alignment(wrap_text=True)
top(db,"RECIPE DATABASE — INGREDIENTS PER CRAFT",14)
hdr(db,3,["Product_ID","Product_Name","Tier","Category","Subcategory","Variant_No","Material_ID","Material_Name","Qty_Per_Craft","Output_Per_Craft","Silver_Fee","Focus_Cost","Time_Seconds","Lookup_Key"])
for i,r in enumerate(recipes,4):
 vals=[r["Product_ID"],r["Product_Name"],r["Tier"],r["Category"],r["Subcategory"],r["Variant_No"],r["Material_ID"],r["Material_Name"],float(r["Qty_Per_Craft"] or 0),float(r["Output_Per_Craft"] or 1),float(r["Silver_Fee"] or 0),float(r["Focus_Cost"] or 0),float(r["Time_Seconds"] or 0)]
 for j,v in enumerate(vals,1): db.cell(i,j,v)
 db.cell(i,14,f'=A{i}&"|"&F{i}&"|"&COUNTIFS($A$4:A{i},A{i},$F$4:F{i},F{i})')
db.freeze_panes="A4"; db.auto_filter.ref=f"A3:N{len(recipes)+3}"
for c,w in {"A":28,"B":30,"C":9,"D":16,"E":18,"F":12,"G":36,"H":38,"I":15,"J":16,"K":14,"L":14,"M":14,"N":3}.items(): db.column_dimensions[c].width=w
db.column_dimensions["N"].hidden=True
top(items,"ITEM DATABASE",3); hdr(items,3,["Item_ID","Item_Name","Recipe_Available"])
for i,(pid,name) in enumerate(products.items(),4): items.cell(i,1,pid); items.cell(i,2,name); items.cell(i,3,"Yes")
items.freeze_panes="A4"; items.auto_filter.ref=f"A3:C{len(products)+3}"; items.column_dimensions["A"].width=42; items.column_dimensions["B"].width=48
top(prices,"MARKET PRICES — ENTER CURRENT PRICES",4); hdr(prices,3,["Material_ID","Material_Name","Buy Price / Unit (Silver)","Last Updated"])
for i,(mid,name) in enumerate(materials.items(),4):
 prices.cell(i,1,mid); prices.cell(i,2,name); prices.cell(i,3,0).fill=PatternFill("solid",fgColor="FFF2CC")
prices.freeze_panes="A4"; prices.auto_filter.ref=f"A3:D{len(materials)+3}"
for c,w in {"A":46,"B":48,"C":28,"D":20,"F":55}.items(): prices.column_dimensions[c].width=w
prices["F3"]="NOTE"; prices["F4"]="Yellow cells are manual price inputs. This dataset does not include live market prices."
top(settings,"CALCULATOR SETTINGS",3); hdr(settings,3,["Setting","Value","Explanation"])
for row in [["Resource Return Rate",0,"Enter decimal percentage; default 0%."],["Market Sell Tax",0.04,"Editable assumption; verify current in-game rate."],["Station Fee per Batch",0,"Enter silver fee per batch."],["Other Fee per Batch",0,"Optional additional silver fee per batch."],["Crafting City","Bridgewatch","Reference only."]]: settings.append(row)
for r in range(4,8): settings.cell(r,2).fill=PatternFill("solid",fgColor="FFF2CC")
settings["B4"].number_format="0.00%"; settings["B5"].number_format="0.00%"; settings.column_dimensions["A"].width=30; settings.column_dimensions["C"].width=70
top(calc,"ALBION ONLINE — CRAFTING PROFIT CALCULATOR",8)
calc["A2"]="Choose an item and recipe variant. Enter market prices and finished-item selling price."; calc.merge_cells("A2:H2")
for cell,label in [("A4","Product Item ID"),("A5","Recipe Variant No."),("A6","Craft Quantity (batches)"),("A7","Finished Goods Sell Price / Unit"),("A8","Resource Return Rate"),("A9","Crafting City")]: calc[cell]=label; calc[cell].font=Font(bold=True)
for cell,val in [("B4",next(iter(products))),("B5",1),("B6",1),("B7",0),("B8","=Settings!B4"),("B9","=Settings!B8")]: calc[cell]=val
for cell in ["B4","B5","B6","B7"]: calc[cell].fill=PatternFill("solid",fgColor="FFF2CC")
calc["B8"].number_format="0.00%"; calc["B7"].number_format="#,##0.00"
wb.defined_names.add(DefinedName("ProductIDs",attr_text="'Item Database'!$A$4:$A$20000"))
dv=DataValidation(type="list",formula1="=ProductIDs"); calc.add_data_validation(dv); dv.add(calc["B4"])
hdr(calc,12,["Material ID","Material Name","Qty / Batch","Total Qty Needed","Unit Buy Price","Material Cost","Cost After Return","Price Status"])
for n in range(30):
 row=13+n; key=f'$B$4&"|"&$B$5&"|"&ROWS($A$13:A{row})'
 formulas=[
 f'=IFERROR(INDEX(\\'Recipe Database\\'!$G$4:$G$20000,MATCH({key},\\'Recipe Database\\'!$N$4:$N$20000,0)),"")',
 f'=IFERROR(INDEX(\\'Recipe Database\\'!$H$4:$H$20000,MATCH({key},\\'Recipe Database\\'!$N$4:$N$20000,0)),"")',
 f'=IFERROR(INDEX(\\'Recipe Database\\'!$I$4:$I$20000,MATCH({key},\\'Recipe Database\\'!$N$4:$N$20000,0)),0)',
 f'=C{row}*$B$6',
 f'=IFERROR(VLOOKUP(A{row},\\'Market Prices\\'!$A$4:$C$20000,3,FALSE),0)',
 f'=D{row}*E{row}',f'=F{row}*(1-$B$8)',f'=IF(A{row}="","",IF(E{row}>0,"Price entered","ENTER PRICE"))']
 for j,formula in enumerate(formulas,1): calc.cell(row,j,formula)
hdr(calc,45,["SUMMARY","Value"])
summ=[("Total Material Cost","=SUM(F13:F42)"),("Material Cost After Return","=SUM(G13:G42)"),("Crafting Station Fees","=Settings!B6*$B$6"),("Other Fees","=Settings!B7*$B$6"),("TOTAL CRAFTING COST","=B47+B48+B49"),("Total Output Units",'=IFERROR(INDEX(\\'Recipe Database\\'!$J$4:$J$20000,MATCH($B$4&"|"&$B$5&"|1",\\'Recipe Database\\'!$N$4:$N$20000,0))*$B$6,0)'),("Gross Sales Value","=B51*$B$7"),("Market Sell Tax","=B52*Settings!B5"),("NET SALES VALUE","=B52-B53"),("ESTIMATED PROFIT / LOSS","=B54-B50"),("Profit per Output Unit","=IFERROR(B55/B51,0)"),("ROI","=IFERROR(B55/B50,0)")]
for i,(label,formula) in enumerate(summ,46): calc.cell(i,1,label); calc.cell(i,2,formula); calc.cell(i,2).number_format="#,##0.00"
calc["B57"].number_format="0.00%"; calc["A59"]="Price Check"; calc["B59"]='=IF(COUNTIF(H13:H42,"ENTER PRICE")>0,"Some material prices are missing","All listed material prices entered")'; calc.merge_cells("B59:H59")
for c,w in {"A":30,"B":42,"C":15,"D":19,"E":19,"F":19,"G":20,"H":20}.items(): calc.column_dimensions[c].width=w
calc.freeze_panes="A13"
top(dash,"ALBION CRAFTING PRO — DATABASE CONNECTED",4)
for r,(a,b) in enumerate([("Products with recipes",len(products)),("Recipe ingredient rows",len(recipes)),("Unique materials to price",len(materials))],3): dash.cell(r,1,a); dash.cell(r,2,b)
dash["A8"]="START HERE"; dash["A8"].font=Font(bold=True,color="FFFFFF"); dash["A8"].fill=PatternFill("solid",fgColor="0F6B78")
for i,s in enumerate(["1. Select Product Item ID in Crafting Calculator B4.","2. Choose the recipe variant number in B5.","3. Enter craft quantity and finished goods sell price.","4. Enter buy prices in yellow cells on Market Prices.","5. Set resource return, tax and fees in Settings; verify current values.","6. Review ingredient quantities, crafting cost, net sales and profit/loss."],9): dash.cell(i,1,s); dash.merge_cells(start_row=i,start_column=1,end_row=i,end_column=4)
top(guide,"READ ME — IMPORTANT NOTES",3)
for i,(a,b) in enumerate([("Dataset","Community-sourced recipe data; verify important recipes in-game."),("Prices","Manual inputs; not live market prices."),("Variant","Selects an alternative recipe for the chosen item."),("Resource return","Default 0%; enter only a verified rate."),("Fees and tax","Editable assumptions; check current in-game values."),("Compatibility","Formula-based; no macros required; designed for Excel 2013.")],3):
 guide.cell(i,1,a).font=Font(bold=True); guide.cell(i,2,b); guide.merge_cells(start_row=i,start_column=2,end_row=i,end_column=3)
guide.column_dimensions["A"].width=25; guide.column_dimensions["B"].width=45
for ws in wb.worksheets: ws.sheet_view.zoomScale=90; ws.sheet_view.showGridLines=False
wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode="auto"
wb.save(os.path.join(base,"Albion_Crafting_Calculator_Excel2013_FINAL.xlsx"))
print("Created workbook:",len(products),"products;",len(recipes),"ingredient rows;",len(materials),"materials")
