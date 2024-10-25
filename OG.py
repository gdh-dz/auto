# %%
import pandas as pd
import numpy as np
import openpyxl
from customtkinter import *


# %%
dataframe =  pd.read_excel("KNX - Export to SSP 12-06-23.xlsx")
dataframe2 = pd.read_excel("SAP - Export to SSP 09-06-23.xlsx")
dataframe3 = pd.read_excel("Part Properties.xlsx", skiprows=[0])
# %%
#dataframe2.rename(columns={'Vendor NO.':'Supplier Id','Vendor name':'Supplier Name','Plant name':'Part Site Description','Material':'Part Name','Part Description':'Material Descrp', 'Part Planner Code Value':'MRP controller', 'Part Planner Code Description':'MRP controller descrp', 'Part Buyer Code Value':'Purchasing grp', 'Part Buyer Code Description':'Purchasing grp descrp', 'PrGroup Telf':'Purchasing grp Tel', 'PrGroup Email':'Purchasing grp EMail', 'Pur.Org.': 'Purchasing org', 'Pur.Org. Descr': 'Purchasing org descrp', 'Purchasing UoM': 'Order Unit', 'Open PO Lines': 'Open PO Items', 'PO Balance Qty': 'Open PO Qty', 'M00 PO': 'JUN 23-Ordered Qty', 'M00 FCT': 'JUN 23-Qty to be Ordered', 'M01 PO': 'JUL 23-Ordered Qty', 'M01 FCT': 'JUL 23-Qty to be Ordered', 'M02 PO':'AUG 23-Ordered Qty', 'M02 FCT':'AUG 23-Qty to be Ordered', 'M03 PO':'SEP 23-Ordered Qty', 'M03 FCT': 'SEP 23-Qty to be Ordered', 'M04 PO':'OCT 23-Ordered Qty', 'M04 FCT':'OCT 23-Qty to be Ordered', 'M05 PO':'NOV 23-Ordered Qty', 'M05 FCT':'NOV 23-Qty to be Ordered', 'M06 PO':'DEC 23-Ordered Qty', 'M06 FCT':'DEC 23-Qty to be Ordered', 'M07 PO': 'JAN 24-Ordered Qty', 'M07 FCT':'JAN 24-Qty to be Ordered', 'M08 PO':'FEB 24-Ordered Qty', 'M08 FCT': 'FEB 24-Qty to be Ordered', 'M09 PO': 'MAR 24-Ordered Qty', 'M09 FCT':'MAR 24-Qty to be Ordered', 'M10 PO': 'APR 24-Ordered Qty', 'M10 FCT':'APR 24-Qty to be Ordered', 'M11 PO':'MAY 24-Ordered Qty', 'M11 FCT':'MAY 24-Qty to be Ordered', 'M12 PO':'JUN 24-Ordered Qty', 'M12 FCT':'JUN 24-Qty to be Ordered', 'M13 PO':'JUL 24-Ordered Qty','M13 FCT':'JUL 24-Qty to be Ordered', 'M14 PO':'AUG 24-Ordered Qty','M14 FCT':'AUG 24-Qty to be Ordered', 'M15 PO':'SEP 24-Ordered Qty', 'M15 FCT':'SEP 24-Qty to be Ordered', 'M16 PO':'OCT 24-Ordered Qty', 'M16 FCT':'OCT 24-Qty to be Ordered', 'M17 PO':'NOV 24-Ordered Qty', 'M17 FCT':'NOV 24-Qty to be Ordered', 'M18 PO':'DEC 24-Ordered Qty', 'M18 FCT':'DEC 24-Qty to be Ordered', 'Vendor Material Code': 'Vendor Material No.', 'Contact Name':'Contact at vendor','Deliv LT': 'Delivery Lead Time'}, inplace=True)
dataframe2.columns = dataframe.columns
# %%
dataframe2["Supplier Id"] = dataframe2["Supplier Id"].apply(lambda x: x[2:])
dataframe2["Supplier Id"] = dataframe2["Supplier Id"].apply(str)
dataframe2["Part Name"] = dataframe2["Part Name"].apply(str)
dataframe["Supplier Id"] = dataframe["Supplier Id"].apply(str)
dataframe["Part Name"] = dataframe["Part Name"].apply(str)
# %%
outer_join = dataframe2.merge(dataframe, how = 'outer', left_on = ["Supplier Id", "Part Name"], right_on = ["Supplier Id", "Part Name"], suffixes = ["_SAP", "_KNX"])
# %%
print(dataframe2.columns)
# %%
sup_id_list = []
for vendor_id in outer_join["Supplier Id"].values:
    if vendor_id in dataframe["Supplier Id"]:
        sup_id_list.append(vendor_id)
    else:
        sup_id_list.append("")

nuevoarchivo2 = 'nuevoarchivo2.xlsx'

outer_join["Supplier Id_KNX"] =  outer_join["Supplier Id"].apply(lambda x: x if x in dataframe["Supplier Id"].values else "")
outer_join["Part Name_KNX"] = ""
outer_join["SAP.Concat"] = ""
outer_join["KX.Concat"] = ""
columns = ['SAP.Concat','Supplier Id',
 'Supplier Name_SAP',
 'Plant_SAP',
 'Part Site Description_SAP',
 'Part Name',
 'Part Description_SAP',
 'Part Planner Code Value_SAP',
 'Part Planner Code Description_SAP',
 'Part Buyer Code Value_SAP',
 'Part Buyer Code Description_SAP',
 'PrGroup Telf_SAP',
 'PrGroup Email_SAP',
 'Pur.Org._SAP',
 'Pur.Org. Descr_SAP',
 'Purchasing UoM_SAP',
 'Open PO Lines_SAP',
 'PO Balance Qty_SAP',
 'M00 PO_SAP',
 'M00 FCT_SAP',
 'M01 PO_SAP',
 'M01 FCT_SAP',
 'M02 PO_SAP',
 'M02 FCT_SAP',
 'M03 PO_SAP',
 'M03 FCT_SAP',
 'M04 PO_SAP',
 'M04 FCT_SAP',
 'M05 PO_SAP',
 'M05 FCT_SAP',
 'M06 PO_SAP',
 'M06 FCT_SAP',
 'M07 PO_SAP',
 'M07 FCT_SAP',
 'M08 PO_SAP',
 'M08 FCT_SAP',
 'M09 PO_SAP',
 'M09 FCT_SAP',
 'M10 PO_SAP',
 'M10 FCT_SAP',
 'M11 PO_SAP',
 'M11 FCT_SAP',
 'M12 PO_SAP',
 'M12 FCT_SAP',
 'M13 PO_SAP',
 'M13 FCT_SAP',
 'M14 PO_SAP',
 'M14 FCT_SAP',
 'M15 PO_SAP',
 'M15 FCT_SAP',
 'M16 PO_SAP',
 'M16 FCT_SAP',
 'M17 PO_SAP',
 'M17 FCT_SAP',
 'M18 PO_SAP',
 'M18 FCT_SAP',
 'Vendor Material Code_SAP',
 'Contact Name_SAP',
 'Deliv LT_SAP',
 'Base UoM_SAP',
 'KX.Concat',
 'Supplier Id_KNX',
 'Supplier Name_KNX',
 'Plant_KNX',
 'Part Site Description_KNX',
 'Part Name_KNX',
 'Part Description_KNX',
 'Part Planner Code Value_KNX',
 'Part Planner Code Description_KNX',
 'Part Buyer Code Value_KNX',
 'Part Buyer Code Description_KNX',
 'PrGroup Telf_KNX',
 'PrGroup Email_KNX',
 'Pur.Org._KNX',
 'Pur.Org. Descr_KNX',
 'Purchasing UoM_KNX',
 'Open PO Lines_KNX',
 'PO Balance Qty_KNX',
 'M00 PO_KNX',
 'M00 FCT_KNX',
 'M01 PO_KNX',
 'M01 FCT_KNX',
 'M02 PO_KNX',
 'M02 FCT_KNX',
 'M03 PO_KNX',
 'M03 FCT_KNX',
 'M04 PO_KNX',
 'M04 FCT_KNX',
 'M05 PO_KNX',
 'M05 FCT_KNX',
 'M06 PO_KNX',
 'M06 FCT_KNX',
 'M07 PO_KNX',
 'M07 FCT_KNX',
 'M08 PO_KNX',
 'M08 FCT_KNX',
 'M09 PO_KNX',
 'M09 FCT_KNX',
 'M10 PO_KNX',
 'M10 FCT_KNX',
 'M11 PO_KNX',
 'M11 FCT_KNX',
 'M12 PO_KNX',
 'M12 FCT_KNX',
 'M13 PO_KNX',
 'M13 FCT_KNX',
 'M14 PO_KNX',
 'M14 FCT_KNX',
 'M15 PO_KNX',
 'M15 FCT_KNX',
 'M16 PO_KNX',
 'M16 FCT_KNX',
 'M17 PO_KNX',
 'M17 FCT_KNX',
 'M18 PO_KNX',
 'M18 FCT_KNX',
 'Vendor Material Code_KNX',
 'Contact Name_KNX',
 'Deliv LT_KNX',
 'Base UoM_KNX']

outer_join = outer_join[columns]

# HEURISTICA PARA QUE LOS REGISTROS DE KN QUE NO EXISTAN NO TENGAN DATOS EN LA COLUMNA PART_NAME_KNX
outer_join["Part Name_KNX"] = np.where(outer_join["Supplier Name_KNX"].isna(), "", outer_join["Part Name"])
outer_join["Supplier Id_KNX"] = np.where(outer_join["Supplier Name_KNX"].isna(), "", outer_join["Supplier Id_KNX"])
print("Nombres de columnas cambiados")
# %%
outer_join["SAP FCT avrg 12m"] = ""
outer_join["KX FCT avrg 12m"] = ""
outer_join["Avg Kx / Avg SAP Match"] = ""
outer_join["Sum FCT SAP"] = ""
outer_join["Sum FCT Kx"] = ""
outer_join["Is missing in Kx ?"] = ""
outer_join["Is PRO2I"] = ""
outer_join["Is BULK ? "] = ""
outer_join["Gap"] = ""
outer_join["Comment gap - List of critical references"] = ""

outer_join.to_excel(nuevoarchivo2, index=False)
print("Archivo guardado como 'nuevoarchivo2.xlsx'")
# %%
wb = openpyxl.load_workbook("nuevoarchivo2.xlsx")
sheet1 = wb["Sheet1"]
sheet1['R7009'] = "=SUBTOTAL(9,R2:R7008)"
sheet1['S7009'] = "=SUBTOTAL(9,S2:S7008)"
sheet1['T7009'] = "=SUBTOTAL(9,T2:T7008)"
sheet1['U7009'] = "=SUBTOTAL(9,U2:U7008)"
sheet1['V7009'] = "=SUBTOTAL(9,V2:V7008)"
sheet1['W7009'] = "=SUBTOTAL(9,W2:W7008)"
sheet1['X7009'] = "=SUBTOTAL(9,X2:X7008)"
sheet1['Y7009'] = "=SUBTOTAL(9,Y2:Y7008)"
sheet1['Z7009'] = "=SUBTOTAL(9,Z2:Z7008)"
sheet1['AA7009'] = "=SUBTOTAL(9,AA2:AA7008)"
sheet1['AB7009'] = "=SUBTOTAL(9,AB2:AB7008)"
sheet1['AC7009'] = "=SUBTOTAL(9,AC2:AC7008)"
sheet1['AD7009'] = "=SUBTOTAL(9,AD2:AD7008)"
sheet1['AE7009'] = "=SUBTOTAL(9,AE2:AE7008)"
sheet1['AF7009'] = "=SUBTOTAL(9,AF2:AF7008)"
sheet1['AG7009'] = "=SUBTOTAL(9,AG2:AG7008)"
sheet1['AH7009'] = "=SUBTOTAL(9,AH2:AH7008)"
sheet1['AI7009'] = "=SUBTOTAL(9,AI2:AI7008)"
sheet1['AJ7009'] = "=SUBTOTAL(9,AJ2:AJ7008)"
sheet1['AK7009'] = "=SUBTOTAL(9,AK2:AK7008)"
sheet1['AL7009'] = "=SUBTOTAL(9,AL2:AL7008)"
sheet1['AM7009'] = "=SUBTOTAL(9,AM2:AM7008)"
sheet1['AN7009'] = "=SUBTOTAL(9,AN2:AN7008)"
sheet1['AO7009'] = "=SUBTOTAL(9,AO2:AO7008)"
sheet1['AP7009'] = "=SUBTOTAL(9,AP2:AP7008)"
sheet1['AQ7009'] = "=SUBTOTAL(9,AQ2:AQ7008)"
sheet1['AR7009'] = "=SUBTOTAL(9,AR2:AR7008)"
sheet1['AS7009'] = "=SUBTOTAL(9,AS2:AS7008)"
sheet1['AT7009'] = "=SUBTOTAL(9,AT2:AT7008)"
sheet1['AU7009'] = "=SUBTOTAL(9,AU2:AU7008)"
sheet1['AV7009'] = "=SUBTOTAL(9,AV2:AV7008)"
sheet1['AW7009'] = "=SUBTOTAL(9,AW2:AW7008)"
sheet1['AX7009'] = "=SUBTOTAL(9,AX2:AX7008)"
sheet1['AY7009'] = "=SUBTOTAL(9,AY2:AY7008)"
sheet1['AZ7009'] = "=SUBTOTAL(9,AZ2:AZ7008)"
sheet1['BA7009'] = "=SUBTOTAL(9,BA2:BA7008)"
sheet1['BB7009'] = "=SUBTOTAL(9,BB2:BB7008)"
sheet1['BC7009'] = "=SUBTOTAL(9,BC2:BC7008)"
sheet1['BW8786'] = "=SUBTOTAL(9,BW2:BW8785)"
sheet1['BX8786'] = "=SUBTOTAL(9,BX2:BX8785)"
sheet1['BY8786'] = "=SUBTOTAL(9,BY2:BY8785)"
sheet1['BZ8786'] = "=SUBTOTAL(9,BZ2:BZ8785)"
sheet1['CA8786'] = "=SUBTOTAL(9,CA2:CA8785)"
sheet1['CB8786'] = "=SUBTOTAL(9,CB2:CB8785)"
sheet1['CC8786'] = "=SUBTOTAL(9,CC2:CC8785)"
sheet1['CD8786'] = "=SUBTOTAL(9,CD2:CD8785)"
sheet1['CE8786'] = "=SUBTOTAL(9,CE2:CE8785)"
sheet1['CF8786'] = "=SUBTOTAL(9,CF2:CF8785)"
sheet1['CG8786'] = "=SUBTOTAL(9,CG2:CG8785)"
sheet1['CH8786'] = "=SUBTOTAL(9,CH2:CH8785)"
sheet1['CI8786'] = "=SUBTOTAL(9,CI2:CI8785)"
sheet1['CJ8786'] = "=SUBTOTAL(9,CJ2:CJ8785)"
sheet1['CK8786'] = "=SUBTOTAL(9,CK2:CK8785)"
sheet1['CL8786'] = "=SUBTOTAL(9,CL2:LC8785)"
sheet1['CM8786'] = "=SUBTOTAL(9,CM2:CM8785)"
sheet1['CN8786'] = "=SUBTOTAL(9,CN2:CN8785)"
sheet1['CO8786'] = "=SUBTOTAL(9,CO2:CO8785)"
sheet1['CP8786'] = "=SUBTOTAL(9,CP2:CP8785)"
sheet1['CQ8786'] = "=SUBTOTAL(9,CQ2:CQ8785)"
sheet1['CR8786'] = "=SUBTOTAL(9,CR2:CR8785)"
sheet1['CS8786'] = "=SUBTOTAL(9,CS2:CS8785)"
sheet1['CT8786'] = "=SUBTOTAL(9,CT2:CT8785)"
sheet1['CU8786'] = "=SUBTOTAL(9,CU2:CU8785)"
sheet1['CV8786'] = "=SUBTOTAL(9,CV2:CV8785)"
sheet1['CW8786'] = "=SUBTOTAL(9,CW2:CW8785)"
sheet1['CX8786'] = "=SUBTOTAL(9,CX2:CX8785)"
sheet1['CY8786'] = "=SUBTOTAL(9,CY2:CY8785)"
sheet1['CZ8786'] = "=SUBTOTAL(9,CZ2:CZ8785)"
sheet1['DA8786'] = "=SUBTOTAL(9,DA2:DA8785)"
sheet1['DB8786'] = "=SUBTOTAL(9,DB2:DB8785)"
sheet1['DC8786'] = "=SUBTOTAL(9,DC2:DA8785)"
sheet1['DD8786'] = "=SUBTOTAL(9,DD2:DD8785)"
sheet1['DE8786'] = "=SUBTOTAL(9,DE2:DE8785)"
sheet1['DF8786'] = "=SUBTOTAL(9,DF2:DF8785)"
sheet1['DG8786'] = "=SUBTOTAL(9,DG2:DG8785)"
sheet1['DH8786'] = "=SUBTOTAL(9,DH2:DH8785)"
sheet1['DI8786'] = "=SUBTOTAL(9,DI2:DI8785)"

print("Formulas agregadas 'nuevoarchivo2.xlsx'")
 # %%
for i in range(2,8787):
    sheet1[f'A{i}'] = f'=+B{i}&"-"&F{i}'
    sheet1[f'BI{i}'] = f'=+BJ{i}&"-"&BN{i}'
    sheet1[f'DQ{i}'] = f'=IF(ISERROR(AVERAGE(S{i},U{i},W{i},Y{i},AA{i},AC{i},AE{i},AG{i},AI{i},AK{i},AM{i},AO{i})),"0", AVERAGE(S{i},U{i},W{i},Y{i},AA{i},AC{i},AE{i},AG{i},AI{i},AK{i},AM{i},AO{i}))'
    sheet1[f'DR{i}'] = f'=IF(ISERROR(AVERAGE(BY{i},CA{i},CC{i},CE{i},CG{i},CI{i},CK{i},CM{i},CO{i},CQ{i},CS{i},CU{i})),"0", AVERAGE(BY{i},CA{i},CC{i},CE{i},CG{i},CI{i},CK{i},CM{i},CO{i},CQ{i},CS{i},CU{i}))'
    sheet1[f'DS{i}'] = f'=IF(ISERROR(DQ{i}/DR{i}),"Kx or SAP FCT = 0", DQ{i}/DR{i})'
    sheet1[f'DT{i}'] = f'=S{i}+U{i}+W{i}+Y{i}+AA{i}+AC{i}+AE{i}+AG{i}+AI{i}+AK{i}+AM{i}+AO{i}'
    sheet1[f'DU{i}'] = f'=BY{i}+CA{i}+CC{i}+CE{i}+CG{i}+CI{i}+CK{i}+CM{i}+CO{i}+CQ{i}+CS{i}+CU{i}'
    sheet1[f'DV{i}'] = f'=IF(OR(A{i}=BI{i},ISBLANK(BO{i}) = FALSE),"N","Y")'
    sheet1[f'DW{i}'] = f'=IF(ISERROR(VLOOKUP(A{i},#REF!,1,0)),"N","Y")'

wb.save('nuevoarchivo2.xlsx')
print("Tu validación está lista, se guardará en el archivo 'nuevoarchivo2.xlsx'")
# %%