import sys, win32com.client, datetime, time, pyperclip
from openpyxl import load_workbook
#SAPconnection
SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
session = SapGui.FindById("ses[0]")


#tagesdatum
current_time = datetime.datetime.now()

#auslesen excel
excel = load_workbook(r"C:\Users\zerzavy\Desktop\BA\PCBAbestellung.xlsx", data_only=True)
Tabelle1 = excel.active
mat1 = Tabelle1['A2'].value
mat2 = Tabelle1['A3'].value
qty1 = Tabelle1['B2'].value
qty2 = Tabelle1['B3'].value

#datum
today = (str(current_time.day)+"."+str(current_time.month)+"."+str(current_time.year))



#sapwindow
def user():
    
    User = session.findById("wnd[0]/usr")
    zaehler = User.Children.Count - 1
    return User.Children(int(zaehler)).Name

#MY-PO
session.StartTransaction(Transaction="me21n")

#Lieferant
userstring = user()
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text = "10810981"

#Bestelldaten
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]").text = mat1
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]").text = mat2
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]").text = qty1
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]").text = qty2
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EEIND[9,0]").text = today
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EEIND[9,1]").text = today
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[16,0]").text = "MY01"
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[16,1]").text = "MY01"
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB1:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4000/btnDYN_4000-BUTTON").press()

#ORG Daten
userstring = user()
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9").select
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").text = "MY01"
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").text = "si3"
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").text = "MY01"
#Siemenserweiterung
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT11").select()
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT11/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1227/ssubCUSTOMER_DATA_HEADER:SAPLXM06:0101/ctxtEKKO-YYAWV1").text = "N"
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT11/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1227/ssubCUSTOMER_DATA_HEADER:SAPLXM06:0101/ctxtEKKO-YYAWV2").text = "N"
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT11/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1227/ssubCUSTOMER_DATA_HEADER:SAPLXM06:0101/ctxtEKKO-YYBAULS").text = "N"
#speichern
session.findById("wnd[0]").sendVKey (11)
#erstellte PO-nummer
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/tbar[1]/btn[17]").press()
pono = session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text
print(pono)

#auftragsbearbeitung
session.StartTransaction(Transaction="va02")
#auftragssuche
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = ""
session.findById("wnd[0]/usr/txtRV45S-BSTNK").text = "4500791860"
session.findById("wnd[0]/usr/btnBT_SUCH").press()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").key = " "
#strecke
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PSTYV[11,0]").text = "TAS"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PSTYV[11,1]").text = "TAS"

#kopfdaten
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()
session.findById("wnd[0]/tbar[1]/btn[18]").press()
session.findById("wnd[0]/tbar[1]/btn[18]").press()
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07").select()
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,6]").text = "VKAF2F"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08").select()
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem ("ZV02","Column1")
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem ("ZV02","Column1")
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem ("ZV02","Column1")
#kopdaten.mytext
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = "***************************************\nAttention:\nMaterial listed below belongs to applied SMT Machine Parts\n– HS Code 8479.90.000\n***************************************"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10").select()
#objektstatus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4305/btnBT_KSTC").press()
session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,2]").selected = True
session.findById("wnd[0]/tbar[0]/btn[3]").press()
session.findById("wnd[0]/tbar[0]/btn[3]").press()
session.findById("wnd[0]/tbar[0]/btn[11]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()

#PO aus BANF
session.StartTransaction(Transaction="me21n")
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").pressContextButton("SELECT")
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").selectContextMenuItemByPosition (9)
session.findById("wnd[0]/usr/ctxtSP$00034-LOW").text = ""
session.findById("wnd[0]/usr/ctxtSP$00036-LOW").text = ""
session.findById("wnd[0]/tbar[1]/btn[8]").press()
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").selectItem ("          1","&Hierarchy")
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").pressButton ("COPY")

userstring = user()
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]").text = qty1
session.findById("wnd[0]/usr/sub"+userstring+"/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]").text = qty2