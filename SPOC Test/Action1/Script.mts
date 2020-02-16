

SAPGuiSession("Session").SAPGuiWindow("SAP").Maximize @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
SAPGuiSession("Session").SAPGuiWindow("SAP").SAPGuiEdit("User").Set "DEV_ECQ2" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_
SAPGuiSession("Session").SAPGuiWindow("SAP").SAPGuiEdit("Password").Set "Init@1234" @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf1.xml_;_
SAPGuiSession("Session").SAPGuiWindow("SAP").SAPGuiEdit("Password").SetFocus @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf1.xml_;_
SAPGuiSession("Session").SAPGuiWindow("SAP").SendKey ENTER @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_

' Code to handle Pop up message

If SAPGuiSession("Session").SAPGuiWindow("System Messages").SAPGuiButton("Continue   (F12)").Exist Then
	
	
	SAPGuiSession("Session").SAPGuiWindow("System Messages").SAPGuiButton("Continue   (F12)").Click
End If


If SAPGuiSession("Session").SAPGuiWindow("License Information for").SAPGuiRadioButton("Continue with this logon,").Exist Then
	
	
SAPGuiSession("Session").SAPGuiWindow("License Information for").SAPGuiRadioButton("Continue with this logon,").Set @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf2.xml_;_
SAPGuiSession("Session").SAPGuiWindow("License Information for").SAPGuiRadioButton("Continue with this logon,").SetFocus @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf2.xml_;_
SAPGuiSession("Session").SAPGuiWindow("License Information for").SAPGuiButton("Confirm Selection   (Enter)").Click	

End If

 
SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access").SAPGuiOKCode("OKCode").Set "ZDEVOPS_BOM_USAGE" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf3.xml_;_
SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access").SendKey ENTER @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf3.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Program to call CDS").SAPGuiButton("Multiple selection").CaptureBitmap "FirstScreen.png"

SAPGuiSession("Session").SAPGuiWindow("Program to call CDS").SAPGuiButton("Multiple selection").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf4.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Multiple Selection for").SAPGuiTable("SAPLALDBSINGLE").SetCellData 1, "Single value", DataTable("Material", dtGlobalSheet) @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf5.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Multiple Selection for").SAPGuiTable("SAPLALDBSINGLE").SetCellData 2, "Single value", DataTable("Material1", dtGlobalSheet) @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf5.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Multiple Selection for").SAPGuiTable("SAPLALDBSINGLE").SetCellData 3, "Single value", DataTable("Material2", dtGlobalSheet) @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf5.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Multiple Selection for").SAPGuiTable("SAPLALDBSINGLE").SelectCell 3,"Single value" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf5.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Multiple Selection for").CaptureBitmap "Material.png"
'Code to check the validation part

SAPGuiSession("Session").SAPGuiWindow("Multiple Selection for").SAPGuiButton("Copy   (F8)").Click @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf5.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Program to call CDS").SAPGuiButton("Execute   (F8)").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf6.xml_;_
'Code to check the validation part
SAPGuiSession("Session").SAPGuiWindow("Program to call CDS_2").CaptureBitmap "Materialdescription.png"

Rownumber=SAPGuiSession("Session").SAPGuiWindow("Program to call CDS_2").SAPGuiGrid("GridViewCtrl").FindRowByCellContent("Material description", "IND-CENTRIFUGAL")

'msgbox Rownumber

' Get Content from table

Ccontent=SAPGuiSession("Session").SAPGuiWindow("Program to call CDS_2").SAPGuiGrid("GridViewCtrl").GetCellData(Rownumber, "Component quantity")

'msgbox Ccontent

' Compare the Quantity with Actual if both are equal make it pass

If Ccontent="8,000" Then
	Reporter.ReportEvent micPass, "BOM Quantity:", "BOM quantity matched with expected quantity"
	else
	Reporter.ReportEvent micFail, "BOM Quantity:", "BOM quantity does not match with expected quantity"
	
End If

'SAPGuiSession("Session").SAPGuiWindow("Program to call CDS_2").SAPGuiEdit("Class Name:=SAPGuiOKCode", "type:=GuiOkCodeField").Set "/nex"

SAPGuiSession("Session").SAPGuiWindow("Program to call CDS_2").SAPGuiOKCode("OKCode").Set "/nex" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf13.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Program to call CDS_2").SendKey ENTER @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf13.xml_;_


Systemutil.CloseProcessByWndTitle ("SAP Logon 740")


