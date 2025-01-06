' Basics initializations

Public sapSession As Object
Public sapSupplier As String
Public commonTextUsed As String

Private Sub handleInitializeBasics()
     Call initializeSAPConnection
     Call initializeGeralVariables
End Sub

Private Sub initializeSAPConnection()
     Dim SapGuiAuto As Object: Set SapGuiAuto = GetObject("SAPGUI")
     Dim Applic As Object: Set Applic = SapGuiAuto.GetScriptingEngine
     Dim Connection As Object: Set Connection = Applic.Children(0)
     Set sapSession = Connection.Children(0)
End Sub

Private Sub initializeGeralVariables()
     sapSupplier = Range("Forn_sap").Value
     commonTextUsed = Range("Texto_padrao_AP").Value
End Sub


' SAP Generics (Export, opentable, exit, etc)

Private Sub navigateToSAPTable(table As String)
     sapSession.findById("wnd[0]/tbar[0]/okcd").text = "/n" & table
     sapSession.findById("wnd[0]").sendVKey 0
End Sub

Private Sub exitPages(Optional numberExits As Integer = 1)
     Dim exitIndex As Integer
     For exitIndex = 1 To numberExits
          sapSession.findById("wnd[0]").sendVKey 3
          Next exitIndex
End Sub

Private Sub ResizeAndCopyData(matrixToBeCopied As Variant)
     Dim tempWorksheet As Worksheet: Set tempWorksheet = ThisWorkbook.Sheets("Temp")
     Dim initialRange As Range: Set initialRange = tempWorksheet.Range("A1")
     tempWorksheet.Cells.Clear
     initialRange.Resize(UBound(matrixToBeCopied, 1), UBound(matrixToBeCopied, 2)).Value = matrixToBeCopied
     tempWorksheet.Range(initialRange, initialRange.Offset(UBound(matrixToBeCopied))).Copy
End Sub

Private Sub exportSAPtoClipboard(Optional typeExport As Integer = 1)
     Select Case typeExport
     
     Case 1
          sapSession.findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
          sapSession.findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").selectContextMenuItem "&PC"
     Case 2
          sapSession.findById("wnd[0]/tbar[1]/btn[45]").press
     End Select
          sapSession.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
          sapSession.findById("wnd[1]/tbar[0]/btn[0]").press
     Application.Wait (Now + TimeValue("0:00:02"))
End Sub


' Processes Callings

Sub sap_cancelInfoRecord()
     Call processInfoRecord(True)
End Sub

Sub sap_uncancelInfoRecord()
     Call processInfoRecord(False)
End Sub

Sub sap_changeIVAInforRecord()
     Call mainUpdateValuesME12(Range("R10"), 1)
End Sub

Sub sap_changePriceValue()
     If MsgBox("Confirmar que gostaria de alterar o valor do valores abaixo manualmente", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
     Call mainUpdateValuesME12(Range("U10"), 0)
End Sub

Sub sap_changePricebyLoad()
     Call updateLoad
End Sub

Sub sap_updatePriceReportSAP()
     Call importPriceReportSAP
End Sub

Sub sap_NewPriceCodeSAP()
     Call createNewPriceCodeSAP
End Sub

Sub sap_useCodeDirectly()

     Call useQuotationCodeinZI9(Range("P7").Value, Range("M7").Value)
End Sub



' Main subs -> Using ME12 and ME15
Private Sub processInfoRecord(cancel As Boolean)
     Dim rowIndex As Integer
     Dim materialList() As Variant
     Dim cancelType As String: cancelType = IIf(cancel, "Cancelamento", "Descancelamento")
     
     Call handleInitializeBasics
     Call initializeMaterialList(materialList, 4, Range("B10"))
     Call navigateToSAPTable("ME15")

     For rowIndex = 1 To UBound(materialList)
          Call cancelItemME15(materialList, rowIndex, cancel)
     Next rowIndex

     Call navigateToSAPTable("ME01")
     For rowIndex = 1 To UBound(materialList)
          Call cancelME01(materialList, rowIndex)
     Next rowIndex

End Sub

Private Sub cancelItemME15(materialList As Variant, rowIndex As Integer, cancel As Boolean)
     If materialList(rowIndex, 4) <> "" Then Exit Sub
     If materialList(rowIndex, 3) = "AMBOS" Then
          Call populateHeaderME(CDbl(materialList(rowIndex, 1)), CStr(materialList(rowIndex, 2)), "0212")
          Call changeME15(cancel)
          Call populateHeaderME(CDbl(materialList(rowIndex, 1)), CStr(materialList(rowIndex, 2)), "0304")
          Call changeME15(cancel)
     Else
          Call populateHeaderME(CDbl(materialList(rowIndex, 1)), CStr(materialList(rowIndex, 2)), CStr(materialList(rowIndex, 3)))
          Call changeME15(cancel)
     End If
End Sub

Private Sub initializeMaterialList(ByRef materialList As Variant, numberColumns As Integer, initialItem As Range)
     Dim listSize As Integer
     Dim rowIndex As Integer, colIndex As Integer
     If IsEmpty(initialItem.Offset(1, 0)) Then
          listSize = 1
     Else
          listSize = initialItem.End(xlDown).Row - initialItem.Row + 1
     End If
     ReDim materialList(1 To listSize, 1 To numberColumns)
     For colIndex = 1 To numberColumns
          For rowIndex = 1 To UBound(materialList)
               materialList(rowIndex, colIndex) = initialItem.Offset(rowIndex - 1, colIndex - 1).Value
          Next rowIndex
     Next colIndex
End Sub

Private Sub populateHeaderME(material As Double, Optional supplier As String = "", Optional plantCode As String = "0212")
     If supplier = "" Then supplier = sapSupplier
     sapSession.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = supplier
     sapSession.findById("wnd[0]/usr/ctxtEINA-MATNR").text = material
     sapSession.findById("wnd[0]/usr/ctxtEINE-EKORG").text = "1500"
     sapSession.findById("wnd[0]/usr/ctxtEINE-WERKS").text = plantCode
     sapSession.findById("wnd[0]").sendVKey 0
End Sub

Private Sub changeME15(cancel As Boolean)
     If handleME15Text() = 0 Then Exit Sub
     If InStr(sapSession.findById("wnd[0]/sbar/pane[0]").text, "não existe") Then Exit Sub
     sapSession.findById("wnd[0]/usr/chkEINA-LOEKZ").Selected = cancel
     sapSession.findById("wnd[0]/usr/chkEINE-LOEKZ").Selected = cancel
     sapSession.findById("wnd[0]").sendVKey 11
     End Sub
     
Private Function handleME15Text()
     Dim statusBar As String: statusBar = sapSession.findById("wnd[0]/sbar/pane[0]").text
     handleME15Text = 1
     If InStr(statusBarText, "o existe") > 0 Then handleME15Text = 0
     If InStr(statusBarText, "dados de organiz") > 0 Then
          Call exitPages
          handleME15Text = 0
          End If
End Function

Private Sub cancelME01(materialList As Variant, rowIndex As Integer)
     If materialList(rowIndex, 3) = "AMBOS" Then
          Call fillME01(materialList, rowIndex, "0212")
          Call fillME01(materialList, rowIndex, "0304")
     Else
          Call fillME01(materialList, rowIndex, CStr(materialList(rowIndex, 3)))
     End If
End Sub

Private Sub fillME01(materialList As Variant, rowIndex As Integer, plantCode As String)
     sapSession.findById("wnd[0]/usr/ctxtEORD-MATNR").text = materialList(rowIndex, 1)
     sapSession.findById("wnd[0]/usr/ctxtEORD-WERKS").text = plantCode
     sapSession.findById("wnd[0]").sendVKey 0
     Call insideME01(materialList, rowIndex)
End Sub
Private Sub insideME01(materialList As Variant, rowIndex As Integer)
     If InStr(sapSession.findById("wnd[0]/sbar/pane[0]").text, "Bloqueado Somente Entrada") Then
          materialList(rowIndex, 4) = "Mat Bloqueado"
          Exit Sub
     End If
     Dim supplier As String: supplier = materialList(rowIndex, 2)
     If sapSession.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").text = supplier Then
          sapSession.findById("wnd[0]/usr/tblSAPLMEORTC_0205").getAbsoluteRow(0).Selected = True
          sapSession.findById("wnd[0]").sendVKey 14
               sapSession.findById("wnd[1]/usr/btnSPOP-OPTION1").press
          sapSession.findById("wnd[0]").sendVKey 11
     Else
          Call exitPages
     End If
     materialList(rowIndex, 4) = "Cancelado"
End Sub

Private Sub mainUpdateValuesME12(initialRange As Range, Optional changeType As Integer = 0)
     Dim rowIndex As Integer
     Dim changeText As String: changeText = IIf(changeType = 0, "Alt. valor Manualmente", "Alt. IVA")
     Dim materialList() As Variant
     
     Call initializeMaterialList(materialList, 2, initialRange)
     Call handleInitializeBasics
     Call navigateToSAPTable("ME12")
     For rowIndex = 1 To UBound(materialList)
          Call UpdateValuesME12(materialList, rowIndex, changeType)
     Next rowIndex
     Call exitPages
End Sub

Private Sub UpdateValuesME12(materialList As Variant, rowIndex As Integer, changeType As Integer)
     If changeType = 1 Then
          Call populateHeaderME(CDbl(materialList(rowIndex, 1)))
          Call updateIVAcodeME12Internal(CStr(materialList(rowIndex, 2)))
          Call populateHeaderME(CDbl(materialList(rowIndex, 1)), , "0304")
          Call updateIVAcodeME12Internal(CStr(materialList(rowIndex, 2)))
     Else
          Call populateHeaderME(CDbl(materialList(rowIndex, 1)))
          Call updatePriceValueME12Internal(CStr(materialList(rowIndex, 2)))
          Call populateHeaderME(CDbl(materialList(rowIndex, 1)), , "0304")
          Call updatePriceValueME12Internal(CStr(materialList(rowIndex, 2)))
     End If
End Sub

Private Sub updateIVAcodeME12Internal(ivaCode As String)
     ivaCode = Replace(ivaCode, " ", "")
     sapSession.findById("wnd[0]").sendVKey 0
     sapSession.findById("wnd[0]/usr/ctxtEINE-MWSKZ").text = ivaCode
     sapSession.findById("wnd[0]").sendVKey 11
End Sub

Private Sub updatePriceValueME12Internal(newPriceValue As Double)
     newPriceValue = newPriceValue
     sapSession.findById("wnd[0]").sendVKey 8
     sapSession.findById("wnd[1]").sendVKey 7
     sapSession.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/txtKONP-KBETR[2,0]").text = newPriceValue
     sapSession.findById("wnd[0]").sendVKey 11
End Sub






' Working with ZI9_MM_RegInfo

Private Sub updateLoad()
     Dim initialRange As Range: Set initialRange = Range("G10")
     Dim plantCode As String: plantCode = Range("I7").Value
     Dim loadcode As String
     
     Call initializeMaterialList(materialList, 2, initialRange)
     Call roundValuesMatrixes(materialList)
     If plantCode = "AMBOS" Then
          Call handleInitializeBasics
          Call updateLoadPlantCode(materialList, "0212")
          loadcode = getLoadCode
          
          Sheets("Cadastros - SAP").Activate
          Call updateLoadPlantCode(materialList, "0304")
          loadcode = getLoadCode
     Else
          Call handleInitializeBasics
          Call updateLoadPlantCode(materialList, plantCode)
          
          loadcode = getLoadCode
     End If
End Sub

Private Sub updateLoadPlantCode(materialList As Variant, Optional plantCode As String = "0212")
     Dim rowIndex As Integer: rowIndex = 0

     Call navigateToSAPTable("ZI9_MM_REGINFO")
     Call populateHeaderInputsZI9(plantCode)
     Call ResizeAndCopyData(materialList)
     Call populateMaterialFromClipboard
     For rowIndex = 0 To UBound(materialList) - 1
          If CDbl(sapSession.findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").getcellvalue(rowIndex, "MATNR")) <> CDbl(materialList(rowIndex + 1, 1)) Then MsgBox ("Erro")
          
          sapSession.findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").firstVisibleRow = rowIndex
          sapSession.findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").modifyCell rowIndex, "ZPB0", materialList(rowIndex + 1, 2)
          Next rowIndex
     sapSession.findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").triggerModified
     Call populateHeaderZI9(plantCode)
End Sub

Private Sub populateHeaderInputsZI9(Optional plantCode As String = "0212")
     sapSession.findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSEKORG").text = "1500"
     sapSession.findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSLIFNR").text = sapSupplier
     sapSession.findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSWERKS-LOW").text = plantCode
End Sub

Private Sub populateMaterialFromClipboard()
     sapSession.findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/btn%_SMATNR_%_APP_%-VALU_PUSH").press
          sapSession.findById("wnd[1]/tbar[0]/btn[24]").press
          sapSession.findById("wnd[1]/tbar[0]/btn[8]").press
     sapSession.findById("wnd[0]").sendVKey 8
End Sub

Private Sub roundValuesMatrixes(matrixToBeRound As Variant, Optional columnIndex As Integer = 2)
     Dim rowIndex As Integer
     For rowIndex = 1 To UBound(matrixToBeRound)
          matrixToBeRound(rowIndex, columnIndex) = Round(matrixToBeRound(rowIndex, columnIndex), 2)
          Next rowIndex
End Sub

Private Sub populateHeaderZI9(Optional plantCode As String = "0212")
     sapSession.findById("wnd[0]/usr/txtCPO_CENTRO").text = plantCode
     sapSession.findById("wnd[0]/usr/txtCPO_TEXT").text = commonTextUsed
     sapSession.findById("wnd[0]/tbar[1]/btn[8]").press
End Sub

Private Function getLoadCode()
     getLoadCode = Mid(sapSession.findById("wnd[0]/sbar").text, 28, 4)
     Call exitPages
End Function

Private Sub importPriceReportSAP()
     Call handleInitializeBasics
     Call navigateToSAPTable("ZI9_MM_REGINFO")
     Call importPriceZI9
     Call importPriceZI9("0304")
     Call exitPages(1)
     Call Importar_AP2
End Sub

Private Sub importPriceZI9(Optional plantCode As String = "0212")
     Dim resultWorksheet As Worksheet: Set resultWorksheet = ThisWorkbook.Sheets(plantCode)
     Dim wsRANGE As Range: Set wsRANGE = resultWorksheet.Range("B1")
     
     resultWorksheet.Columns("B:R").ClearContents
     Application.CutCopyMode = False
     Call populateHeaderInputsZI9(plantCode)
     sapSession.findById("wnd[0]").sendVKey 8
     If InStr(sapSession.findById("wnd[0]/sbar/pane[0]").text, "Não") Then Exit Sub
     Call exportSAPtoClipboard
     wsRANGE.PasteSpecial
     Application.CutCopyMode = False
     resultWorksheet.Range("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="|", TrailingMinusNumbers:=True
     Call exitPages
End Sub

Private Sub createNewPriceCodeSAP()
     Dim initialRange As Range: Set initialRange = Range("K10")
     Dim plantCode As String: plantCode = Range("M7").Value
     Dim loadcode As String
     Dim checkCode As Boolean
     
     Call handleInitializeBasics
     Call initializeMaterialList(materialList, 3, initialRange)
     Call roundValuesMatrixes(materialList)
     loadcode = handleXLStoSAP(materialList, plantCode)
     
     checkCode = loopingUntilCheckIsDone(loadcode)
     
     If checkCode Then Call useQuotationCodeinZI9(Range("O11").Value, plantCode)
End Sub

Private Function handleXLStoSAP(materialList As Variant, Optional plantCode As String = "0212")

     ReDim matrixtoFill(1 To UBound(materialList), 1 To 9)
     Dim xlsPath As String: xlsPath = Range("caminho_xlsx").Value
     Dim loadcode As String
     
     Call createQuotationXLS(materialList, matrixtoFill, plantCode)
     Call createXLS(matrixtoFill, xlsPath)
     handleXLStoSAP = sendXLStoSAP(xlsPath)
     
     Call handleErrorXls
     
End Function

Private Sub handleErrorXls()
     If InStr(sapSession.findById("wnd[0]/sbar/pane[0]").text, "Dados inconsistentes") = 0 Then Exit Sub
     Call exportSAPtoClipboard(2)
     Call exportSAPToLog
     Call resumelog

End Sub

Private Sub resumelog()
     Dim Worksheet As Worksheet: Set Worksheet = Sheets("LOG")
     Dim lastRow As Long: lastRow = Cells(Rows.Count, "Q").End(xlUp).Row
     Dim log As String: log = ""
     Dim cellValue As String
     
     For Each cell In Range("Q6:Q" & lastRow)
          cellValue = Trim(cell.Value)
          If cellValue <> "" Then log = log & cellValue & vbCrLf
     Next cell
     
     MsgBox "Erro conforme: " & vbCrLf & log
     Range("A1").Value = log
     End
End Sub

Private Sub createQuotationXLS(materialList As Variant, matrixtoFill As Variant, plantCode As String)
     Dim rowIndex As Integer
     Dim enterpriseCode As String: enterpriseCode = IIf(plantCode = "0212", "0200", "0300")
     ReDim matrixtoFill(1 To UBound(materialList), 1 To 9)
     
     For rowIndex = 1 To UBound(materialList)
          matrixtoFill(rowIndex, 1) = "1500"
          matrixtoFill(rowIndex, 2) = "103"
          matrixtoFill(rowIndex, 3) = enterpriseCode
          matrixtoFill(rowIndex, 4) = materialList(rowIndex, 1)
          matrixtoFill(rowIndex, 5) = "1"
          matrixtoFill(rowIndex, 6) = plantCode
          matrixtoFill(rowIndex, 7) = sapSupplier
          matrixtoFill(rowIndex, 8) = materialList(rowIndex, 2)
          matrixtoFill(rowIndex, 9) = materialList(rowIndex, 3)
          Next rowIndex
End Sub

Private Sub createXLS(matrixtoFill As Variant, xlsPath As String)
     Workbooks.Open (xlsPath)
     Range(Range("A2:I2"), Range("A2:I2").End(xlDown)).ClearContents
     Range("A2").Resize(UBound(matrixtoFill, 1), UBound(matrixtoFill, 2)).Value = matrixtoFill
     Windows("Template - Cotacao.xlsx").Close SaveChanges:=True
End Sub

Private Function sendXLStoSAP(xlsPath As String)
     Call navigateToSAPTable("ZLBRR_MM_0003")
     sapSession.findById("wnd[0]/usr/ctxtP_FILE").text = xlsPath
     sapSession.findById("wnd[0]").sendVKey 8
     sapSession.findById("wnd[0]").sendVKey 20
     sendXLStoSAP = Right$(sapSession.findById("wnd[0]/sbar").text, 4)
     Application.Wait (Now + TimeValue("0:00:05"))
End Function
Private Sub openCreateQuotationPage()
     sapSession.findById("wnd[0]").sendVKey 25
End Sub
Private Sub checkZLBR003forStatus(loadcode As String)
     sapSession.findById("wnd[0]/usr/ctxtS_PROC-LOW").text = loadcode
     sapSession.findById("wnd[0]").sendVKey 8
     Call exportSAPtoClipboard(2)
     Application.Wait (Now + TimeValue("0:00:02"))
End Sub

Private Sub exportSAPToLog()
     Dim Worksheet As Worksheet: Set Worksheet = Sheets("LOG")
     ThisWorkbook.Activate
     Worksheet.Range("B1").PasteSpecial
     Worksheet.Range("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="|", TrailingMinusNumbers:=True
     Application.Wait (Now + TimeValue("0:00:01"))
End Sub

Private Function checkifQuotationCodeIsExported() As Boolean
     Dim Worksheet As Worksheet: Set Worksheet = Sheets("Cadastros - SAP")
     If IsError(Range("P7").Value) Then
          checkifQuotationCodeIsExported = False
          Exit Function
          End If
          
     If Left(Range("P7").Value, 1) <> "6" Then
          checkifQuotationCodeIsExported = False
          Exit Function
     Else
          Range("O11").Value = Range("P7").Value
          Range("P11").Value = ""
          checkifQuotationCodeIsExported = True
     End If
End Function

Private Function loopingUntilCheckIsDone(loadcode As String) As Boolean
     Dim iterator As Integer
     Dim checkIfExported As Boolean: checkIfExported = False
     openCreateQuotationPage
     Do While iterator < 4
          Call checkZLBR003forStatus(loadcode)
          Call exportSAPToLog
          Call exitPages
          checkIfExported = checkifQuotationCodeIsExported
          iterator = IIf(checkIfExported, 4, iterator + 1)
          Application.Wait (Now + TimeValue("0:00:02"))
     Loop
     loopingUntilCheckIsDone = checkIfExported
End Function

Sub useQuotationCodeinZI9(quotationCode As String, plantCode As String)
     Call initializeSAPConnection
     Call navigateToSAPTable("ZI9_MM_REGINFO")

     sapSession.findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC2").Select
     sapSession.findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC2/ssubTBS_100_SCA:ZI9_MM_REGINFO:0102/subSBS_0105:ZI9_MM_REGINFO:0105/ctxtS_EBELN-LOW").text = quotationCode
     sapSession.findById("wnd[0]").sendVKey 8
     sapSession.findById("wnd[0]/usr/txtCPO_TEXT").text = Range("H8").Value
     sapSession.findById("wnd[0]/tbar[1]/btn[8]").press
     quotationCode = Right$(Left$(sapSession.findById("wnd[0]/sbar").text, 31), 4)
End Sub
