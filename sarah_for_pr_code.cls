'variables pa SAP
Public SapGuiAuto As Object
Public SAPApp As Object
Public SAPCon As Object
Public session As Object

'libros y worksheets
Public rpdFile As Workbook
Public sbFile As Workbook
Public rpdPath As String
Public sbPath As String
Public poPath As String

'variables de trabajo tecnico
Public prQnty As Integer
Public prWs As Worksheet
Public extraMail As String
Public info_error As Boolean
Public auto_correo As Boolean
Public txt As String
Public confirmation As String
Public needSap As Boolean

'variables de trabajo
Public pr As String
Public vc As String
Public payTerms As String
Public rs As String
Public poNumber As String

'variables extraidas de la me5a (rpd)
Public moneda As String
Public pgr As String
Public plant As String
Public req As String
Public tracknum As String
Public Buyer As String

'variables extraidas de MK03 (rpd)
Public duns As String
Public supplier As String
Public TermsPay As String
Public inco1 As String
Public inco2 As String

'diccionario de payments
Public dicTerminos As Object

'carpeta y eso
Public carpeta As String
Public archivo As String
Public destino As String
Public correoRequestor As String
Public buyer_mail As String
Public correoSupplier As String

Sub cargarArchivos()
    With Application.FileDialog(msoFileDialogFilePicker)
                .Title = "Selecciona el archivo RPO"
                If .Show = -1 Then
                    rpdPath = .SelectedItems(1)
                    Range("rpd_path").Value = rpdPath
                Else
                    Exit Sub
                End If
    End With

    With Application.FileDialog(msoFileDialogFilePicker)
                .Title = "Selecciona el archivo SB APPROVAL"
                If .Show = -1 Then
                    sbPath = .SelectedItems(1)
                    Range("sb_path").Value = sbPath
                Else
                    Exit Sub
                End If
    End With
End Sub

Sub limpiar()
    ThisWorkbook.Sheets("PR").Range("A2:F56").ClearContents
End Sub

Sub auto_pr()
    'se crea un diccionario para los incoterms
    Set dicTerminos = CreateObject("Scripting.Dictionary")
        ' Agregar términos al diccionario
        dicTerminos.Add "ZMN2", "MNS2-2: Payment Settled on 2nd day, 2nd mnth"
        dicTerminos.Add "Z23", "MNS2-3: 2nd day, third month"
        dicTerminos.Add "Z24", "MNS2-4: 2nd day, fourth month"
        dicTerminos.Add "ZMNX", "MNS2-5: 2nd day, 5th month"
        dicTerminos.Add "ZF45", "45 Days End of Month from Invoice (French)"
        dicTerminos.Add "ZCSH", "Cash In Advance"
        dicTerminos.Add "Z107", "Net 107 Days"
        dicTerminos.Add "Z120", "Net 120 Days"
        dicTerminos.Add "ZZ37", "Net 137 Days"
        dicTerminos.Add "Z015", "Net 15 days"
        dicTerminos.Add "Z025", "Net 25th Prox (pay on 25th of following month)"
        dicTerminos.Add "Z030", "Net 30 days"
        dicTerminos.Add "Z047", "Net 47 Days"
        dicTerminos.Add "Z056", "Net 56 days"
        dicTerminos.Add "Z060", "Net 60 days"
        dicTerminos.Add "Z077", "Net 77 days"
        dicTerminos.Add "Z090", "Net 90 days"
        dicTerminos.Add "ZCAD", "Net Immediate"


    'se crean variables utiles para el codigo
    Dim i As Integer
    Dim txt As String
    Dim coment As String

    'se crea conexion con sap
    confirmation = MsgBox("Necesitas el uso de SAP?", _
                          vbQuestion + vbYesNo, "SARAH")
    Application.ScreenUpdating = False
    
    If confirmation = vbYes Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set SAPApp = SapGuiAuto.GetScriptingEngine
        Set SAPCon = SAPApp.Children(0)
        Set session = SAPCon.Children.ElementAt(0)
        me5a_ff
        needSap = True
    Else
        needSap = False
    End If

    'se ajusta las variables de trabajo
    Set prWs = ThisWorkbook.Sheets("PR")
    prQnty = prWs.Range("prQnty").Value

    'variables de los libros a utilizar
    sbPath = prWs.Range("sb_path").Value
    rpdPath = prWs.Range("rpd_path").Value

    'bucle principal del programa
    For i = 2 To (prQnty + 1)
        auto_correo = True
        extraMail = ""
        txt = "RPO"
        
        coment = prWs.Range("F" & i).Value 'accion que se va realizar _
                                            con el pr
        pr = prWs.Range("A" & i).Value
        vc = prWs.Range("C" & i).Value  'variables para SAP
        rs = prWs.Range("H" & i).Value
        
        If coment = "In Process Center" Then
            poNumber = prWs.Range("N" & i).Value 'se toma el valor de PO
            
            If poNumber <> "" Then 'se revisa si existe
                sendpo (i)
            End If
    
            GoTo NextIteration
        ElseIf needSap = False Then
            GoTo NextIteration
        ElseIf coment = "rpd" Or coment = "xp" Then
            'se genera el rpd con correo
            auto_rpd (i) 'funcion rpd
            
            'funcion correo
            Savenmail (i)
            
            GoTo NextIteration
        ElseIf coment = "rpd nc" Or coment = "xp nc" Then
            auto_correo = False
            'se genera rpd sin correo
            auto_rpd (i) 'funcion rpd
            
            'funcion correo
            Savenmail (i)
            
            GoTo NextIteration
        ElseIf coment = "sb" Then
            txt = "SB APPROVAL"
            auto_correo = False
            'funcion approval

            prWs.Range("F" & i).Value = "rpd"
            Savenmail (i)
            GoTo NextIteration
        End If
        

NextIteration:
    Next i
End Sub

Function sendpo(i As Integer)  ' funcition
    Dim objOutlook As Object
    Dim objMail As Object
    
    buyer_mail = prWs.Range("L" & i).Value
    correoRequestor = prWs.Range("K" & i).Value
    correoSupplier = prWs.Range("J" & i).Value
    extraMail = prWs.Range("M" & i).Value
    
    confirmation = MsgBox("Desea enviar la PO '" & poNumber & "' a supplier?", _
                          vbQuestion + vbYesNo, "SARAH")
    Application.ScreenUpdating = False

    If confirmation = vbYes Then
        'se escoge el path de la hard copy
        With Application.FileDialog(msoFileDialogFilePicker)
                .Title = "Selecciona la Hard Copy de la PO: " & poNumber
                If .Show = -1 Then
                    poPath = .SelectedItems(1)
                Else
                    Exit Function
                End If
        End With
        ' Crear una instancia de Outlook
        Set objOutlook = CreateObject("Outlook.Application")
        
        ' Crear un nuevo correo
        Set objMail = objOutlook.CreateItem(0)
    
        With objMail
                .Subject = "APTIV Acknowledgment PO#: " & poNumber & " | " & prWs.Range("G" & i).Value
                .Body = "Dear Supplier: " & vbCrLf & _
                        "APTIV generated the the Spot Buy PO #: " & poNumber & vbCrLf & vbCrLf & _
                        "Please review it and confirm everything is correct and send me " _
                        & "your response as soon as possible. Otherwise, if we do not receive " _
                        & "your confirmation, we will assume the all the information containing " _
                        & "on the P.O. is correct and you do not have any discrepancies on it. " & vbCrLf & _
                        "Thank you."
                .To = correoSupplier  ' Utiliza la dirección de correo electrónico del Requestor
                .CC = "Roberto.Ontiveros@aptiv.com; " & correoRequestor & "; " & buyer_mail & "; " & extraMail ' Resto del código para adjuntar el archivo, enviar el correo...
                .Attachments.Add poPath
                .Display
        End With
        prWs.Range("F" & i).Value = "PO ENVIADA"
        prWs.Range("N" & i).Value = poNumber & " - Enviada"
    End If
End Function

Function Savenmail(i As Integer) 'funcion con la que se guarda el documento y se envia a mail
    Dim objOutlook As Object
    Dim objMail As Object
    Dim Adjunto As String
    Dim ArchivoActual As String
    Dim dicRequestors As Object
    Dim archivoFinal As String
    txt = "RPD"
    Set rpdFile = Workbooks.Open(rpdPath)

    buyer_mail = prWs.Range("L" & i).Value
    correoRequestor = prWs.Range("K" & i).Value
    extraMail = prWs.Range("M" & i).Value

    If correoRequestor = "Req. No registrado" Then
        auto_correo = False
        prWs.Range("F" & i).Value = "no enviado"
    End If


    If auto_correo = False Then
        rpdFile.Sheets(1).Range("Comment_Fld").Value = pr & " - " & tracknum & " - " & supplier & " - VC: " & vc
    End If

    destino = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    carpeta = Format(Date, "YYYY-MM-DD")
    destino = destino & "\" & carpeta & "\"

    If Dir(destino, vbDirectory) = "" Then
        MkDir destino
        destino = destino & "POs" & "\"
        MkDir destino
    End If

    archivo = txt & "-" & pr & "-" & tracknum & ".xlsm"
    rpdFile.SaveAs destino & archivo
    archivoFinal = destino & archivo

    If auto_correo = True Then
        ' Crear una instancia de Outlook
            Set objOutlook = CreateObject("Outlook.Application")
        
            ' Crear un nuevo correo
            Set objMail = objOutlook.CreateItem(0)
        
            ' Guardar el archivo actual
            ThisWorkbook.Save
        
            ArchivoActual = archivo_final
         
            ' Configurar el correo
            With objMail
                .Subject = "Request for Purchase – PO – CR - " & pr & " - " & tracknum & " - " & supplier & " - VC: " & vc
                .Body = "Recipient(s): " & vbCrLf & _
                        "If you are receiving this request as an Approver, please review the attached Request for Purchasing Document and provide your approval decision as feedback to the Requester [" & Buyer & "]." & vbCrLf & _
                        "If you are receiving this request as Process Center personnel, please proceed with processing the attached request as deemed appropriate." & vbCrLf & _
                        "Thank you."
                .To = "juarez.delphi.gsm.process.center.direct@aptiv.com"  ' Utiliza la dirección de correo electrónico del Requestor
                .CC = "Roberto.Ontiveros@aptiv.com; " & correoRequestor & "; " & buyer_mail & "; " & extraMail ' Resto del código para adjuntar el archivo, enviar el correo...
                .Attachments.Add archivoFinal
                .Display
            End With
    
            objMail.send
            rpdFile.Close
        Else
        End If
       
        ' Liberar recursos
        Set objMail = Nothing
        Set objOutlook = Nothing
        Set dicRequestors = Nothing
End Function

Function auto_rpd(i As Integer) 'para esta funcion la variable pr, cv y paytrms ya existe
    Dim texto As String
    Dim firstime As Boolean
    texto = "In Process Center"
    
    Set rpdFile = Workbooks.Open(rpdPath)
    
    'llamada a la transaccion me5a de SAP
    me5a
    prWs.Range("E" & i).Value = req
    'llamada a la transaccion mk03 de SAP
    mk03 'aqui mismo se llena el archivo
    
    'detector de errores
    check_rpd (i)
    
    If info_error = True Then
        texto = "Missing info"
        info_error = False
        auto_correo = False
    End If
    
    prWs.Range("F" & i).Value = texto
End Function

Function check_rpd(i As Integer)
    Set rpdFile = Workbooks.Open(rpdPath)

    If i = 2 Then
        If rpdFile.Sheets(1).Range("POrg_Fld").Value = "" Then
            MsgBox ("Error: El codigo no funciono correctamente, intente de nuevo mas tarde")
            rpdFile.Close
            End
        End If
    End If

    If rpdFile.Sheets(1).Range("venGS_Fld").Value = "" Or rpdFile.Sheets(1).Range("venCA_Fld").Value = "" Or _
    rpdFile.Sheets(1).Range("venFA_Fld").Value = "" Then
        info_error = True
    End If
End Function

Function mk03()
    Dim tax As String
    Dim error1 As String
    Dim terms As String

    session.StartTransaction ("MK03")
        session.findById("wnd[0]/usr/ctxtRF02K-LIFNR").Text = vc
        session.findById("wnd[0]/usr/ctxtRF02K-EKORG").Text = "dpe2"
        session.findById("wnd[0]/usr/chkRF02K-D0110").Selected = True
        session.findById("wnd[0]/usr/chkRF02K-D0120").Selected = True
        session.findById("wnd[0]/usr/chkRF02K-D0310").Selected = True
        session.findById("wnd[0]/usr/chkWRF02K-D0320").Selected = True

        session.findById("wnd[0]/tbar[0]/btn[0]").press

        duns = session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-SORT1").Text

        session.findById("wnd[0]/tbar[1]/btn[8]").press

        supplier = session.findById("wnd[0]/usr/txtLFA1_INT-NAME1").Text

        session.findById("wnd[0]/tbar[1]/btn[8]").press

        inco1 = session.findById("wnd[0]/usr/ctxtLFM1-INCO1").Text
        inco2 = session.findById("wnd[0]/usr/txtLFM1-INCO2").Text
        TermsPay = session.findById("wnd[0]/usr/ctxtLFM1-ZTERM").Text

        session.findById("wnd[0]/tbar[1]/btn[8]").press

        'incia llenado del formato
        Set rpdFile = Workbooks.Open(rpdPath)

        With rpdFile.Sheets(1) ' Asegúrate de ajustar el índice de hoja si es necesario
            .Range("venVN_Fld").Value = session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,2]").Text
            .Range("venOA_Fld").Value = session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,0]").Text
            .Range("venFA_Fld").Value = session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,1]").Text
            .Range("venCA_Fld").Value = session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,5]").Text
            .Range("venGS_Fld").Value = session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,6]").Text
            .Range("SupplierNm_Fld").Value = supplier
            .Range("Vendor_manf").Value = duns
            .Range("Div_Reg_Fld").Value = "SPS-NA"
            .Range("POrg_Fld").Value = "DPE2"
            .Range("Plant_Fld").Value = plant
            .Range("PGrp_Fld").Value = pgr
            
            .Range("PR_Num_Fld").Value = pr
            
            If plant = "HN11" Then
                tax = "I1"
            Else
                tax = "U0"
            End If
            
            .Range("TaxCode_Fld").Value = tax
            .Range("PayOn_Fld").Value = rs
            
            If dicTerminos.Exists(TermsPay) Then
                ' El código está en el diccionario, actualizar con el texto correspondiente
                TermsPay = dicTerminos(TermsPay)
            Else
                    ' El código no está en el diccionario, establecer valores predeterminados
                error1 = "Create PO with Payment terms: " & TermsPay
                terms = "Net 90 days"
                .Range("Comment_Fld").Value = error1
            End If
            
            .Range("PayTerms_Fld").Value = TermsPay
            .Range("Currency_Fld").Value = moneda
            .Range("INCO_Term").Value = inco1
            .Range("INCO_Txt").Value = inco2
            
            .Range("ValidityStart_Fld").FormulaR1C1 = "=TODAY()"
            .Range("ValidityStart_Fld").Value = .Range("ValidityStart_Fld").Value
            
            .Range("ValidityEnd_Fld").FormulaR1C1 = "=DATE(YEAR(ValidityStart_Fld),MONTH(ValidityStart_Fld)+6,DAY(ValidityStart_Fld))"
            .Range("ValidityEnd_Fld").Value = .Range("ValidityEnd_Fld").Value
            
            .Range("M58").Value = .Range("ValidityStart_Fld").Value
            .Range("G58").Value = .Range("ValidityStart_Fld").Value
        End With
End Function

Function me5a()
    session.StartTransaction ("ME5A") 'entra a la transccion
        session.findById("wnd[0]/usr/ctxtBA_BANFN-LOW").Text = pr 'se agrega el pr al buscador
        If Err.Number <> 0 Then
            session.findById("wnd[0]/usr/chkP_ERLBA").Selected = True 'Es usado nomas cuando es un pr ya procesado
        End If

        session.findById("wnd[0]/tbar[1]/btn[8]").press

        Err.Clear
        
        On Error Resume Next
            session.findById("wnd[0]/usr/lbl[3,6]").SetFocus
            session.findById("wnd[0]/usr/lbl[3,6]").caretPosition = 3
                
                If Err.Number <> 0 Then
                    Exit Function
                End If
                
            session.findById("wnd[0]").sendVKey 2
        On Error GoTo 0
        
        
        'seleccion contact person
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT11").Select

        'aqui se guarda si es emma o cesar
        Buyer = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT11/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:3328/txtMEREQ3328-EKNAM").Text

        plant = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").GetCellValue(0, "WERKS")
        pgr = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").GetCellValue(0, "EKGRP")
        req = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").GetCellValue(0, "AFNAM")
        tracknum = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").GetCellValue(0, "BEDNR")
        moneda = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").GetCellValue(0, "WAERS")
End Function

Function me5a_ff()
    'funcion pa limpiar sap
    session.StartTransaction ("ME5A")
        session.findById("wnd[0]/usr/ctxtBA_BANFN-LOW").Text = ""
        session.findById("wnd[0]/usr/ctxtBA_BANFN-LOW").caretPosition = 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtBA_EKGRP-HIGH").SetFocus
        session.findById("wnd[0]/usr/ctxtBA_EKGRP-HIGH").caretPosition = 0
        session.findById("wnd[0]/usr/btn%_BA_EKGRP_%_APP_%-VALU_PUSH").press
        session.findById("wnd[1]/tbar[0]/btn[16]").press
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[0]/usr/ctxtP_LSTUB").Text = "a"
        session.findById("wnd[0]/usr/ctxtP_LSTUB").SetFocus
        session.findById("wnd[0]/usr/ctxtP_LSTUB").caretPosition = 1
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtS_FRGDT-LOW").Text = ""
        session.findById("wnd[0]/usr/ctxtS_FRGDT-LOW").SetFocus
        session.findById("wnd[0]/usr/ctxtS_FRGDT-LOW").caretPosition = 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtS_FRGDT-HIGH").Text = ""
        session.findById("wnd[0]/usr/ctxtS_FRGDT-HIGH").SetFocus
        session.findById("wnd[0]/usr/ctxtS_FRGDT-HIGH").caretPosition = 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/chkP_ERLBA").Selected = False
        session.findById("wnd[0]/usr/chkP_ERLBA").SetFocus
    session.StartTransaction ("SE38")
End Function