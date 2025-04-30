# verificar_proxima_linha_me01_sap
Buscar linha vazia na lista de lofs me01 sap


Sub PreencherProximaLinhaVazia_Click()

    Dim row As Integer
    Dim cellText As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Planilha1")
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
'------------------------------------------------------------------------------------------
    ' Conecta ao SAP
    If Not IsObject(Appl) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set Appl = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
       Set connection = Appl.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = connection.Children(0)
    End If
'------------------------------------------------------------------------------------------

    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n me01"
    session.findById("wnd[0]").sendVKey 0
    
    For i = 2 To ultimaLinha
    
        On Error Resume Next
           
        session.findById("wnd[0]/usr/ctxtEORD-MATNR").Text = ws.Cells(i, 1).Value
        session.findById("wnd[0]/usr/ctxtEORD-WERKS").Text = "BR35"
        session.findById("wnd[0]/usr/ctxtEORD-WERKS").SetFocus
        session.findById("wnd[0]/usr/ctxtEORD-WERKS").caretPosition = 4
        session.findById("wnd[0]").sendVKey 0
    
        row = 0
    
        Do
            On Error Resume Next
            cellText = session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2," & row & "]").Text
            If Err.Number <> 0 Then
                ' Se deu erro, provavelmente acabou as linhas visíveis
                Exit Do
            End If
            On Error GoTo 0
            
            If Trim(cellText) = "" Then
                ' Linha vazia encontrada, preencher
                session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-VDATU[0," & row & "]").Text = ws.Cells(i, 2).Value '"03.04.2025"
                session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-BDATU[1," & row & "]").Text = "31.12.9999"
                session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2," & row & "]").Text = ws.Cells(i, 3).Value '"101236"
                session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EKORG[3," & row & "]").Text = "BR35"
                'session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-AUTET[10," & row & "]").Text = "1"
                'session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-AUTET[10," & row & "]").SetFocus
                Exit Do
            End If
            
            row = row + 1
        Loop
        
        session.findById("wnd[0]").sendVKey 11
        
'------------------------------------------------------------------------------------------
        
        statusMessage = session.findById("wnd[0]/sbar").Text
        ws.Cells(i, "F").Value = statusMessage
            
        typeMessage = session.findById("wnd[0]/sbar").messagetype
        ws.Cells(i, "G").Value = typeMessage
            
        If typeMessage <> "S" Then
            ws.Cells(i, "E").Value = "Erro"
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/n me01"
            session.findById("wnd[0]").sendVKey 0
        Else
            ws.Cells(i, "E").Value = "Sucesso"
        End If
            
        application.Wait Now + TimeValue("00:00:03")
        
'------------------------------------------------------------------------------------------
        
    Next i
    
End Sub

