Attribute VB_Name = "Module1"
Sub GetEmailInfo()

Dim outlookApp As Object, testInspect As Inspector, numEmails As Integer, eMail As MailItem
Dim targetEmail As MailItem, potEmails As Integer, crmURL As String, pos1 As Integer, pos2 As Integer
Dim recipAdd As String, custName As String, stAdd As String, zipAdd As String, ctryAdd As String
Dim machModel As String, contName As String, nSplit As Integer, nSplitTest As String, lineNum As Integer
Dim emailText As String, i As Integer, j As Integer, k As Integer, salesExec As String
Dim commaPos As Integer, parenthPos As Integer

On Error Resume Next
Set outlookApp = GetObject(, "Outlook.Application")
On Error GoTo 0

If outlookApp Is Nothing Then
    Exit Sub
End If

Set testInspect = outlookApp.ActiveInspector

If testInspect Is Nothing Then
    Exit Sub
End If

numEmails = outlookApp.Inspectors.Count

If numEmails = 0 Then
    Exit Sub
End If

potEmails = 0 'number of potential emails
For i = 1 To numEmails
    Set eMail = outlookApp.Inspectors.Item(i).CurrentItem
    For Each recip In eMail.Recipients
        If recip.AddressEntry.AddressEntryUserType = 1 Then 'distribution list
            recipAdd = recip.AddressEntry.GetExchangeDistributionList.PrimarySmtpAddress
            If recipAdd = "DLPSA-CallReportDistribution@PSAngelus.com" Then
                Set targetEmail = eMail 'Distr. list is a recipient
                potEmails = potEmails + 1
                Exit For
            End If
        End If
    Next recip
Next

If potEmails > 1 Then 'more than one open email addressed to Distr. List
    'Set targetEmail = Nothing 'reset targetemail
    potEmails = 0 'reset potemails
    For i = 1 To numEmails
        Set eMail = outlookApp.Inspectors.Item(i).CurrentItem
        For Each recip In eMail.Recipients
            If recip.AddressEntry.AddressEntryUserType = 1 Then 'distribution list
                recipAdd = recip.AddressEntry.GetExchangeDistributionList.PrimarySmtpAddress
                If recipAdd = "DLPSA-CallReportDistribution@PSAngelus.com" Then
                    If InStr(1, UCase(eMail.Body), "MATEER") > 0 Or InStr(1, UCase(eMail.Body), "AUGER") > 0 Then
                        Set targetEmail = eMail
                        potEmails = potEmails + 1
                        Exit For
                    End If
                End If
            End If
        Next recip
    Next
ElseIf potEmails = 0 Then 'email isn't open, or isn't addressed to DL
    For i = 1 To numEmails
        Set eMail = outlookApp.Inspectors.Item(i).CurrentItem
        If InStr(1, UCase(eMail.Body), "MATEER") > 0 Or InStr(1, UCase(eMail.Body), "AUGER") > 0 Then
            Set targetEmail = eMail 'Distr. list is a recipient
            potEmails = potEmails + 1
        End If
    Next
End If

'If potEmails > 1 after this step Then the last one processed will be used

If potEmails = 0 Or targetEmail Is Nothing Then
    Range("A1").Value = "Nothing"
    Exit Sub
End If

salesExec = targetEmail.Sender 'England, Tyer (PSA-CLW)
commaPos = InStr(targetEmail.Sender, ",")
parenthPos = InStr(targetEmail.Sender, "(") - 1
salesExec = Mid(targetEmail.Sender, commaPos + 2, parenthPos - commaPos - 2)
salesExec = salesExec & " " & Left(targetEmail.Sender, commaPos - 1)

pos1 = InStr(1, targetEmail.Body, "Link:") + 6
pos2 = InStr(1, targetEmail.Body, "Sales") - 3
crmURL = Mid(targetEmail.Body, pos1, pos2 - pos1)
crmURL = Application.WorksheetFunction.Trim(Replace(crmURL, vbLf, ""))

nSplitTest = ""
nSplit = 0
Do While nSplitTest <> "Account"

    nSplit = nSplit + 1
    nSplitTest = Split(targetEmail.Body, ":")(nSplit)
    nSplitTest = Right(nSplitTest, 7)
    nSplitTest = Application.WorksheetFunction.Trim(Replace(nSplitTest, vbLf, ""))

    If nSplit > 10 Then
        Exit Do
    End If
Loop

nSplit = nSplit + 1
lineNum = 0

Do While custName = ""
    lineNum = lineNum + 1
    custName = Split(Split(targetEmail.Body, ":")(nSplit), vbCr)(lineNum)
    custName = Application.WorksheetFunction.Trim(Replace(custName, vbLf, ""))
Loop

Do While stAdd = ""
    lineNum = lineNum + 1
    stAdd = Split(Split(targetEmail.Body, ":")(nSplit), vbCr)(lineNum)
    stAdd = Application.WorksheetFunction.Trim(Replace(stAdd, vbLf, ""))
Loop

Do While zipAdd = ""
    lineNum = lineNum + 1
    zipAdd = Split(Split(targetEmail.Body, ":")(nSplit), vbCr)(lineNum)
    zipAdd = Application.WorksheetFunction.Trim(Replace(zipAdd, vbLf, ""))
Loop

Do While ctryAdd = ""
    lineNum = lineNum + 1
    ctryAdd = Split(Split(targetEmail.Body, ":")(nSplit), vbCr)(lineNum)
    ctryAdd = Application.WorksheetFunction.Trim(Replace(ctryAdd, vbLf, ""))
Loop

pos1 = InStr(1, targetEmail.Body, "Contacts:") + 12
pos2 = InStr(1, targetEmail.Body, "Phone:") - 2
If pos2 - pos1 > 0 Then
    contName = Mid(targetEmail.Body, pos1, pos2 - pos1)
    contName = Application.WorksheetFunction.Trim(Replace(contName, vbCrLf, ""))
Else
    contName = " "
End If

emailText = Split(targetEmail.Body, crmURL)(0) & " " & Split(targetEmail.Body, crmURL)(1) 'don't consider the link

If InStr(1, UCase(emailText), "MLX") > 0 Then 'model was specified
    pos2 = InStr(1, UCase(emailText), "MLX") + 3
    j = 0
    For i = pos2 To 1 Step -1
        If Not IsNumeric(Mid(emailText, i, 1)) Then
            If j = 1 Then
               pos1 = i
               Exit For
            End If
        Else
            j = 1
        End If
    Next i
    machModel = Mid(emailText, pos1, pos2 - pos1)
    machModel = Application.WorksheetFunction.Trim(machModel)
    
Else
    If InStr(1, emailText, "1800") > 0 Then '1800
        If InStr(1, UCase(emailText), "1800B") > 0 Or InStr(1, UCase(emailText), "BELT") > 0 Then 'belt drive
            If InStr(1, UCase(emailText), "1800B/D") > 0 Then 'not specified
                machModel = "1800D MLX"
            Else 'belt drive
                machModel = "1800B MLX"
            End If
        Else 'not belt drive
            machModel = "1800D MLX"
        End If
    ElseIf InStr(1, emailText, "1900") > 0 Then '1900
        If InStr(1, UCase(emailText), "CERAMIC") > 0 Or InStr(1, UCase(emailText), "1900C") Then 'ceramic
            machModel = "1900C MLX"
        Else 'normal/standard
            machModel = "1900 MLX"
        End If
    ElseIf InStr(1, UCase(emailText), "SEMI") > 0 Then 'semi
        If InStr(1, UCase(emailText), "CLUTCH") > 0 Then 'clutch/brake
            machModel = "1900 MLX"
        Else 'servo
            machModel = "1800D MLX"
        End If
    ElseIf InStr(1, UCase(emailText), "ROTARY") > 0 Then 'rotary
        If InStr(1, emailText, "6600") > 0 Then '6600
            machModel = "6600"
        Else '6700
            machModel = "6700"
        End If
    ElseIf InStr(1, UCase(emailText), "AUTO") > 0 Then 'auto
        If InStr(1, UCase(emailText), "SINGLE HEAD") Then
            If InStr(1, UCase(emailText), "CLUTCH") > 0 Then
                machModel = "3900 MLX"
            Else
                machModel = "3800D MLX"
            End If
        ElseIf InStr(1, UCase(emailText), "DUAL HEAD") Then
            If InStr(1, UCase(emailText), "CLUTCH") > 0 Then
                machModel = "4900 MLX"
            Else
                machModel = "4800D MLX"
            End If
        End If
        
        If machModel = "" Then 'single/dual head don't appear
            For i = 3 To 4
                For j = 8 To 9
                    For k = 1 To 3
                        If InStr(1, emailText, Str(i) & Str(j) & Str(k) & "0") > 0 Then
                            machModel = Str(i) & Str(j) & Str(k) & "0"
                            If j = 8 Then
                                machModel = machModel & "D MLX"
                            Else
                                machModel = machModel & " MLX"
                            End If
                        End If
                    Next k
                Next j
            Next i
        End If
    End If
End If

Range("A1").Value = crmURL
Range("A2").Value = custName
Range("A3").Value = stAdd
Range("A4").Value = zipAdd
Range("A5").Value = ctryAdd
Range("A6").Value = contName
Range("A7").Value = machModel
Range("A8").Value = salesExec

Exit Sub

errhandler:
MsgBox "Error in GetEmailInfo sub"

End Sub
