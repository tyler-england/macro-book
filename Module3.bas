Attribute VB_Name = "Module3"
Option Explicit

Public Sub ExportModules(wb As Workbook)

    If InStr(UCase(Application.UserName), "ENGLAND") = 0 Then Exit Sub
    If InStr(UCase(Application.UserName), "TYLER") = 0 Then Exit Sub
    
    Dim s1DPath As String, sFolderPath As String, sSubFolder As String, sFileFolder As String
    Dim varVar As Variant, bNewFolder As Boolean, sExt As String
    Dim sFailed() As String, x As Integer
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    '''''''hardcoded'''''''
    s1DPath = "C:\Users\englandt\*"
    '''''''''''''''''''''''
    On Error GoTo errhandler
    sSubFolder = Dir(Replace(s1DPath, "*", "OneDrive*"), vbDirectory)
    If sSubFolder = "" Then
        Exit Sub      'no OneDrive
    ElseIf UCase(sSubFolder) = "ONEDRIVE" Then
        MsgBox "May be using wrong OneDrive folder (not BW directory)"
    End If
    sFolderPath = Replace(s1DPath, "*", sSubFolder) & "\"
    sSubFolder = Dir(sFolderPath & "scripts*", vbDirectory)
    If sSubFolder = "" Then
        MkDir sFolderPath & "Scripts" 'make directory
        sSubFolder = "Scripts"
    End If
    sFolderPath = sFolderPath & sSubFolder & "\"
    sSubFolder = Dir(sFolderPath & "VBA*", vbDirectory)
    If sSubFolder = "" Then
        MkDir sFolderPath & "VBA_Modules" 'make directory
        sSubFolder = sFolderPath & "VBA_Modules"
    End If
    sFolderPath = sFolderPath & sSubFolder & "\" 'vba modules folder
    sFileFolder = Replace(Replace(Replace(Replace(Replace(wb.Path & "\" & wb.Name, "\", "-"), ".", ""), ":", "+"), " ", ""), "/", "-")
    sFileFolder = Replace(sFileFolder, "https+--bw1-mysharepointcom-personal-tyler_england_bwpackagingsystems_com-Documents", "OneDrive--")
    sSubFolder = Dir(sFolderPath & sFileFolder, vbDirectory)
    If sSubFolder = "" Then 'folder doesn't exist
        bNewFolder = True
        sSubFolder = Dir(sFolderPath & "*" & Replace(wb.Name, ".", "") & "*", vbDirectory)
        Do While sSubFolder <> "" 'check for any partial matches (diff path, etc)
            varVar = MsgBox("No folder exists with the following name..." & vbCrLf & sFileFolder & _
                    vbCrLf & vbCrLf & "However this folder does exist..." & vbCrLf & sSubFolder & _
                    vbCrLf & vbCrLf & "Do you want to rename the existing one and use it?", vbYesNo, "VBA Modules")
            If varVar = vbYes Then 'use this folder -> don't make a new one
                Name sFolderPath & sSubFolder As sFolderPath & sFileFolder
                bNewFolder = False
                Exit Do
            End If
            sSubFolder = Dir()
        Loop
        If bNewFolder Then 'make new folder
            MkDir sFolderPath & sFileFolder
        End If
        sFolderPath = sFolderPath & sFileFolder
    Else
        sFolderPath = sFolderPath & sSubFolder
    End If
    If Right(sFolderPath, 1) <> "\" Then sFolderPath = sFolderPath & "\"
    x = 0
    ReDim sFailed(x)
    For Each varVar In wb.VBProject.VBComponents
        On Error GoTo errhandler
        Select Case varVar.Type
            Case ClassModule, Document
                sExt = ".cls"
            Case Form
                sExt = ".frm"
            Case Module
                sExt = ".bas"
            Case Else
                sExt = ".txt"
        End Select
        If sExt = ".bas" Or sExt = ".cls" Then 'only care about modules/sheets
            On Error Resume Next
            Err.Clear
            Call varVar.Export(sFolderPath & varVar.Name & sExt)
            If Err.Number <> 0 Then
                ReDim Preserve sFailed(x)
                sFailed(x) = varVar.Name
                x = x + 1
            End If
        End If
    Next
    If x > 0 Then
        MsgBox "Failed to export the following modules:" & vbCrLf & vbCrLf & _
                Join(sFailed, vbCrLf)
    End If
errhandler:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End If
End Sub

Public Sub ErrorRep(rouName, rouType, curVal, errNum, errDesc, miscInfo)
    
    Dim oApp As Object, oEmail As MailItem, arrEmailTxt(10) As String
    Dim outlookOpen As Boolean, emailTxt As String, varMsg As Variant
    
    Application.ScreenUpdating = False
    arrEmailTxt(2) = "--Issue finding Workbook"
    arrEmailTxt(3) = "--Issue finding User"
    arrEmailTxt(4) = "--Issue finding Workbook path"
    arrEmailTxt(5) = "--Issue finding Routine name"
    arrEmailTxt(6) = "--Issue finding Routine type"
    arrEmailTxt(7) = "--Issue finding Current value"
    arrEmailTxt(8) = "--Issue finding Error number"
    arrEmailTxt(9) = "--Issue finding Error description"
    arrEmailTxt(10) = "--Issue finding Misc. add'l info"
    
    On Error Resume Next
        Set oApp = GetObject(, "Outlook.Application")
        outlookOpen = True
        
        ''''''can't use error handler because these varTypes might be problematic
        If Not VarType(curVal) = vbString Then 'make into string
            If VarType(curVal) > 8000 Then 'array of some sort
                curVal = Join(curVal, ";")
            Else 'hopefully this will make it a string
                curVal = Str(curVal)
            End If
        End If
        
        If Not VarType(miscInfo) = vbString Then 'make into string
            If VarType(miscInfo) > 8000 Then 'array of some sort
                curVal = Join(miscInfo, ";")
            Else 'hopefully this will make it a string
                curVal = Str(miscInfo)
            End If
        End If
        
    On Error Resume Next 'types might cause errors
        arrEmailTxt(0) = "REPORT"
        arrEmailTxt(1) = "Error occurred in VBA program. Details are listed below." & vbCrLf
        arrEmailTxt(2) = Right(arrEmailTxt(2), Len(arrEmailTxt(2)) - 16) & ": " & ThisWorkbook.Name
        arrEmailTxt(3) = Right(arrEmailTxt(3), Len(arrEmailTxt(3)) - 16) & ": " & Application.UserName & vbCrLf
        arrEmailTxt(4) = Right(arrEmailTxt(4), Len(arrEmailTxt(4)) - 16) & ": " & ThisWorkbook.Path
        arrEmailTxt(5) = Right(arrEmailTxt(5), Len(arrEmailTxt(5)) - 16) & ": " & rouName
        arrEmailTxt(6) = Right(arrEmailTxt(6), Len(arrEmailTxt(6)) - 16) & ": " & rouType
        arrEmailTxt(7) = Right(arrEmailTxt(7), Len(arrEmailTxt(7)) - 16) & ": " & curVal & vbCrLf
        arrEmailTxt(8) = Right(arrEmailTxt(8), Len(arrEmailTxt(8)) - 16) & ": " & errNum
        arrEmailTxt(9) = Right(arrEmailTxt(9), Len(arrEmailTxt(9)) - 16) & ": " & errDesc & vbCrLf
        arrEmailTxt(10) = Right(arrEmailTxt(10), Len(arrEmailTxt(10)) - 16) & ": " & vbCrLf & miscInfo
    On Error GoTo errhandler
    
    emailTxt = Join(arrEmailTxt, vbCrLf)
    
    'see if emailTxt has been sent already this session
    bNewMsg = True 'default value
    If iNumMsgs > 0 Then 'at least one email has been generated already
        For Each varMsg In arrErrorEmails 'see if there were any matches
            If UCase(varMsg) = UCase(emailTxt) Then 'this was already sent this session
                bNewMsg = False
                Exit For
            End If
        Next
    End If
    
    If bNewMsg Then 'new message -> add to array for next time
        iNumMsgs = iNumMsgs + 1
        ReDim Preserve arrErrorEmails(iNumMsgs)
        arrErrorEmails(iNumMsgs) = emailTxt
    Else 'repeat message
        Exit Sub
    End If
    
    If oApp Is Nothing Then
        Set oApp = CreateObject("Outlook.Application")
        outlookOpen = False
    End If
    
    Set oEmail = oApp.CreateItem(0)

    With oEmail
        .To = "tyler.england@bwpackagingsystems.com"
        .Subject = "VBA Program Error Report"
        .Body = emailTxt
        If InStr(UCase(Application.UserName), "ENGLAND, TYLER") > 0 Then
            .Display 'it me
        Else:
            .Send
        End If
    End With
    
    If Not outlookOpen Then oApp.Close
errhandler:
End Sub
