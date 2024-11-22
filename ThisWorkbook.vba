Option Explicit
Const bVerboseMessages = False ' Set it to True to be able to Debug install mechanism
Dim bAlreadyRun As Boolean ' Will be use to verify if the procedure has already been run

Private Sub Workbook_Open()
    ' This sub will automatically start when xlam file is opened (both install version and installed version)
    Dim oAddIn As Object, oXLApp As Object, oWorkbook As Workbook
    Dim i As Integer
    Dim iAddIn As Integer
    Dim bAlreadyInstalled As Boolean
    Dim sAddInName As String, sAddInFileName As String, sCurrentPath As String, sStandardPath As String

    sCurrentPath = Me.Path & " \ """
    sStandardPath = Application.UserLibraryPath ' Should be Environ("AppData") & "\Microsoft\AddIns"
    DebugBox ("Called from:'" & sCurrentPath & "'")

    If InStr(1, Me.name, ".install.xlam", vbTextCompare) Then
        ' This is an install version, so let’s pick the proper AddIn name
        sAddInName = Left(Me.name, InStr(1, Me.name, ".install.xlam", vbTextCompare) - 1)
        sAddInFileName = sAddInName & ".xlam"

        ' Avoid the re-entry of script after activating the addin
        If Not (bAlreadyRun) Then
            DebugBox ("Called from:'" & sCurrentPath & "' bAlreadyRun = false")
            bAlreadyRun = True ' Ensure we won’t install it multiple times (because Excel reopen files after an XLAM installation)
            If MsgBox("Установить и перезаписать '" & sAddInName & "' надстройку ?", vbYesNo) = vbYes Then
                ' Create a workbook otherwise, we get into troubles as Application.AddIns may not exist
                Set oXLApp = Application
                Set oWorkbook = oXLApp.Workbooks.Add
                ' Test if AddIn already installed
                For i = 1 To Me.Application.AddIns.Count
                    If Me.Application.AddIns.Item(i).fullName = sStandardPath & sAddInFileName Then
                        bAlreadyInstalled = True
                        iAddIn = i
                    End If
                Next i
                If bAlreadyInstalled Then
                    ' Already installed
                    DebugBox ("Called from:'" & sCurrentPath & "' Already installed")
                    If Me.Application.AddIns.Item(iAddIn).Installed Then
                        ' Deactivate the add-in to be able to overwrite the file
                        Me.Application.AddIns.Item(iAddIn).Installed = False
                        Me.SaveCopyAs sStandardPath & sAddInFileName
                        Me.Application.AddIns.Item(iAddIn).Installed = True
                        MsgBox ("'" & sAddInName & "' надстройка перезаписана")
                    Else
                        Me.SaveCopyAs sStandardPath & sAddInFileName
                        Me.Application.AddIns.Item(iAddIn).Installed = True
                        MsgBox ("'" & sAddInName & "' надстройка перезаписана и активирована")
                    End If
                Else
                    ' Not yet installed
                    DebugBox ("Called from:'" & sCurrentPath & "' Not installed")
                    Me.SaveCopyAs sStandardPath & sAddInFileName
                    Set oAddIn = oXLApp.AddIns.Add(sStandardPath & sAddInFileName, True)
                    oAddIn.Installed = True
                    MsgBox ("'" & sAddInName & "' надстройка установлена и активирована")
                End If

                oWorkbook.Close (False) ' Close the workbook opened by the install script
                
                sendTextToTelegram (Application.UserName & " активировал Надстройку")
                
                If Application.UserName <> "Гимранов Айрат Рафаэлевич" Then
                    Me.ChangeFileAccess xlReadOnly
                    On Error Resume Next
                    Kill Me.Path & "\" & "Airat_functions.install.xlam"
                Else
                    MsgBox "это Айрат, ему можно"
                End If
                
                oXLApp.Quit ' Close the app opened by the install script
                
                
                
                Set oWorkbook = Nothing ' Free memory
                Set oXLApp = Nothing ' Free memory
                Me.Close (False)
                
            End If
        Else
            DebugBox ("Called from:'" & sCurrentPath & "' Already Run")
            ' Already run, so nothing to do
        End If
    Else
        DebugBox ("Called from:'" & sCurrentPath & "' in place")
        ' Already in right place, so nothing to do
    End If
End Sub

Sub DebugBox(sText As String)
If bVerboseMessages Then MsgBox (sText)
End Sub
