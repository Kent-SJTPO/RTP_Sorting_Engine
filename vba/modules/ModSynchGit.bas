Attribute VB_Name = "ModSynchGit"
Option Compare Database
Option Explicit

'============================================================
' modSynchGit
'
' Exports:
'   VBA modules/classes/forms/reports
'   Saved SQL queries
'
' Uses:
'   tblRepoSettings
'       SettingName = RepoRoot
'       SettingValue = path to repo root
'
' Creates:
'   \backup
'   \vba\modules
'   \vba\classes
'   \vba\forms
'   \vba\reports
'   \sql
'
' Keeps:
'   one .bak copy before overwrite
'============================================================

Public Sub ExportAccessSourceToRepo()

    On Error GoTo ErrHandler

    Dim repoRoot As String

    repoRoot = Nz(DLookup("SettingValue", _
                          "tblRepoSettings", _
                          "SettingName='RepoRoot'"), "")

    If repoRoot = "" Then
        MsgBox "RepoRoot not found in tblRepoSettings.", vbExclamation
        Exit Sub
    End If

    EnsureFolder repoRoot
    EnsureFolder repoRoot & "\backup"
    EnsureFolder repoRoot & "\vba"
    EnsureFolder repoRoot & "\vba\modules"
    EnsureFolder repoRoot & "\vba\classes"
    EnsureFolder repoRoot & "\vba\forms"
    EnsureFolder repoRoot & "\vba\reports"
    EnsureFolder repoRoot & "\sql"

    ExportVBAComponents repoRoot
    ExportSavedQueries repoRoot

    MsgBox "Git sync complete.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Git sync failed: " & Err.Description, vbCritical

End Sub


Private Sub ExportVBAComponents(ByVal repoRoot As String)

    Dim vbComp As Object
    Dim outPath As String
    Dim fileExt As String
    Dim folderPath As String

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents

        Select Case vbComp.Type

            Case 1
                folderPath = repoRoot & "\vba\modules"
                fileExt = ".bas"

            Case 2
                folderPath = repoRoot & "\vba\classes"
                fileExt = ".cls"

            Case 3
                folderPath = repoRoot & "\vba\forms"
                fileExt = ".cls"

            Case 100
                If Left(vbComp.Name, 6) = "Report" Then
                    folderPath = repoRoot & "\vba\reports"
                ElseIf Left(vbComp.Name, 4) = "Form" Then
                    folderPath = repoRoot & "\vba\forms"
                Else
                    folderPath = repoRoot & "\vba\classes"
                End If
                fileExt = ".cls"

            Case Else
                folderPath = repoRoot & "\vba\classes"
                fileExt = ".cls"

        End Select

        outPath = folderPath & "\" & CleanFileName(vbComp.Name) & fileExt

        BackupFile repoRoot, outPath

        If FileExists(outPath) Then Kill outPath

        vbComp.Export outPath

        Debug.Print "Exported: " & vbComp.Name

    Next vbComp

End Sub


Private Sub ExportSavedQueries(ByVal repoRoot As String)

    Dim qdf As DAO.QueryDef
    Dim outPath As String
    Dim fileNum As Integer

    For Each qdf In CurrentDb.QueryDefs

        If ShouldExportQuery(qdf.Name) Then

            outPath = repoRoot & "\sql\" & CleanFileName(qdf.Name) & ".sql"

            BackupFile repoRoot, outPath

            If FileExists(outPath) Then Kill outPath

            fileNum = FreeFile
            Open outPath For Output As #fileNum
            Print #fileNum, qdf.sql
            Close #fileNum

            Debug.Print "Exported: " & qdf.Name

        End If

    Next qdf

End Sub


Private Sub BackupFile(ByVal repoRoot As String, ByVal filePath As String)

    Dim backupPath As String
    Dim fileName As String

    If Not FileExists(filePath) Then Exit Sub

    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    backupPath = repoRoot & "\backup\" & fileName & ".bak"

    If FileExists(backupPath) Then Kill backupPath

    Name filePath As backupPath

End Sub


Private Function ShouldExportQuery(ByVal queryName As String) As Boolean

    If Left(queryName, 1) = "~" Then
        ShouldExportQuery = False
    ElseIf Left(queryName, 4) = "MSys" Then
        ShouldExportQuery = False
    Else
        ShouldExportQuery = True
    End If

End Function


Private Sub EnsureFolder(ByVal folderPath As String)

    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If

End Sub


Private Function FileExists(ByVal filePath As String) As Boolean

    FileExists = (Len(Dir(filePath)) > 0)

End Function


Private Function CleanFileName(ByVal rawName As String) As String

    Dim s As String

    s = rawName
    s = Replace(s, "\", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, """", "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")

    CleanFileName = s

End Function

