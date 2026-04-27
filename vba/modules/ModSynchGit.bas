Attribute VB_Name = "ModSynchGit"
Option Compare Database
Option Explicit

'============================================================
' modExportAccessSource
'
' Purpose:
'   Export Access VBA modules and saved query SQL to text files
'   for Git/GitHub version tracking.
'
' Requirements:
'   In Access:
'   File > Options > Trust Center > Trust Center Settings > Macro Settings
'   Enable: "Trust access to the VBA project object model"
'
' Notes:
'   - Standard modules export as .bas
'   - Class modules export as .cls
'   - Form/report code modules export as .cls
'   - Saved queries export as .sql
'   - System/temp queries are skipped
'============================================================

Private Const REPO_ROOT As String = "c:\Users\Kschellinger\SJTPO_Git\RTP_Sorting_Engine"

Public Sub ExportAccessSourceToRepo()

    Dim exportRoot As String

    exportRoot = REPO_ROOT

    EnsureFolder exportRoot
    EnsureFolder exportRoot & "\vba"
    EnsureFolder exportRoot & "\vba\modules"
    EnsureFolder exportRoot & "\vba\classes"
    EnsureFolder exportRoot & "\vba\forms"
    EnsureFolder exportRoot & "\vba\reports"
    EnsureFolder exportRoot & "\sql"

    ExportVBAComponents exportRoot
    ExportSavedQueries exportRoot

    MsgBox "Access source export complete:" & vbCrLf & exportRoot, vbInformation

End Sub

Private Sub ExportVBAComponents(ByVal exportRoot As String)

    Dim vbComp As Object
    Dim outPath As String
    Dim fileExt As String
    Dim folderPath As String

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents

        Select Case vbComp.Type

            Case 1      ' Standard module
                folderPath = exportRoot & "\vba\modules"
                fileExt = ".bas"

            Case 2      ' Class module
                folderPath = exportRoot & "\vba\classes"
                fileExt = ".cls"

            Case 3      ' Form
                folderPath = exportRoot & "\vba\forms"
                fileExt = ".cls"

            Case 100    ' Document module / Report
                If Left(vbComp.Name, 6) = "Report" Then
                    folderPath = exportRoot & "\vba\reports"
                ElseIf Left(vbComp.Name, 4) = "Form" Then
                    folderPath = exportRoot & "\vba\forms"
                Else
                    folderPath = exportRoot & "\vba\classes"
                End If
                fileExt = ".cls"

            Case Else
                folderPath = exportRoot & "\vba\classes"
                fileExt = ".cls"

        End Select

        outPath = folderPath & "\" & CleanFileName(vbComp.Name) & fileExt

        If FileExists(outPath) Then
            Kill outPath
        End If

        vbComp.Export outPath

    Next vbComp

End Sub

Private Sub ExportSavedQueries(ByVal exportRoot As String)

    Dim qdf As DAO.QueryDef
    Dim outPath As String
    Dim fileNum As Integer

    For Each qdf In CurrentDb.QueryDefs

        If ShouldExportQuery(qdf.Name) Then

            outPath = exportRoot & "\sql\" & CleanFileName(qdf.Name) & ".sql"

            fileNum = FreeFile

            Open outPath For Output As #fileNum
            Print #fileNum, qdf.sql
            Close #fileNum

        End If

    Next qdf

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
