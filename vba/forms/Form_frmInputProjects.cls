VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmInputProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFindProject_Click()
    Dim s As String
    Dim rs As DAO.Recordset

    s = Trim(Nz(Me.txtFindProject, ""))

    MsgBox "Search text = [" & s & "]"

    If s = "" Then
        MsgBox "Enter a project ID, project name, DBNUM, or fund to search for."
        Exit Sub
    End If

    Set rs = Me.RecordsetClone

    MsgBox "Trying project_id"
    rs.FindFirst "[project_id] Like '*" & Replace(s, "'", "''") & "*'"

    If rs.NoMatch Then
        MsgBox "No match in project_id, trying project_name"
        rs.FindFirst "[project_name] Like '*" & Replace(s, "'", "''") & "*'"
    End If

    If rs.NoMatch Then
        MsgBox "No match in project_name, trying DBNUM"
        rs.FindFirst "[DBNUM] Like '*" & Replace(s, "'", "''") & "*'"
    End If

    If rs.NoMatch Then
        MsgBox "No match in DBNUM, trying fund"
        rs.FindFirst "[fund] Like '*" & Replace(s, "'", "''") & "*'"
    End If

    If rs.NoMatch Then
        MsgBox "No matching project was found."
    Else
        Me.Bookmark = rs.Bookmark
        MsgBox "Match found."
    End If

    rs.Close
    Set rs = Nothing
End Sub
