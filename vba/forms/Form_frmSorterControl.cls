VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSorterControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrepareTables_Click()
    PrepareSorterTables
    Me.frmDebugLogSub.Form.Requery
End Sub

Private Sub cmdRunSorter_Click()
    RunSorter
    Me.frmDebugLogSub.Form.Requery
End Sub
Public Sub RefreshDebugWindow()
    On Error Resume Next

    Me.frmDebugLogSub.Requery
    DoEvents

    With Me.frmDebugLogSub.Form.RecordsetClone
        If .RecordCount > 0 Then
            .MoveLast
            Me.frmDebugLogSub.Form.Bookmark = .Bookmark
        End If
    End With

    Me.Repaint
    DoEvents

    PauseSeconds 0.15
End Sub
