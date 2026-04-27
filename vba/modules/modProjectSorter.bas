Attribute VB_Name = "modProjectSorter"
Option Compare Database
Option Explicit

Private Const LOG_TABLE As String = "DebugLog"
Private Const WORK_PROJECTS As String = "WorkProjects"
Private Const WORK_FUNDS As String = "WorkFunds"
Private Const OUTPUT_AWARDS As String = "OutputAwardLog"

Public Sub PrepareSorterTables()

    On Error GoTo ErrHandler

    ' Clear working and output tables
    CurrentDb.Execute "qryReset_WorkProjects", dbFailOnError
    CurrentDb.Execute "qryReset_WorkFunds", dbFailOnError
    CurrentDb.Execute "qryReset_OutputAwards", dbFailOnError
    CurrentDb.Execute "qryReset_InvalidProjects", dbFailOnError
    CurrentDb.Execute "qryReset_DebugLog", dbFailOnError

    LogMessage "START", "PrepareSorterTables started"
    LogMessage "STEP", "Working tables cleared"

    ' Append new working data
    CurrentDb.Execute "qryAppend_InvalidProjects", dbFailOnError
    CurrentDb.Execute "qryAppend_WorkProjects", dbFailOnError
    CurrentDb.Execute "qryAppend_WorkFunds", dbFailOnError

    LogMessage "STEP", "Working tables populated"

    LogMessage "END", "PrepareSorterTables completed"

    Exit Sub

ErrHandler:
    LogMessage "ERROR", "PrepareSorterTables failed: " & Err.Description

End Sub


Public Sub RunSorter()
    On Error GoTo ErrHandler

    LogMessage "START", "RunSorter started"

    ResetAwards
    ProcessYears 2026, 2050

    LogMessage "END", "RunSorter completed successfully"
    MsgBox "Sorter completed successfully.", vbInformation
    Exit Sub

ErrHandler:
    LogMessage "ERROR", "RunSorter failed: " & Err.Number & " - " & Err.Description
    MsgBox "RunSorter failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Function GetAwardCost(ByVal baseCost As Double, ByVal DBNUM As String, ByVal awardYear As Long) As Double
    Dim yrs As Long

    If Nz(DBNUM, "") Like "S*" Then
        GetAwardCost = baseCost
    Else
        yrs = awardYear - 2024
        If yrs < 0 Then yrs = 0
        GetAwardCost = Round(baseCost * (1.03 ^ yrs), 3)
    End If
End Function

Private Sub ResetAwards()
    On Error GoTo ErrHandler

    LogMessage "STEP", "ResetAwards started"

    CurrentDb.Execute "DELETE FROM " & OUTPUT_AWARDS, dbFailOnError
    CurrentDb.Execute "UPDATE " & WORK_PROJECTS & " SET awarded = False, award_year = 0", dbFailOnError

    LogMessage "STEP", "ResetAwards completed"
    Exit Sub

ErrHandler:
    LogMessage "ERROR", "ResetAwards failed: " & Err.Description
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub ProcessYears(ByVal StartYear As Long, ByVal EndYear As Long)
    Dim yr As Long

    For yr = StartYear To EndYear
        LogMessage "YEAR", "Processing year " & yr

        ProcessTIPProjects yr
        ProcessYearAC yr
        ProcessYearFlexible yr
    Next yr
End Sub

Private Sub ProcessTIPProjects(ByVal yr As Long)
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim fundAmt As Double
    Dim newAmt As Double

    sql = "SELECT * FROM " & WORK_PROJECTS & _
          " WHERE awarded = False" & _
          " AND year_eligible = " & yr & _
          " AND DBNUM Like 'S*'" & _
          " ORDER BY project_id ASC"

    Set rs = CurrentDb.OpenRecordset(sql, dbOpenDynaset)

    Do While Not rs.EOF
        fundAmt = Nz(GetFundAmount(rs!fund, yr), 0)
        newAmt = fundAmt - Nz(rs!cost, 0)

        If newAmt < 0 Then newAmt = 0

        CurrentDb.Execute _
            "UPDATE " & WORK_PROJECTS & _
            " SET awarded = True, award_year = " & yr & _
            " WHERE project_id = '" & Replace(rs!project_id, "'", "''") & "'", _
            dbFailOnError

        UpdateFundAmount rs!fund, yr, newAmt

        CurrentDb.Execute _
            "INSERT INTO " & OUTPUT_AWARDS & _
            " (award_year, award_round, fund, project_id, project_name, project_cost, fund_used, pool_used, residual_to_pool) VALUES (" & _
            yr & ", " & _
            "'TIP Projects are automatically funded without regard to estimated funding in the RTP', " & _
            "'" & Replace(rs!fund, "'", "''") & "', " & _
            "'" & Replace(rs!project_id, "'", "''") & "', " & _
            "'" & Replace(rs!project_name, "'", "''") & "', " & _
            Nz(rs!cost, 0) & ", " & _
            Nz(rs!cost, 0) & ", " & _
            "0, " & _
            "0)", _
            dbFailOnError

        LogMessage "TIP", "TIP project deducted: " & rs!project_id & _
                          ", year=" & yr & _
                          ", fund=" & rs!fund & _
                          ", cost=" & Format(Nz(rs!cost, 0), "0.000") & _
                          ", remaining fund=" & Format(newAmt, "0.000")

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Sub ProcessYearAC(ByVal yr As Long)
    Dim available As Double
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim projCost As Double
    Dim awardedThisPass As Boolean

    available = Nz(GetFundAmount("STBGP-AC", yr), 0)

    Do While available > 0.1
        awardedThisPass = False

        sql = "SELECT * FROM " & WORK_PROJECTS & _
              " WHERE awarded = False" & _
              " AND (DBNUM Not Like 'S*' OR DBNUM Is Null)" & _
              " AND fund = 'STBGP-AC'" & _
              " AND year_eligible <= " & yr & _
              " ORDER BY score DESC, year_eligible ASC, cost ASC, project_id ASC"

        Set rs = CurrentDb.OpenRecordset(sql, dbOpenDynaset)

        Do While Not rs.EOF
            projCost = GetAwardCost(Nz(rs!cost, 0), Nz(rs!DBNUM, ""), yr)

            If projCost <= available Then
                AwardProject rs!project_id, yr, rs!fund, projCost, rs!project_name, Nz(rs!DBNUM, ""), _
                             "AC", projCost, 0, 0

                available = available - projCost
                UpdateFundAmount "STBGP-AC", yr, available

                LogMessage "AWARD", "AC award: " & rs!project_id & _
                                   ", year=" & yr & _
                                   ", inflated cost=" & Format(projCost, "0.000") & _
                                   ", remaining AC fund=" & Format(available, "0.000")

                awardedThisPass = True
                Exit Do
            Else
                LogMessage "SKIP", "AC project skipped: " & rs!project_id & _
                                   ", year=" & yr & _
                                   ", inflated cost=" & Format(projCost, "0.000") & _
                                   ", available=" & Format(available, "0.000")
            End If

            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing

        If Not awardedThisPass Then Exit Do
    Loop

    LogMessage "INFO", "AC processing complete for year " & yr & _
                       "; remaining AC fund=" & Format(available, "0.000")
End Sub

Private Sub ProcessYearFlexible(ByVal yr As Long)
    Dim pool As Double
    Dim changed As Boolean
    Dim roundNum As Long

    roundNum = 1
    pool = SeedPoolForYear(yr)

    Do
        changed = False
        LogMessage "ROUND", "Year " & yr & ", round " & roundNum & ", starting pool=" & Format(pool, "0.000")

        changed = EvaluateFlexibleRound(yr, pool, roundNum)

        If changed Then roundNum = roundNum + 1
    Loop While changed = True

    CleanupPoolAwards yr, pool
End Sub

Private Function SeedPoolForYear(ByVal yr As Long) As Double
    Dim total As Double

    total = 0
    total = total + GetResidualIfActivated("STBGP-B50K200K", yr)
    total = total + GetResidualIfActivated("STBGP-B5K50K", yr)
    total = total + GetResidualIfActivated("STBGP-L5K", yr)

    LogMessage "POOL", "Initial pool for year " & yr & " = " & Format(total, "0.000")
    SeedPoolForYear = total
End Function

Private Function EvaluateFlexibleRound(ByVal yr As Long, ByRef pool As Double, ByVal roundNum As Long) As Boolean
    Dim funds(1 To 3) As String
    Dim i As Integer
    Dim awardedThisRound As Boolean

    funds(1) = "STBGP-B50K200K"
    funds(2) = "STBGP-B5K50K"
    funds(3) = "STBGP-L5K"

    awardedThisRound = False

    For i = 1 To 3
        If TryAwardTopFlexible(funds(i), yr, pool, roundNum) Then
            awardedThisRound = True
        End If
    Next i

    EvaluateFlexibleRound = awardedThisRound
End Function

Private Function TryAwardTopFlexible(ByVal fundName As String, ByVal yr As Long, ByRef pool As Double, ByVal roundNum As Long) As Boolean
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim fundAmt As Double
    Dim totalAvailable As Double
    Dim fundUsed As Double
    Dim poolUsed As Double
    Dim residualToPool As Double
    Dim projCost As Double

    sql = "SELECT * FROM " & WORK_PROJECTS & _
          " WHERE awarded = False" & _
          " AND (DBNUM Not Like 'S*' OR DBNUM Is Null)" & _
          " AND fund = '" & Replace(fundName, "'", "''") & "'" & _
          " AND year_eligible <= " & yr & _
          " ORDER BY score DESC, year_eligible ASC, cost ASC, project_id ASC"

    Set rs = CurrentDb.OpenRecordset(sql, dbOpenDynaset)

    fundAmt = Nz(GetFundAmount(fundName, yr), 0)

    Do While Not rs.EOF
        projCost = GetAwardCost(Nz(rs!cost, 0), Nz(rs!DBNUM, ""), yr)
        totalAvailable = fundAmt + pool

        If projCost <= totalAvailable Then
            If projCost <= fundAmt Then
                fundUsed = projCost
                poolUsed = 0
                residualToPool = fundAmt - projCost

                UpdateFundAmount fundName, yr, 0
                pool = pool + residualToPool
            Else
                fundUsed = fundAmt
                poolUsed = projCost - fundAmt
                residualToPool = 0

                UpdateFundAmount fundName, yr, 0
                pool = pool - poolUsed
            End If

            AwardProject rs!project_id, yr, rs!fund, projCost, rs!project_name, Nz(rs!DBNUM, ""), _
                         "ROUND " & roundNum, fundUsed, poolUsed, residualToPool

            LogMessage "AWARD", "Flexible award: " & rs!project_id & _
                                   ", year=" & yr & _
                                   ", fund=" & fundName & _
                                   ", fund_used=" & Format(fundUsed, "0.000") & _
                                   ", pool_used=" & Format(poolUsed, "0.000") & _
                                   ", residual_to_pool=" & Format(residualToPool, "0.000") & _
                                   ", remaining pool=" & Format(pool, "0.000")

            TryAwardTopFlexible = True
            rs.Close
            Set rs = Nothing
            Exit Function
        Else
            LogMessage "SKIP", "Flexible project skipped: " & rs!project_id & _
                               ", year=" & yr & _
                               ", fund=" & fundName & _
                               ", cost=" & Format(projCost, "0.000") & _
                               ", available=" & Format(totalAvailable, "0.000")
        End If

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    TryAwardTopFlexible = False
End Function
Private Sub CleanupPoolAwards(ByVal yr As Long, ByRef pool As Double)
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim awardCost As Double

    Do While pool > 0.1

        sql = "SELECT TOP 1 * FROM " & WORK_PROJECTS & _
              " WHERE awarded = False" & _
              " AND (DBNUM Not Like 'S*' OR DBNUM Is Null)" & _
              " AND fund In ('STBGP-B50K200K','STBGP-B5K50K','STBGP-L5K')" & _
              " AND year_eligible <= " & yr & _
              " ORDER BY score DESC, year_eligible ASC, cost ASC, project_id ASC"

        Set rs = CurrentDb.OpenRecordset(sql, dbOpenDynaset)

        If rs.EOF Then
            rs.Close
            Set rs = Nothing
            Exit Do
        End If

        awardCost = GetAwardCost(Nz(rs!cost, 0), Nz(rs!DBNUM, ""), yr)

        LogMessage "CLEANUP", "Cleanup check: " & rs!project_id & _
                              ", year=" & yr & _
                              ", base cost=" & Format(Nz(rs!cost, 0), "0.000") & _
                              ", award cost=" & Format(awardCost, "0.000") & _
                              ", pool=" & Format(pool, "0.000")

        If awardCost <= pool Then
            AwardProject rs!project_id, yr, rs!fund, awardCost, rs!project_name, Nz(rs!DBNUM, ""), _
                         "CLEANUP", 0, awardCost, 0

            pool = pool - awardCost

            LogMessage "CLEANUP", "Cleanup pool award: " & rs!project_id & _
                                  ", year=" & yr & _
                                  ", award cost=" & Format(awardCost, "0.000") & _
                                  ", pool now=" & Format(pool, "0.000")
        Else
            LogMessage "CLEANUP", "Cleanup stopped: top candidate " & rs!project_id & _
                                  " costs " & Format(awardCost, "0.000") & _
                                  " but remaining pool is " & Format(pool, "0.000")
            rs.Close
            Set rs = Nothing
            Exit Do
        End If

        rs.Close
        Set rs = Nothing
    Loop
End Sub

Private Sub AwardProject(ByVal projectID As String, ByVal yr As Long, ByVal fundName As String, _
                         ByVal amt As Double, ByVal projectName As String, ByVal DBNUM As String, _
                         ByVal awardRound As String, ByVal fundUsed As Double, ByVal poolUsed As Double, _
                         ByVal residualToPool As Double)

    Dim eligibleYear As Long

    eligibleYear = Nz(DLookup("year_eligible", WORK_PROJECTS, _
                    "project_id='" & Replace(projectID, "'", "''") & "'"), 0)

    If eligibleYear > yr Then
        LogMessage "ERROR", "Blocked early award: " & projectID & _
                            ", eligible=" & eligibleYear & _
                            ", attempted year=" & yr
        Err.Raise vbObjectError + 1000, , _
                  "Project " & projectID & " is not eligible until year " & eligibleYear
    End If

    CurrentDb.Execute _
        "UPDATE " & WORK_PROJECTS & _
        " SET awarded = True, award_year = " & yr & _
        " WHERE project_id = '" & Replace(projectID, "'", "''") & "'", _
        dbFailOnError

    CurrentDb.Execute _
        "INSERT INTO " & OUTPUT_AWARDS & _
        " (award_year, award_round, fund, project_id, project_name, project_cost, fund_used, pool_used, residual_to_pool) VALUES (" & _
        yr & ", " & _
        "'" & Replace(awardRound, "'", "''") & "', " & _
        "'" & Replace(fundName, "'", "''") & "', " & _
        "'" & Replace(projectID, "'", "''") & "', " & _
        "'" & Replace(projectName, "'", "''") & "', " & _
        amt & ", " & _
        fundUsed & ", " & _
        poolUsed & ", " & _
        residualToPool & ")", _
        dbFailOnError
End Sub

Private Function GetFundAmount(ByVal fundName As String, ByVal yr As Long) As Double
    GetFundAmount = Nz(DLookup("projected_amount", WORK_FUNDS, _
                    "fund='" & Replace(fundName, "'", "''") & "' AND [year]=" & yr), 0)
End Function

Private Sub UpdateFundAmount(ByVal fundName As String, ByVal yr As Long, ByVal newAmt As Double)
    If newAmt < 0 Then newAmt = 0

    CurrentDb.Execute _
        "UPDATE " & WORK_FUNDS & _
        " SET projected_amount = " & Replace(Format(newAmt, "0.000"), ",", ".") & _
        " WHERE fund='" & Replace(fundName, "'", "''") & "'" & _
        " AND [year]=" & yr, _
        dbFailOnError
End Sub

Private Function GetResidualIfActivated(ByVal fundName As String, ByVal yr As Long) As Double
    Dim amt As Double
    amt = Nz(GetFundAmount(fundName, yr), 0)

    If HadTIPOrAwardActivity(fundName, yr) Then
        GetResidualIfActivated = amt
    Else
        GetResidualIfActivated = 0
    End If
End Function

Private Function HadTIPOrAwardActivity(ByVal fundName As String, ByVal yr As Long) As Boolean
    Dim n As Long

    n = Nz(DCount("*", WORK_PROJECTS, _
        "fund='" & Replace(fundName, "'", "''") & "'" & _
        " AND year_eligible=" & yr & _
        " AND DBNUM Like 'S*'"), 0)

    HadTIPOrAwardActivity = (n > 0)
End Function


Public Sub LogMessage(ByVal logType As String, ByVal msg As String)

    CurrentDb.Execute _
        "INSERT INTO DebugLog (log_time, log_type, log_message) VALUES (" & _
        "Now(), '" & Replace(logType, "'", "''") & "', '" & Replace(msg, "'", "''") & "')"

    On Error Resume Next
    Forms!frmSorterControl.RefreshDebugWindow
    DoEvents

End Sub
Public Sub PauseSeconds(ByVal seconds As Double)
    Dim endTime As Double
    endTime = Timer + seconds

    Do While Timer < endTime
        DoEvents
    Loop
End Sub

Private Sub UpdateButtonStates()

    Dim hasWork As Boolean
    Dim hasInvalid As Boolean
    Dim hasAwards As Boolean

    hasWork = (DCount("*", "WorkProjects") > 0)
    hasInvalid = (DCount("*", "InvalidProjects") > 0)
    hasAwards = (DCount("*", "WorkProjects", "awarded=True") > 0)

    ' Always available
    Me.cmdProjectEditor.Enabled = True
    Me.Command7.Enabled = True ' Edit Funds
    Me.cmdPrepareTables.Enabled = True

    ' After PrepareTables
    Me.cmdInvalidProjectsFormOpen.Enabled = hasWork
    Me.cmdRunSorter.Enabled = hasWork

    ' After RunSorter
    Me.[Examine RTP Table].Enabled = hasAwards
    Me.[Preview Table 27].Enabled = hasAwards
    Me.Command16.Enabled = hasAwards

End Sub
