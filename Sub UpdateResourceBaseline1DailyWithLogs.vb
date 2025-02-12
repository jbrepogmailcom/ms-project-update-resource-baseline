Sub UpdateResourceBaseline1DailyWithLogs240731h()
    Dim proj As Project
    Dim res As Resource
    Dim assgn As Assignment
    Dim tsdWork As TimeScaleValues
    Dim tsvWork As TimeScaleValue
    Dim startDate As Date
    Dim endDate As Date
    Dim totalBaseline1Work As Double
    Dim periodStart As Date
    Dim periodEnd As Date
    Dim taskExists As Boolean
    Dim assignmentsToDelete As Collection
    Dim item As Variant
    Dim logFilePath As String
    Dim logFile As Object
    Dim logEntry As String
    Dim taskWBS As String
    Dim tsdResWork As TimeScaleValues

    Set proj = ActiveProject
    startDate = proj.ProjectStart
    endDate = proj.ProjectFinish

    ' Initialize collection to hold assignments to delete
    Set assignmentsToDelete = New Collection

    ' Set log file path
    logFilePath = "C:\Temp\Baseline1WorkLog.txt" ' Change this path as needed

    ' Initialize log file
    Set logFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(logFilePath, 2, True)
    logFile.WriteLine "WBS,Resource,Date,Value"

    ' Check each resource and its assignments
    For Each res In proj.Resources
        If Not res Is Nothing Then
            For Each assgn In res.Assignments
                ' Check if the related task still exists
                On Error Resume Next
                taskExists = Not proj.Tasks(assgn.taskID) Is Nothing
                On Error GoTo 0

                If Not taskExists Then
                    ' Add the assignment to the collection for deletion
                    assignmentsToDelete.Add assgn
                End If
            Next assgn
        End If
    Next res

    ' Delete all assignments to deleted tasks
    For Each item In assignmentsToDelete
        item.Delete
    Next item

    ' Update Baseline1 Work summary
    For Each res In proj.Resources
        If Not res Is Nothing Then
            ' Initialize period start to the project start date
            periodStart = startDate

            ' Initialize total baseline work for the resource
            Dim totalResourceBaseline1Work As Double
            totalResourceBaseline1Work = 0

            Do While periodStart <= endDate
                ' Set the end of the current period to the end of the day
                periodEnd = periodStart

                ' Reset the total baseline work for the current period
                totalBaseline1Work = 0

                ' Process each assignment for the resource
                For Each assgn In res.Assignments
                    ' Check if the assignment is not to a summary task
                    If Not proj.Tasks(assgn.taskID).Summary Then
                        ' Get the daily time-phased data for Baseline1 Work
                        Set tsdWork = assgn.TimeScaleData(periodStart, periodEnd, pjAssignmentTimescaledBaseline1Work, pjTimescaleDays)

                        For Each tsvWork In tsdWork
                            If IsNumeric(tsvWork.value) Then
                                totalBaseline1Work = totalBaseline1Work + CDbl(tsvWork.value)
                                totalResourceBaseline1Work = totalResourceBaseline1Work + CDbl(tsvWork.value)

                                ' Get the WBS value safely
                                On Error Resume Next
                                taskWBS = proj.Tasks(assgn.taskID).WBS
                                On Error GoTo 0

                                ' Ensure WBS is not null
                                If taskWBS = "" Then taskWBS = "N/A"

                                ' Log the details
                                logEntry = taskWBS & "," & res.Name & "," & Format(periodStart, "yyyy-mm-dd") & "," & tsvWork.value
                                logFile.WriteLine logEntry
                            End If
                        Next tsvWork
                    End If
                Next assgn

                ' Zero out the existing time-phased baseline work for the resource summary for this day
                Set tsdResWork = res.TimeScaleData(periodStart, periodEnd, pjResourceTimescaledBaseline1Work, pjTimescaleDays)
                If tsdResWork.Count > 0 Then
                    tsdResWork(1).value = totalBaseline1Work
                Else
                    ' Add a new entry if it doesn't exist
                    Set tsvWork = tsdResWork.Add(periodStart, periodEnd)
                    tsvWork.value = totalBaseline1Work
                End If

                ' Move to the next period (next day)
                periodStart = DateAdd("d", 1, periodStart)
            Loop

            ' Update the total Baseline1 Work for the resource
            res.Baseline1Work = totalResourceBaseline1Work
        End If
    Next res

    ' Close the log file
    logFile.Close

    MsgBox "Resource Baseline1 Work updated for daily time-phased data. Log created at " & logFilePath
End Sub

