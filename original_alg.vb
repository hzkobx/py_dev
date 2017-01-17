'-----------------------------
' Input Constants Definitions
'-----------------------------

' Following is a set of X,Y coordinates where inputs are taken (from "Inputs" Worksheet)

'X coordinate (Column) where most of inputs are located
Const InputsStartC As Integer = 2

' Time Window (ns)
Public Const IN_timeWindow_C As Integer = InputsStartC
Public Const IN_timeWindow_R As Integer = 4
' Idle Task Id
Public Const IN_idleTaskId_C As Integer = InputsStartC
Public Const IN_idleTaskId_R As Integer = 5
' No Interrupt Id
Public Const IN_NoIntId_C As Integer = InputsStartC
Public Const IN_NoIntId_R As Integer = 6
' Number of ISR Direct Functions
Public Const IN_NbDirectISRs_C As Integer = InputsStartC
Public Const IN_NbDirectISRs_R As Integer = 7
' Direct ISR Sheet Name(s)
Public Const IN_DirectISRsName_C As Integer = InputsStartC
Public Const IN_DirectISRsName_R As Integer = 8
' No. of Background Tasks
Public Const IN_NbBkgTasks_C As Integer = InputsStartC
Public Const IN_NbBkgTasks_R As Integer = 9
' Background Tasks Ids
Public Const IN_BkgTasksId_C As Integer = InputsStartC
Public Const IN_BkgTasksId_R As Integer = 10
' Tasks Data File Name
Public Const IN_TaskDataFile_C As Integer = InputsStartC
Public Const IN_TaskDataFile_R As Integer = 11
' ISRs Data File Name
Public Const IN_IsrsDataFile_C As Integer = InputsStartC
Public Const IN_IsrsDataFile_R As Integer = 12
' Working Directory
Public Const IN_WorkingDir_C As Integer = InputsStartC
Public Const IN_WorkingDir_R As Integer = 13


'----------------------
' Global Constants
'----------------------
Const FuncNofVars As Integer = 7
Const FuncMaxNumberOf As Integer = 30
Const vTRUE As Integer = 1
Const vFALSE As Integer = 0
Const ResultsSlots As Integer = 10
Const ResultsColumns As Integer = 3 ' Time, CPU load with ISR, CPU load without ISR

'----------------------
' Variables for ISRs
'----------------------
Dim FuncNumberOf As Integer
Dim FuncCurrent As Integer
Dim FuncVariables(FuncMaxNumberOf, FuncNofVars) As Variant
Dim FuncNoneId As Variant
'Variables definitions for ISR functions
Const Fvar_StartTime As Integer = 0
Const Fvar_EndTime As Integer = 1
Const Fvar_CurrentTime As Integer = 2
Const Fvar_WindowAccTime As Integer = 3
Const Fvar_AccumulatedTime As Integer = 4
Const Fvar_Measuring As Integer = 5
Const Fvar_notNumber As Integer = 6
Const Fvar_CurrentRow As Integer = 7

Dim FuncAllAccTime As Variant
' Number of processed windows
Dim Executed As Variant
' Variables for percentage average
Dim AccPercentage As Variant
Dim AccPercentageNoISR As Variant
' Variables for percentage average
Dim AccISRsPercentage As Variant
' CPU load progress and cancel
Dim CPUloadProgressLastRow As Variant
Dim CPUloadCalcCancel As Integer
' Objects IDs
Dim IdleTaskId As Variant
Dim ISRnoneId As Variant
Const MaxBGtasks As Integer = 10
Dim BGtasksIds(MaxBGtasks) As Variant
Dim BGcurrentTasksNumber As Variant



'================================================================================
' Name: CPU_load_calculation
' -------------------------------------------------------------------------------
' Description: Main Sub to perform the CPU load calculation
' Called: Within mainApp form
' Pre-conditions: inputs already verified
' Post-conditions: CPU load already calculated, including max peak, average (both
'                  with and without ISRs
'================================================================================
Sub CPU_load_calculation()
    
    'Generic variables
    Dim ForCounter As Integer
    Dim NoCPUusageTask As Integer
    
    ' variables related to the current window of time being analyzed
    Dim TimeWindow As Variant
    Dim StartTime As Variant
    Dim CurrentTime As Variant
    Dim CurrentRow As Variant
    
    'IDLE task analysis
    Dim IdleTaskAcummulatedTime As Variant
    Dim IdleTaskStartTime As Variant
    Dim IdleTaskEndTime As Variant
    Dim IdleTaskMeasuring As Variant
    
    'Error variables
    Dim TaskTimeError As Integer
    
    'Results variables
    Dim ResultsArrayRowIndex As Integer
    Dim ResultsArrayColIndex As Integer
    Dim TempTime As Variant
    Dim ResultsSortAux As Integer
    Dim CPUloadWithoutISRs As Variant
    Dim CPUloadWithISRs As Variant
    Dim ResultsMatrix(ResultsColumns - 1, ResultsSlots - 1) As Variant
    Dim ISRsMax(ResultsColumns - 1, ResultsSlots - 1) As Variant
    
    'Graphs variables
    Dim GraphDataRowIdx As Integer
    Dim GraphCPUloadTemp As Double
    Dim GraphISRloadTemp As Double
    Dim GraphCPUloadWithISRsTemp As Double
    
    '---------------------------
    ' Initializations
    '---------------------------
    CPUloadCalcCancel = 0
    AccPercentage = 0
    AccPercentageNoISR = 0
    AccISRsPercentage = 0
    
    GraphDataRowIdx = 0
    
    'Init results matrix
    For ResultsArrayRowIndex = 0 To (ResultsSlots - 1) Step 1
        For ResultsArrayColIndex = 0 To (ResultsColumns - 1) Step 1
            ResultsMatrix(ResultsArrayColIndex, ResultsArrayRowIndex) = 0
        Next ResultsArrayColIndex
    Next ResultsArrayRowIndex
    ' Init ISR usage matrix
    For ResultsArrayRowIndex = 0 To (ResultsSlots - 1) Step 1
        For ResultsArrayColIndex = 0 To (ResultsColumns - 1) Step 1
            ISRsMax(ResultsArrayColIndex, ResultsArrayRowIndex) = 0
        Next ResultsArrayColIndex
    Next ResultsArrayRowIndex
    'Clear prevous results
    Worksheets("Results").Cells(3, 1).Value = ""
    Worksheets("Results").Cells(3, 2).Value = ""
    Worksheets("Results").Cells(3, 3).Value = ""
    Worksheets("Results").Cells(5, 2).Value = ""
    'Clear previous messages
    Worksheets("Tasks").Activate
    Range(Cells(2, 4), Cells(65500, 5)).ClearContents
    Worksheets("ISRs").Activate
    Range(Cells(2, 4), Cells(65500, 4)).ClearContents
    ' Init inputs
    ' Time is in nanoseconds so for a 1 ms window value must be 1000000 and so on
    TimeWindow = Worksheets("Inputs").Cells(IN_timeWindow_R, IN_timeWindow_C).Value
    IdleTaskId = Worksheets("Inputs").Cells(IN_idleTaskId_R, IN_idleTaskId_C).Value
    FuncNoneId = Worksheets("Inputs").Cells(IN_NoIntId_R, IN_NoIntId_C).Value
    ' Init variables
    TaskTimeError = 0
    CurrentRow = 2 ' Init Cell is 2 because data starts in row 2
    Executed = 0
    'Init number of ISR Direct Functions
    FuncNumberOf = Worksheets("Inputs").Cells(IN_NbDirectISRs_R, IN_NbDirectISRs_C).Value
    ' ISRs is always present so at least 1 ISR function
    FuncNumberOf = FuncNumberOf + 1
    If FuncNumberOf > 0 Then
        'Init direct interrupts
        For FuncCurrent = 0 To (FuncNumberOf) Step 1
            FuncVariables(FuncCurrent, Fvar_CurrentRow) = 2
        Next FuncCurrent
    End If
    'Init number of Background Tasks
    BGcurrentTasksNumber = Worksheets("Inputs").Cells(IN_NbBkgTasks_R, IN_NbBkgTasks_C).Value
    'Init Background tasks Ids
    If BGcurrentTasksNumber > 0 Then
        'Init direct interrupts
        For ForCounter = 0 To (BGcurrentTasksNumber) Step 1
            BGtasksIds(ForCounter) = Worksheets("Inputs").Cells(IN_BkgTasksId_R, IN_BkgTasksId_C + ForCounter).Value
        Next ForCounter
    End If
    
    ' ----------------------------------------------------------
    ' Calculate CPU load for the each Time Window
    ' ----------------------------------------------------------
    While ((Not (Worksheets("Tasks").Cells(CurrentRow, 1).Value = "")) And TaskTimeError = 0 And CPUloadCalcCancel = 0)
        
        
        ' Update CPU load progress label
        If (Timer > (Inicio + TiempoPausa)) Then
            'Timer has expired, reload it
            Inicio = Timer
            mainApp.Label2.Caption = "CPU Load Progress... Analyzed " + Format(CurrentRow) + " rows of " + Format(getLastRow)
            Application.StatusBar = Executed
            DoEvents
        Else
            ' Timer not expired
        End If
        
        ' Increment number of executed Windows of Time
        Executed = Executed + 1
        
        ' Get the starting time for current window
        StartTime = Worksheets("Tasks").Cells(CurrentRow, 1).Value
        ' At this point current time is the same than start time
        CurrentTime = StartTime
        ' Init variables for CPU load calculation if current Window
        IdleTaskMeasuring = 0
        IdleTaskAcummulatedTime = 0
        FuncAllAccTime = 0
        TaskTimeError = 0
        
        ' For the time of the Window, look for Idle task running time, if an error occurs, exit this window and continue with the
        ' next one
        While (CurrentTime <= (StartTime + TimeWindow)) And (TaskTimeError = 0) And (CPUloadCalcCancel = 0)
            ' Does current row belongs to a background task??
            NoCPUusageTask = 0 'FALSE
            If BGcurrentTasksNumber > 0 Then
                'Init direct interrupts
                For ForCounter = 0 To (BGcurrentTasksNumber) Step 1
                    If Worksheets("Tasks").Cells(CurrentRow, 2).Value = BGtasksIds(ForCounter) Then
                        NoCPUusageTask = 1 'TRUE
                        Exit For
                    End If
                Next ForCounter
            End If
            
            'Does current row belongs to an Idle task or background task??
            If Worksheets("Tasks").Cells(CurrentRow, 2).Value = IdleTaskId Or NoCPUusageTask Then
                If IdleTaskMeasuring = 0 Then
                    'New Idle task measurement
                    IdleTaskMeasuring = 1
                    IdleTaskStartTime = Worksheets("Tasks").Cells(CurrentRow, 1).Value
                Else
                    'do nothing, Idle task detected in a previous slow and not ended yet
                End If
            Else
                If IdleTaskMeasuring = 0 Then
                    'do nothing
                Else
                    'End Idle task measurement
                    IdleTaskMeasuring = 0
                    IdleTaskEndTime = Worksheets("Tasks").Cells(CurrentRow, 1).Value
                    'Calculate Idle task total time (within the time window)
                    IdleTaskAcummulatedTime = IdleTaskAcummulatedTime + (IdleTaskEndTime - IdleTaskStartTime)
                    
                    ' -----------------------------------------------------------------------------
                    ' Logic to substract CPU time used by ISRs during IDLE task execution
                    ' -----------------------------------------------------------------------------
                    FuncAllAccTime = FuncAllAccTime + ProcessISRfunctions(0, IdleTaskStartTime, IdleTaskEndTime, "ISRs")
                    If FuncNumberOf > 1 Then
                        'Process direct interrupts
                        For FuncCurrent = 1 To (FuncNumberOf - 1) Step 1
                            FuncAllAccTime = FuncAllAccTime + ProcessISRfunctions(1 + FuncCurrent, IdleTaskStartTime, IdleTaskEndTime, Worksheets("Inputs").Cells(IN_DirectISRsName_R, ((IN_DirectISRsName_C - 1) + FuncCurrent)).Value)
                        Next FuncCurrent
                    End If
                    ' -----------------------------------------------------------------------------
                    ' End of Logic to substract CPU time used by ISRs during IDLE task execution
                    ' -----------------------------------------------------------------------------
                End If
            End If
            ' Increment Task row, to continue looking for Idle task
            CurrentRow = CurrentRow + 1
            ' Is next row to be analyzed empty?
            If Worksheets("Tasks").Cells(CurrentRow, 1).Value <> "" Then
                ' Not empty, Ok
                CurrentTime = Worksheets("Tasks").Cells(CurrentRow, 1).Value
            Else
                ' Empty, and error is flagged
                TaskTimeError = 1
            End If
            ' DoEvents allow Windows OS to execute queued events, for example request to cancel this Script
            DoEvents
        Wend
        
        'Window has been partially analyzed, determine if an Idle task measurement where incomplete at the time
        ' the window ended
        If IdleTaskMeasuring And TaskTimeError = 0 Then
            ' Idle task truncated, force the end of it equal to the end of the window
            IdleTaskAcummulatedTime = IdleTaskAcummulatedTime + ((StartTime + TimeWindow) - IdleTaskStartTime)
            'End Task IDLE measurement has been truncated, force ISR within IDLE task to be calculated
            IdleTaskMeasuring = 0
            IdleTaskEndTime = (StartTime + TimeWindow)

            ' -----------------------------------------------------------------------------
            ' Logic to substract CPU time used by ISRs during IDLE task execution
            ' -----------------------------------------------------------------------------
            FuncAllAccTime = FuncAllAccTime + ProcessISRfunctions(0, IdleTaskStartTime, IdleTaskEndTime, "ISRs")
            If FuncNumberOf > 1 Then
                'Process direct interrupts
                For FuncCurrent = 1 To (FuncNumberOf - 1) Step 1
                    FuncAllAccTime = FuncAllAccTime + ProcessISRfunctions(1 + FuncCurrent, IdleTaskStartTime, IdleTaskEndTime, Worksheets("Inputs").Cells(IN_DirectISRsName_R, ((IN_DirectISRsName_C - 1) + FuncCurrent)).Value)
                Next FuncCurrent
            End If
            ' -----------------------------------------------------------------------------
            ' End of Logic to substract CPU time used by ISRs during IDLE task execution
            ' -----------------------------------------------------------------------------
        End If
        
        'WINDOW HAS BEEN PROCESSED COMPLETELY
        ' If no error occurred, process window results with respect to global results
        If TaskTimeError = 0 Then
            'Write Graph Data
            GraphDataRowIdx = GraphDataRowIdx + 1
            GraphCPUloadTemp = 100 - ((IdleTaskAcummulatedTime / TimeWindow) * 100)
            GraphCPUloadWithISRsTemp = 100 - (((IdleTaskAcummulatedTime / TimeWindow) - ((FuncAllAccTime) / TimeWindow)) * 100)
            GraphISRloadTemp = GraphCPUloadWithISRsTemp - GraphCPUloadTemp
            
            Worksheets("Graphs").Cells(GraphDataRowIdx + 1, 1).Value = GraphDataRowIdx
            Worksheets("Graphs").Cells(GraphDataRowIdx + 1, 2).Value = Worksheets("Tasks").Cells(CurrentRow, 1).Value
            Worksheets("Graphs").Cells(GraphDataRowIdx + 1, 3).Value = GraphCPUloadTemp
            Worksheets("Graphs").Cells(GraphDataRowIdx + 1, 4).Value = GraphCPUloadWithISRsTemp
            Worksheets("Graphs").Cells(GraphDataRowIdx + 1, 5).Value = GraphISRloadTemp
            
            ' Write window results into Tasks worksheet
            Worksheets("Tasks").Cells(CurrentRow, 4).Value = GraphCPUloadTemp
            Worksheets("Tasks").Cells(CurrentRow, 5).Value = GraphCPUloadWithISRsTemp
            ' Calculate percentage of ISR for this window
            Worksheets("Tasks").Cells(CurrentRow, 6).Value = GraphISRloadTemp
            ' Logic for CPU load average
            AccPercentage = AccPercentage + (100 - (((IdleTaskAcummulatedTime / TimeWindow) - ((FuncAllAccTime) / TimeWindow)) * 100))
            AccPercentageNoISR = AccPercentageNoISR + (100 - ((IdleTaskAcummulatedTime / TimeWindow) * 100))
            ' Logic for ISRs load average
            AccISRsPercentage = AccISRsPercentage + ((100 - (((IdleTaskAcummulatedTime / TimeWindow) - ((FuncAllAccTime) / TimeWindow)) * 100)) - (100 - ((IdleTaskAcummulatedTime / TimeWindow) * 100)))
            ' Logic for CPU load biggest results
            ' See if a new CPU load must be pushed into the results matrix
            For ResultsSortAux = 0 To (ResultsSlots - 1) Step 1
                TempTime = 100 - (((IdleTaskAcummulatedTime / TimeWindow) - ((FuncAllAccTime) / TimeWindow)) * 100)
                If TempTime >= ResultsMatrix(1, ResultsSortAux) Then
                    ' Shift previous values
                    For ResultsArrayRowIndex = (ResultsSlots - 2) To ResultsSortAux Step -1
                        For ResultsArrayColIndex = 0 To (ResultsColumns - 1) Step 1
                            ResultsMatrix(ResultsArrayColIndex, ResultsArrayRowIndex + 1) = ResultsMatrix(ResultsArrayColIndex, ResultsArrayRowIndex)
                        Next ResultsArrayColIndex
                    Next ResultsArrayRowIndex
                    ' Put new values
                    ResultsMatrix(0, ResultsSortAux) = StartTime
                    ResultsMatrix(1, ResultsSortAux) = 100 - (((IdleTaskAcummulatedTime / TimeWindow) - ((FuncAllAccTime) / TimeWindow)) * 100)
                    ResultsMatrix(2, ResultsSortAux) = 100 - ((IdleTaskAcummulatedTime / TimeWindow) * 100)
                    Exit For
                End If
            Next ResultsSortAux
            ' Logic to push ISR max usage
            For ResultsSortAux = 0 To (ResultsSlots - 1) Step 1
                ' TempTime = CPU load with ISRs - CPU load without ISRs
                TempTime = ((100 - (((IdleTaskAcummulatedTime / TimeWindow) - ((FuncAllAccTime) / TimeWindow)) * 100)) - (100 - ((IdleTaskAcummulatedTime / TimeWindow) * 100)))
                If TempTime >= ISRsMax(1, ResultsSortAux) Then
                    ' Shift previous values
                    For ResultsArrayRowIndex = (ResultsSlots - 2) To ResultsSortAux Step -1
                        For ResultsArrayColIndex = 0 To (ResultsColumns - 1) Step 1
                            ISRsMax(ResultsArrayColIndex, ResultsArrayRowIndex + 1) = ISRsMax(ResultsArrayColIndex, ResultsArrayRowIndex)
                        Next ResultsArrayColIndex
                    Next ResultsArrayRowIndex
                    ' Put new values
                    ISRsMax(0, ResultsSortAux) = StartTime
                    ISRsMax(1, ResultsSortAux) = TempTime
                    ISRsMax(2, ResultsSortAux) = 0
                    Exit For
                End If
            Next ResultsSortAux
        End If
        ' End Calculation and results of the current Time Window, continue looking..
        CurrentRow = CurrentRow + 1
        ' Update Progress bar
        Call updateProgressBar(CurrentRow)
        ' DoEvents allow Windows OS to execute queued events, for example request to cancel this Script
        DoEvents
    Wend
    ' ----------------------------------------------------------
    ' End of Calculate CPU load for the each Time Window
    ' ----------------------------------------------------------
    
    ' ----------------------------------------------------------
    ' Process Final results
    ' ----------------------------------------------------------
    'Clear prevous results
    Worksheets("Results").Cells(3, 1).Value = ""
    Worksheets("Results").Cells(3, 2).Value = ""
    Worksheets("Results").Cells(3, 3).Value = ""
    Worksheets("Results").Cells(5, 2).Value = ""
    Worksheets("Results").Activate
    Range(Cells(10, 1), Cells(19, 3)).ClearContents
    Range(Cells(24, 1), Cells(33, 3)).ClearContents
    Worksheets("Results").Cells(35, 2).Value = ""
    
    'Print results
    ' Max CPU load
    Worksheets("Results").Cells(3, 1).Value = ResultsMatrix(0, 0)
    Worksheets("Results").Cells(3, 2).Value = ResultsMatrix(1, 0)
    Worksheets("Results").Cells(3, 3).Value = ResultsMatrix(2, 0)
    'Average CPU load
    AccPercentage = (AccPercentage / Executed)
    Worksheets("Results").Cells(6, 2).Value = AccPercentage
    AccPercentageNoISR = (AccPercentageNoISR / Executed)
    Worksheets("Results").Cells(6, 3).Value = AccPercentageNoISR
    'Average ISRs load
    AccISRsPercentage = (AccISRsPercentage / Executed)
    Worksheets("Results").Cells(35, 2).Value = AccISRsPercentage
    ' Used Time Window
    Worksheets("Results").Cells(5, 2).Value = (TimeWindow / 1000000)
    ' Biggest CPU load matrix
    For ResultsArrayRowIndex = 0 To (ResultsSlots - 1) Step 1
        For ResultsArrayColIndex = 0 To (ResultsColumns - 1) Step 1
            Worksheets("Results").Cells(ResultsArrayRowIndex + 10, ResultsArrayColIndex + 1).Value = ResultsMatrix(ResultsArrayColIndex, ResultsArrayRowIndex)
        Next ResultsArrayColIndex
    Next ResultsArrayRowIndex
    ' Biggest ISR usage matrix
    For ResultsArrayRowIndex = 0 To (ResultsSlots - 1) Step 1
        For ResultsArrayColIndex = 0 To (ResultsColumns - 2) Step 1
            Worksheets("Results").Cells(ResultsArrayRowIndex + 24, ResultsArrayColIndex + 1).Value = ISRsMax(ResultsArrayColIndex, ResultsArrayRowIndex)
        Next ResultsArrayColIndex
    Next ResultsArrayRowIndex
    
        ' Update Graph
    GraphUpdate (GraphDataRowIdx)
    
    If CPUloadCalcCancel = 1 Then
        MsgBox "Operation Cancelled by user", vbInformation, "CPU load calculation"
    End If
End Sub

'================================================================================
' Name: ProcessISRfunctions
' -------------------------------------------------------------------------------
' Description: Function to process ISRs time while an Idle task execution
' Called: From  CPU_load_calculation
' Pre-conditions: inputs already verified
' Post-conditions: Total ISR time within an idle task execution is estimated
'================================================================================
Function ProcessISRfunctions(pFunctionId, pTaskIdleStartTime, pTaskIdleEndTime, pFuncName)


FuncVariables(pFunctionId, Fvar_EndTime) = 0 'Error = 0
FuncVariables(pFunctionId, Fvar_WindowAccTime) = 0 'Error = 0

    ' ===================================================================
    ' Logic to substract CPU time used by ISR during IDLE task execution
    ' ===================================================================
    FuncVariables(pFunctionId, Fvar_StartTime) = pTaskIdleStartTime
    FuncVariables(pFunctionId, Fvar_AccumulatedTime) = 0
    FuncVariables(pFunctionId, Fvar_Measuring) = 0 'FALSE
    FuncVariables(pFunctionId, Fvar_CurrentTime) = 0
    FuncVariables(pFunctionId, Fvar_notNumber) = 0 'Error = 0
    If FuncVariables(pFunctionId, Fvar_CurrentRow) > 6 Then
        FuncVariables(pFunctionId, Fvar_CurrentRow) = FuncVariables(pFunctionId, Fvar_CurrentRow) - 4
    End If
    
    ' -------------------------------------------------------------
    ' Logic to Look for the start point (firs useful row)
    ' -------------------------------------------------------------
    Do While FuncVariables(pFunctionId, Fvar_notNumber) = 0 And (CPUloadCalcCancel = 0)  ' Add secure limit
        If Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 1).Value <> "" Then
            If Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 1).Value >= pTaskIdleStartTime Then
                ' Suspicios ISR within IDLE task
                Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 4).Value = "Starting.."
                If Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 1).Value >= pTaskIdleEndTime Then
                    'ISR found is outside of IDLE task
                    FuncVariables(pFunctionId, Fvar_Measuring) = 0 'FALSE
                Else
                    ' An ISR execution has been found during IDLE task execution
                    FuncVariables(pFunctionId, Fvar_Measuring) = 1 'TRUE
                    Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 4).Value = "Starting Real.."
                    ' Set ISR start Row and start time
                    FuncVariables(pFunctionId, Fvar_StartTime) = Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 1).Value
                    FuncVariables(pFunctionId, Fvar_CurrentTime) = FuncVariables(pFunctionId, Fvar_StartTime)
                End If
                ' Exit Loop
                Exit Do
            Else
                ' Still looking for a possible ISR within IDLE task
                FuncVariables(pFunctionId, Fvar_CurrentRow) = FuncVariables(pFunctionId, Fvar_CurrentRow) + 1
            End If
            FuncVariables(pFunctionId, Fvar_notNumber) = 0
        Else
            Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 4).Value = "Error, row has not a number"
            FuncVariables(pFunctionId, Fvar_notNumber) = 1
        End If
        DoEvents
    Loop
    ' -------------------------------------------------------------
    ' End of Logic to Look for the start point (firs useful row)
    ' -------------------------------------------------------------
    
    ' ----------------------------------------------------------------------------
    ' Logic to estimate ISR accumulated time for curren Idle task execution
    ' ----------------------------------------------------------------------------
    If FuncVariables(pFunctionId, Fvar_Measuring) And FuncVariables(pFunctionId, Fvar_notNumber) = 0 Then
        'An ISR has been found within IDLE task. Process used time
        While FuncVariables(pFunctionId, Fvar_CurrentTime) < pTaskIdleEndTime And FuncVariables(pFunctionId, Fvar_notNumber) = 0 And (CPUloadCalcCancel = 0)
            If Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 2).Value = FuncNoneId Then
                If FuncVariables(pFunctionId, Fvar_Measuring) = 1 Then
                    'ISR has ended
                    FuncVariables(pFunctionId, Fvar_EndTime) = Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 1).Value
                    FuncVariables(pFunctionId, Fvar_AccumulatedTime) = (FuncVariables(pFunctionId, Fvar_EndTime) - FuncVariables(pFunctionId, Fvar_StartTime))
                    FuncVariables(pFunctionId, Fvar_Measuring) = 0
                    FuncVariables(pFunctionId, Fvar_WindowAccTime) = FuncVariables(pFunctionId, Fvar_WindowAccTime) + FuncVariables(pFunctionId, Fvar_AccumulatedTime)
                    Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 4).Value = FuncVariables(pFunctionId, Fvar_AccumulatedTime)
                Else
                    ' do nothing
                End If
            Else
                If FuncVariables(pFunctionId, Fvar_Measuring) = 1 Then
                    ' do nothing
                Else
                    'new ISR detection
                    FuncVariables(pFunctionId, Fvar_StartTime) = Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 1).Value
                    FuncVariables(pFunctionId, Fvar_Measuring) = 1
                End If
            End If
            FuncVariables(pFunctionId, Fvar_CurrentRow) = FuncVariables(pFunctionId, Fvar_CurrentRow) + 1
            If Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 1).Value <> "" Then
                FuncVariables(pFunctionId, Fvar_CurrentTime) = Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 1).Value
            Else
                FuncVariables(pFunctionId, Fvar_notNumber) = 1
            End If
            DoEvents
        Wend
    Else
    End If
    ' ----------------------------------------------------------------------------
    ' End of Logic to estimate ISR accumulated time for curren Idle task execution
    ' ----------------------------------------------------------------------------
    
    If FuncVariables(pFunctionId, Fvar_Measuring) And FuncVariables(pFunctionId, Fvar_notNumber) = 0 Then
        FuncVariables(pFunctionId, Fvar_WindowAccTime) = FuncVariables(pFunctionId, Fvar_WindowAccTime) + ((pTaskIdleEndTime) - FuncVariables(pFunctionId, Fvar_StartTime))
    Else
        FuncVariables(pFunctionId, Fvar_notNumber) = 0
    End If
    
    'Process results
    If FuncVariables(pFunctionId, Fvar_WindowAccTime) > 0 Then
        Worksheets(pFuncName).Cells(FuncVariables(pFunctionId, Fvar_CurrentRow), 5).Value = FuncVariables(pFunctionId, Fvar_WindowAccTime)
        ProcessISRfunctions = FuncVariables(pFunctionId, Fvar_WindowAccTime)
    Else
        ProcessISRfunctions = 0
    End If
    
End Function

'================================================================================
' Name: GetProgress
' -------------------------------------------------------------------------------
' Description: Return number of processed windows
' Called: Can be called anywhere
' Pre-conditions: CPU load calculation running
' Post-conditions: --
'================================================================================
Function GetProgress()
    GetProgress = Executed
End Function

'================================================================================
' Name: updateProgressBar
' -------------------------------------------------------------------------------
' Description: Update CPU load progress bar according to the given parameter
' Called: From CPU load calculation main sub
' Pre-conditions: mainApp form loaded
' Post-conditions: CPU load progress bar updated
'================================================================================
Function updateProgressBar(Value)
    Dim upb_temp As Variant
    'Calculate current progress according to the given parameter (current row)
    upb_temp = ((Value / getLastRow) * 100)
    ' Value of bar can't be greater than 100
    If upb_temp > 100 Then
        upb_temp = 100
    End If
    ' Update progress bar
    mainApp.CPUloadProgressBar.Value = upb_temp
End Function

'================================================================================
' Name: InputProgressBar
' -------------------------------------------------------------------------------
' Description: Update Input Length progress bar according to the given parameter
' Called:
' Pre-conditions: mainApp form loaded
' Post-conditions: Input Length progress bar updated
'================================================================================
Function InputProgressBar(Value As Double)
    Dim ipb_tempVal As Variant
    
    ipb_tempVal = ((Value / 65535) * 100)
    mainApp.InputProgressBar.Value = ipb_tempVal
End Function

'================================================================================
' Name: CheckInputs
' -------------------------------------------------------------------------------
' Description: Perform basic tests of the required inputs
' Called: Before execution of CPU load main sub
' Pre-conditions: --
' Post-conditions: Basic test for Inputs done and errors adviced to user
'================================================================================
Function CheckInputs()

    Dim wSheet As Worksheet
    Dim tempCheck As Variant
    Dim currentWorkSheet As Variant

    ' Initialize return value, 1 means everything is correct
    CheckInputs = 1
    'Does Inputs worksheet exist?
    currentWorkSheet = "Inputs"
    CheckInputs = Check_WorkSheetExist(currentWorkSheet)
    If CheckInputs = 0 Then
        MsgBox "Worksheet " + currentWorkSheet + " does not exist!", vbCritical, "Input checking.."
    End If
    'Check inputs
    If Worksheets("Inputs").Cells(IN_timeWindow_R, IN_timeWindow_C).Value = "" And CheckInputs = 1 Then
        CheckInputs = 0
        MsgBox "Time Window must not be empty! Please enter a valid value", vbCritical, "Input checking.."
    End If
    If Worksheets("Inputs").Cells(IN_idleTaskId_R, IN_idleTaskId_C).Value = "" And CheckInputs = 1 Then
        CheckInputs = 0
        MsgBox "Idle Task Id must not be empty! Please enter a valid value", vbCritical, "Input checking.."
    End If
    If Worksheets("Inputs").Cells(IN_NoIntId_R, IN_NoIntId_C).Value = "" And CheckInputs = 1 Then
        CheckInputs = 0
        MsgBox "No Interrupt Id must not be empty! Please enter a valid value", vbCritical, "Input checking.."
    End If
    If Worksheets("Inputs").Cells(IN_NbDirectISRs_R, IN_NbDirectISRs_C).Value = "" And CheckInputs = 1 Then
        CheckInputs = 0
        MsgBox "Number of ISR Direct Functions must not be empty! Please enter a valid value", vbCritical, "Input checking.."
    End If
    If Worksheets("Inputs").Cells(IN_NbBkgTasks_R, IN_NbBkgTasks_C).Value = "" And CheckInputs = 1 Then
        CheckInputs = 0
        MsgBox "Number of Background Tasks must not be empty! Please enter a valid value", vbCritical, "Input checking.."
    End If
    If CheckInputs = 1 And Worksheets("Inputs").Cells(IN_NbBkgTasks_R, IN_NbBkgTasks_C).Value > 0 Then
        For tempCheck = 0 To (Worksheets("Inputs").Cells(IN_NbBkgTasks_R, IN_NbBkgTasks_C).Value - 1) Step 1
            If Worksheets("Inputs").Cells(IN_BkgTasksId_R, IN_BkgTasksId_C + tempCheck).Value = "" Then
                CheckInputs = 0
                MsgBox "Please enter valid Background Tasks Ids.", vbCritical, "Input checking.."
                Exit For
            End If
        Next tempCheck
    End If

    
    ' Check Direct ISRs worksheets
    If (Worksheets("Inputs").Cells(IN_NbDirectISRs_R, IN_NbDirectISRs_C).Value > 0) And (CheckInputs = 1) Then
        For tempCheck = 0 To (Worksheets("Inputs").Cells(IN_NbDirectISRs_R, IN_NbDirectISRs_C).Value - 1) Step 1
            CheckInputs = Check_WorkSheetExist(Worksheets("Inputs").Cells(IN_DirectISRsName_R, IN_DirectISRsName_C + tempCheck).Value)
            If CheckInputs = 0 Then
                MsgBox "Worksheet for Direct ISR: " + Worksheets("Inputs").Cells(IN_DirectISRsName_R, IN_DirectISRsName_C + tempCheck) + " does not exist!", vbCritical, "Input checking.."
                Exit For
            Else
                ' Check data
                If Worksheets(Worksheets("Inputs").Cells(IN_DirectISRsName_R, IN_DirectISRsName_C + tempCheck).Value).Cells(2, 1).Value <> "" And Worksheets(Worksheets("Inputs").Cells(IN_DirectISRsName_R, IN_DirectISRsName_C + tempCheck).Value).Cells(2, 2).Value <> "" Then
                Else
                    MsgBox "Verify timing data for Direct ISR: " + Worksheets("Inputs").Cells(IN_DirectISRsName_R, IN_DirectISRsName_C + tempCheck), vbCritical, "Input checking.."
                    CheckInputs = 0
                End If
            End If
        Next tempCheck
    End If
    ' Check Tasks worksheet
    If CheckInputs = 1 Then
        currentWorkSheet = "Tasks"
        CheckInputs = Check_WorkSheetExist(currentWorkSheet)
        If CheckInputs = 0 Then
            MsgBox "Worksheet " + currentWorkSheet + " does not exist!", vbCritical, "Input checking.."
        Else
            ' Worksheet exists
            ' Check data
            If Worksheets(currentWorkSheet).Cells(2, 1).Value <> "" And Worksheets(currentWorkSheet).Cells(2, 2).Value <> "" Then
            Else
                MsgBox "Verify timing data for OSEK Tasks", vbCritical, "Input checking.."
                CheckInputs = 0
            End If
        End If
    End If
    ' Check ISRs worksheet
    If CheckInputs = 1 Then
        currentWorkSheet = "ISRs"
        CheckInputs = Check_WorkSheetExist(currentWorkSheet)
        If CheckInputs = 0 Then
            MsgBox "Worksheet " + currentWorkSheet + " does not exist!", vbCritical, "Input checking.."
        Else
            ' Worksheet exists
            ' Check data
            If Worksheets(currentWorkSheet).Cells(2, 1).Value <> "" And Worksheets(currentWorkSheet).Cells(2, 2).Value <> "" Then
            Else
                MsgBox "Verify timing data for OSEK ISRs", vbCritical, "Input checking.."
                CheckInputs = 0
            End If
        End If
    End If
    ' Check Results worksheet
    If CheckInputs = 1 Then
        currentWorkSheet = "Results"
        CheckInputs = Check_WorkSheetExist(currentWorkSheet)
        If CheckInputs = 0 Then
            MsgBox "Worksheet " + currentWorkSheet + " does not exist!", vbCritical, "Input checking.."
        Else
        End If
    End If
End Function

'================================================================================
' Name: Check_WorkSheetExist
' -------------------------------------------------------------------------------
' Description: Function to verify if the specified worksheet exist in the workbook
' Called: From CheckInputs
' Pre-conditions: --
' Post-conditions: worksheet verified. Function returns 1 if it exist, 0 if not
'================================================================================
Function Check_WorkSheetExist(wSheetName)
    On Error Resume Next
    Set wSheet = Sheets(wSheetName)
    If wSheet Is Nothing Then
        ' Fatal Error, worksheet missing
        Check_WorkSheetExist = 0
        Set wSheet = Nothing
    Else 'Does exist
        Set wSheet = Nothing
        Check_WorkSheetExist = 1
    End If
    On Error GoTo 0
End Function

'================================================================================
' Name: setLastRow
' -------------------------------------------------------------------------------
' Description: Verify which is the last row (until detection of the firs
'              empty cell in column 1 of Tasks worksheet
' Called: --
' Pre-conditions: Inputs verified
' Post-conditions: CPUloadProgressLastRow set
'================================================================================
Function setLastRow()
    Dim glr_currentRow As Double
    Dim glr_exitWhile As Variant
    
    CPUloadProgressLastRow = 0
    glr_currentRow = 0
    glr_exitWhile = 0
    setLastRow = 0
    
    While (glr_currentRow < 65536) And (glr_exitWhile = 0) And CPUloadCalcCancel = 0
        ' Update Input progress bar
        Call InputProgressBar(glr_currentRow)
        If Worksheets("Tasks").Cells(2 + glr_currentRow, 1).Value = "" Then
            setLastRow = glr_currentRow
            CPUloadProgressLastRow = setLastRow
            glr_exitWhile = 1
        Else
        End If
        glr_currentRow = glr_currentRow + 1
        DoEvents
    Wend
    ' At the end Input progress bar must be at 100%
    mainApp.InputProgressBar.Value = 100
    
End Function

'================================================================================
' Name: getLastRow
' -------------------------------------------------------------------------------
' Description: Return the last row to be processed (into Task worksheet)
' Called: --
' Pre-conditions: setLastRow already executed
' Post-conditions: value of last row returned
'================================================================================
Function getLastRow()
    getLastRow = CPUloadProgressLastRow
End Function

'================================================================================
' Name: ImportCSVfile
' -------------------------------------------------------------------------------
' Description: Return the last row to be processed (into Task worksheet)
' Called: --
' Pre-conditions: setLastRow already executed
' Post-conditions: value of last row returned
'================================================================================
Sub ImportCSVfile(ByVal file_name As String, ByVal sheet_name As String)
Dim work_sheet As Worksheet
Dim query_table As QueryTable
Dim new_chart As Chart

    ' Load the CSV file.
    Set work_sheet = Sheets(sheet_name)
    Set query_table = work_sheet.QueryTables.Add(Connection:="TEXT;" + file_name, Destination:=work_sheet.Range("A1"))
    With query_table
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False

        ' Set the data types for the columns.
        .TextFileColumnDataTypes = Array(xlGeneralFormat, xlGeneralFormat)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub

Function doFileQuery(filename As String, outSheet As String) As Boolean
    Dim rootDir As String
    rootDir = ActiveWorkbook.Path
    Dim connectionName As String
    'connectionName = "TEXT;" + rootDir + "\" + filename
    connectionName = "TEXT;" + filename
    With Worksheets(outSheet).QueryTables.Add(Connection:=connectionName, Destination:=Worksheets(outSheet).Range("A1"))
        .Name = filename
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .Refresh BackgroundQuery:=False
    End With
End Function

Function GraphUpdate(VarI As Variant)

    Worksheets("Graphs").Activate
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SeriesCollection(1).XValues = ActiveSheet.Range(ActiveSheet.Cells(2, 1), ActiveSheet.Cells(VarI, 1))
    ActiveChart.SeriesCollection(1).Values = ActiveSheet.Range(ActiveSheet.Cells(2, 3), ActiveSheet.Cells(VarI, 3))
    ActiveChart.SeriesCollection(2).XValues = ActiveSheet.Range(ActiveSheet.Cells(2, 1), ActiveSheet.Cells(VarI, 1))
    ActiveChart.SeriesCollection(2).Values = ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(VarI, 4))
    ActiveChart.SeriesCollection(3).XValues = ActiveSheet.Range(ActiveSheet.Cells(2, 1), ActiveSheet.Cells(VarI, 1))
    ActiveChart.SeriesCollection(3).Values = ActiveSheet.Range(ActiveSheet.Cells(2, 5), ActiveSheet.Cells(VarI, 5))
    'ActiveChart.ChartTitle.Select
End Function


