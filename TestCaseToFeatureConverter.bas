Attribute VB_Name = "TestCaseToFeatureConverter"
Option Explicit

' Constants for column mapping (can be easily modified by user)
Private Const COL_TEST_CASE_ID As Integer = 1      ' Column A
Private Const COL_SCENARIO_NAME As Integer = 2     ' Column B
Private Const COL_STEPS As Integer = 3             ' Column C
Private Const COL_EXPECTED_RESULT As Integer = 4   ' Column D
Private Const START_ROW As Integer = 2             ' Data starts from row 2 (header in row 1)

' Main subroutine to convert test cases to feature files
Sub ConvertTestCasesToFeatureFiles()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim outputDir As String
    Dim featureContent As String
    Dim fileName As String
    Dim testCaseID As String
    Dim scenarioName As String
    Dim steps As String
    Dim expectedResult As String
    
    ' Get the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, COL_TEST_CASE_ID).End(xlUp).Row
    
    ' Check if we have any data
    If lastRow < START_ROW Then
        MsgBox "No test case data found. Please ensure data starts from row " & START_ROW & ".", vbExclamation
        Exit Sub
    End If
    
    ' Get output directory from user
    outputDir = GetOutputDirectory()
    If outputDir = "" Then
        MsgBox "No output directory selected. Process cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Process each test case
    For i = START_ROW To lastRow
        ' Extract data from each row
        testCaseID = Trim(ws.Cells(i, COL_TEST_CASE_ID).Value)
        scenarioName = Trim(ws.Cells(i, COL_SCENARIO_NAME).Value)
        steps = Trim(ws.Cells(i, COL_STEPS).Value)
        expectedResult = Trim(ws.Cells(i, COL_EXPECTED_RESULT).Value)
        
        ' Skip empty rows
        If testCaseID = "" And scenarioName = "" Then
            GoTo NextRow
        End If
        
        ' Generate feature content
        featureContent = GenerateFeatureContent(testCaseID, scenarioName, steps, expectedResult)
        
        ' Generate filename
        fileName = GenerateFileName(testCaseID, scenarioName)
        
        ' Create feature file
        CreateFeatureFile outputDir, fileName, featureContent
        
NextRow:
    Next i
    
    ' Show completion message
    MsgBox "Feature files have been successfully created in: " & vbCrLf & outputDir, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

' Function to get output directory from user
Private Function GetOutputDirectory() As String
    Dim folderPath As String
    
    folderPath = Application.GetOpenFilename("All Files (*.*),*.*", , "Select any file in the desired output directory")
    
    If folderPath <> "False" Then
        ' Extract directory path from file path
        GetOutputDirectory = Left(folderPath, InStrRev(folderPath, "\") - 1)
    Else
        GetOutputDirectory = ""
    End If
End Function

' Function to generate feature content in Gherkin format
Private Function GenerateFeatureContent(testCaseID As String, scenarioName As String, steps As String, expectedResult As String) As String
    Dim content As String
    
    ' Create feature header
    content = "Feature: " & scenarioName & vbCrLf & vbCrLf
    
    ' Create scenario
    content = content & "  Scenario: " & scenarioName & " - " & testCaseID & vbCrLf
    
    ' Add Given step
    content = content & "    Given the test setup is ready" & vbCrLf
    
    ' Add When step with actual steps
    content = content & "    When I perform the following steps:" & vbCrLf
    content = content & "      " & steps & vbCrLf
    
    ' Add Then step with expected result
    content = content & "    Then I expect:" & vbCrLf
    content = content & "      " & expectedResult & vbCrLf
    
    GenerateFeatureContent = content
End Function

' Function to generate filename from test case data
Private Function GenerateFileName(testCaseID As String, scenarioName As String) As String
    Dim fileName As String
    
    ' Use test case ID if available, otherwise use scenario name
    If testCaseID <> "" Then
        fileName = CleanFileName(testCaseID)
    ElseIf scenarioName <> "" Then
        fileName = CleanFileName(scenarioName)
    Else
        fileName = "TestCase_" & Format(Now, "yyyymmdd_hhmmss")
    End If
    
    ' Add .feature extension
    GenerateFileName = fileName & ".feature"
End Function

' Function to clean filename (remove invalid characters)
Private Function CleanFileName(fileName As String) As String
    Dim cleanName As String
    Dim i As Integer
    Dim char As String
    
    cleanName = fileName
    
    ' Replace invalid characters with underscore
    For i = 1 To Len(cleanName)
        char = Mid(cleanName, i, 1)
        If InStr("\/:*?""<>|", char) > 0 Then
            cleanName = Replace(cleanName, char, "_")
        End If
    Next i
    
    ' Remove leading/trailing spaces and dots
    cleanName = Trim(cleanName)
    If Right(cleanName, 1) = "." Then
        cleanName = Left(cleanName, Len(cleanName) - 1)
    End If
    
    CleanFileName = cleanName
End Function

' Subroutine to create feature file
Private Sub CreateFeatureFile(outputDir As String, fileName As String, content As String)
    Dim filePath As String
    Dim fileNum As Integer
    
    ' Create full file path
    filePath = outputDir & "\" & fileName
    
    ' Get next available file number
    fileNum = FreeFile
    
    ' Open file for writing
    Open filePath For Output As fileNum
    Print #fileNum, content
    Close fileNum
End Sub

' Optional: Function to export to CSV format
Sub ExportTestCasesToCSV()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim csvContent As String
    Dim outputDir As String
    Dim csvPath As String
    Dim i As Long
    
    ' Get the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, COL_TEST_CASE_ID).End(xlUp).Row
    
    ' Check if we have any data
    If lastRow < START_ROW Then
        MsgBox "No test case data found. Please ensure data starts from row " & START_ROW & ".", vbExclamation
        Exit Sub
    End If
    
    ' Get output directory
    outputDir = GetOutputDirectory()
    If outputDir = "" Then
        MsgBox "No output directory selected. Process cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Create CSV header
    csvContent = "TestCaseID,ScenarioName,Steps,ExpectedResult" & vbCrLf
    
    ' Add data rows
    For i = START_ROW To lastRow
        csvContent = csvContent & """" & Replace(ws.Cells(i, COL_TEST_CASE_ID).Value, """", """""") & ""","
        csvContent = csvContent & """" & Replace(ws.Cells(i, COL_SCENARIO_NAME).Value, """", """""") & ""","
        csvContent = csvContent & """" & Replace(ws.Cells(i, COL_STEPS).Value, """", """""") & ""","
        csvContent = csvContent & """" & Replace(ws.Cells(i, COL_EXPECTED_RESULT).Value, """", """""") & """" & vbCrLf
    Next i
    
    ' Create CSV file
    csvPath = outputDir & "\TestCases.csv"
    Dim fileNum As Integer
    fileNum = FreeFile
    Open csvPath For Output As fileNum
    Print #fileNum, csvContent
    Close fileNum
    
    MsgBox "CSV file has been successfully created at: " & vbCrLf & csvPath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

' Function to validate Excel sheet structure
Sub ValidateSheetStructure()
    Dim ws As Worksheet
    Dim headerRow As Range
    Dim missingColumns As String
    
    Set ws = ActiveSheet
    
    ' Check if we have data in the expected columns
    If ws.Cells(1, COL_TEST_CASE_ID).Value = "" Then
        missingColumns = missingColumns & "Column A (Test Case ID), "
    End If
    
    If ws.Cells(1, COL_SCENARIO_NAME).Value = "" Then
        missingColumns = missingColumns & "Column B (Scenario Name), "
    End If
    
    If ws.Cells(1, COL_STEPS).Value = "" Then
        missingColumns = missingColumns & "Column C (Steps), "
    End If
    
    If ws.Cells(1, COL_EXPECTED_RESULT).Value = "" Then
        missingColumns = missingColumns & "Column D (Expected Result), "
    End If
    
    If missingColumns <> "" Then
        missingColumns = Left(missingColumns, Len(missingColumns) - 2) ' Remove trailing comma and space
        MsgBox "Warning: The following columns appear to be missing or empty:" & vbCrLf & missingColumns & vbCrLf & vbCrLf & "Please ensure your data is structured correctly.", vbExclamation
    Else
        MsgBox "Sheet structure appears to be valid!", vbInformation
    End If
End Sub

' Function to show usage instructions
Sub ShowInstructions()
    Dim instructions As String
    
    instructions = "TEST CASE TO FEATURE FILE CONVERTER" & vbCrLf & vbCrLf
    instructions = instructions & "SETUP INSTRUCTIONS:" & vbCrLf
    instructions = instructions & "1. Ensure your Excel sheet has the following structure:" & vbCrLf
    instructions = instructions & "   - Row 1: Headers" & vbCrLf
    instructions = instructions & "   - Column A: Test Case ID" & vbCrLf
    instructions = instructions & "   - Column B: Scenario Name" & vbCrLf
    instructions = instructions & "   - Column C: Steps" & vbCrLf
    instructions = instructions & "   - Column D: Expected Result" & vbCrLf
    instructions = instructions & "   - Data starts from Row 2" & vbCrLf & vbCrLf
    instructions = instructions & "USAGE:" & vbCrLf
    instructions = instructions & "1. Select the worksheet containing your test cases" & vbCrLf
    instructions = instructions & "2. Run 'ConvertTestCasesToFeatureFiles' to create feature files" & vbCrLf
    instructions = instructions & "3. Run 'ExportTestCasesToCSV' to create a CSV file" & vbCrLf
    instructions = instructions & "4. Run 'ValidateSheetStructure' to check your data format" & vbCrLf & vbCrLf
    instructions = instructions & "CUSTOMIZATION:" & vbCrLf
    instructions = instructions & "Modify the constants at the top of the module to change column mappings."
    
    MsgBox instructions, vbInformation, "Instructions"
End Sub