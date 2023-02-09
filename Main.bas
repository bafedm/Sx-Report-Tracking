Attribute VB_Name = "Main"
Option Explicit

Sub Button3_Click()
    usfEditProject.Show
    Main
End Sub

'@Folder("VBAProject")
Private Sub Main()

'Set debug mode
    Dim debugStatus As Boolean: debugStatus = False

'Generic Variables
    Dim i As Integer, j As Integer, k As Integer

'Assign Excel workbook/worksheets
    Dim wb                      As Workbook:                Set wb = ThisWorkbook
    Dim wsActiveSheet           As Worksheet:               Set wsActiveSheet = wb.ActiveSheet
    Dim wsProjectData           As Worksheet:               Set wsProjectData = wb.Worksheets("project list")

'Fixed activesheet for testing purposes
    If debugStatus = True Then Set wsActiveSheet = wb.Worksheets("23Week06")
    
    Dim tblActiveWeek           As ListObject:              Set tblActiveWeek = GetWeeklyTableListObjectFromWorksheet(wsActiveSheet)

'Assign selected workbook  tables
    Dim tblImportedProjectList  As ListObject:              Set tblImportedProjectList = wsProjectData.ListObjects("q_l_projectList")
    Dim tblManualJobList        As ListObject:              Set tblManualJobList = wsProjectData.ListObjects("tbl_userDefinedProjectList")
    Dim tblProjectStartDates    As ListObject:              Set tblProjectStartDates = wsProjectData.ListObjects("tbl_startDates")

'Dictionary to hold weekly tables.
    'Key:   Worksheet name in format of YYWeekXX
    'Value: weekly table data as List Object
    
    'Late binding (uncomment for distribution)
    'Dim dictWeeklyData          As Object:      Set dictWeeklyData = CreateObject("Scripting.Dictionary")
    
    'Early binding (for development)
    'Dim dictWeeklyData          As Scripting.Dictionary:    Set dictWeeklyData = New Scripting.Dictionary

    'Set dictWeeklyData = GenerateDictionaryOfWeeklyReports(wb)
    
    usfEditProject.RefreshProjectList tblActiveWeek
    
'Debuggin'
    If debugStatus = True Then
    
        Debug.Print "---New Run---"
        DebugOut "Import Project List Row Count", tblImportedProjectList.Range.Rows.Count
        DebugOut "Manual Job List Row Count", tblManualJobList.Range.Rows.Count
        DebugOut "Project Start Dates Row Count", tblProjectStartDates.Range.Rows.Count
        DebugOut "Active Week Table Name", TypeName(tblActiveWeek)

        For i = 0 To dictWeeklyData.Count - 1
            DebugOut dictWeeklyData.Keys()(i) & ": " & dictWeeklyData.Items()(i).Range.Rows.Count, "", ""
        Next i
        
    End If
End Sub




'!!!! NOTE:
'Select appropriate function declartion depending on distribution (Object) or development (Scripting[...])
'For distribution the Set Gen[...] line also needs to be uncommented

'@Description("Creates a dictionary containing weekly reports")
'Private Function GenerateDictionaryOfWeeklyReports(wb As Workbook) As Scripting.Dictionary
Private Function GenerateDictionaryOfWeeklyReports(wb As Workbook) As Object

Dim s       As String
Dim ws      As Worksheet
Dim lo      As ListObject
'Late binding (uncomment for distribution)
'Set GenerateDictionaryOfWeeklyReports = CreateObject("Scripting.Dictionary")

    'Late binding (uncomment for distribution)
    'Dim dict        As Object:      Set dict = CreateObject("Scripting.Dictionary")
    
    'Early binding (for development)
    Dim dict    As Scripting.Dictionary:    Set dict = New Scripting.Dictionary

For Each ws In wb.Worksheets
    If Not (GetWeeklyTableListObjectFromWorksheet(ws) Is Nothing) Then
        s = ws.Name
        Set lo = GetWeeklyTableListObjectFromWorksheet(ws)
        Set dict(s) = lo
    End If
Next ws

Set GenerateDictionaryOfWeeklyReports = dict
            
End Function

'@Description("Checks worksheet to see if it contains weekly report table and if it does returns it as a list object")
Public Function GetWeeklyTableListObjectFromWorksheet(ws As Worksheet) As ListObject

Dim lo As ListObject

For Each lo In ws.ListObjects
    If lo.Range(1, 1) = "Job number" Then
        Set GetWeeklyTableListObjectFromWorksheet = lo
    End If
Next lo


End Function


Private Sub DebugOut(outMessage As String, outValue As String, Optional delim As String = ": ")
    Debug.Print outMessage & delim & outValue
End Sub
