VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfEditProject 
   Caption         =   "Edit or Add Project"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   OleObjectBlob   =   "usfEditProject.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "usfEditProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'@Description("Generates list of project numbers with project names from the active week report table")
Public Sub RefreshProjectList(tblCurrentWeek As ListObject, Optional preserveSelection As String = "")

'Generic Variables
    Dim i As Integer, j As Integer, k As Integer

'Use ArrayList array type to store concated project details.  ArrayList has sort function.
    Dim projectList     As Object:  Set projectList = CreateObject("System.Collections.ArrayList")

'Build concated project details and add to arraylist
    For i = 1 To tblCurrentWeek.DataBodyRange.Rows.Count
        projectList.Add tblCurrentWeek.DataBodyRange.Cells(i, 1) & " - " & tblCurrentWeek.DataBodyRange.Cells(i, 2)
    Next i

projectList.Sort

'Clear combobox contents and reload with arraylist, if being called from userform set the value to the last used
    cmbProjects.Clear
    cmbProjects.List = projectList.toarray()
    cmbProjects.Value = preserveSelection

End Sub

'@Description("Creates entries in related tables for new projects")
Private Sub cmbAddProject_Click()

'Generic Variables
    Dim i As Integer, j As Integer, k As Integer

    Dim wb                      As Workbook:    Set wb = ThisWorkbook
    Dim wsProjectData           As Worksheet:   Set wsProjectData = wb.Worksheets("project list")
    Dim wsActiveSheet           As Worksheet:   Set wsActiveSheet = wb.ActiveSheet
    Dim tblImportedProjectList  As ListObject:  Set tblImportedProjectList = wsProjectData.ListObjects("q_l_projectList")
    Dim tblManualJobList        As ListObject:  Set tblManualJobList = wsProjectData.ListObjects("tbl_userDefinedProjectList")
    Dim tblProjectStartDates    As ListObject:  Set tblProjectStartDates = wsProjectData.ListObjects("tbl_startDates")
    Dim tblActiveWeek           As ListObject:  Set tblActiveWeek = Main.GetWeeklyTableListObjectFromWorksheet(wsActiveSheet)
    
    Dim projectFound            As Boolean:     projectFound = False
    Dim newTblRow               As Range
    Dim budgetHoursColumnNum    As Integer
    
'As current week table set when run check to make sure the activesheet actually contains a weekly report table
    If tblActiveWeek Is Nothing Then
        MsgBox "A worksheet containing weekly reporting figures must be active.  Please select an appropriate worksheet and try again."
        Exit Sub
    End If

'All the following blocks have the same general function
'   (1) check the target table to see if the job number exists.  If it does update the values and set projectFound flag to true
'   (2) if projectFound flag is false, create new entries and populate values
'   ***note*** if job numbers are present that should have been caught by txtAddProjectNumber_Exit method

    'Manual Job List table check and update
    With tblManualJobList.DataBodyRange
        For i = 1 To .Rows.Count
            If .Cells(i, 1) = txtAddProjectNumber Then
                .Cells(i, 2) = txtAddProjectName.Value
                projectFound = True
            End If
        Next i
     End With
        
    'Manual Job List table add new
    If projectFound = False Then
        Set newTblRow = tblManualJobList.ListRows.Add.Range
        With newTblRow
            .Cells(1) = txtAddProjectNumber.Value
            .Cells(2) = txtAddProjectName.Value
        End With
    End If
            
    'Project Start Date table check and update
    projectFound = False
    With tblProjectStartDates.DataBodyRange
        For i = 1 To .Rows.Count
            If .Cells(i, 1) = txtAddProjectNumber Then
                .Cells(i, 2) = txtEditStartDate.Value
                .Cells(i, 3) = txtEditBudgetHours.Value
                projectFound = True
            End If
        Next i
    End With
    
    'Project Start Date table add new
    If projectFound = False Then
        Set newTblRow = tblProjectStartDates.ListRows.Add.Range
        With newTblRow
            .Cells(1) = txtAddProjectNumber.Value
            .Cells(2) = txtAddStartDate.Value
            .Cells(3) = txtAddBudgetHours.Value
        End With
    End If
    
    'Weekly Report table check and update
    projectFound = False
    With tblActiveWeek.DataBodyRange
        For i = 1 To .Rows.Count
            If .Cells(i, 1) = txtAddProjectNumber Then
                .Cells(i, 2) = txtAddProjectName.Value
                projectFound = True
            End If
        Next i
    End With
    
    'Weekly Report table add new
    If projectFound = False Then
        Set newTblRow = tblActiveWeek.ListRows.Add.Range
        With newTblRow
            .Cells(1) = txtAddProjectNumber.Value
        End With
    End If
    
End Sub


'@Description("On selection change check if project is on master list and if it is disable the project name textbox.  Update project name, start date, and budget hours from repective tables")
Private Sub cmbProjects_Change()

'Generic Variables
    Dim i As Integer, j As Integer, k As Integer
    
    Dim wb                      As Workbook:    Set wb = ThisWorkbook
    Dim wsProjectData           As Worksheet:   Set wsProjectData = wb.Worksheets("project list")
    Dim tblImportedProjectList  As ListObject:  Set tblImportedProjectList = wsProjectData.ListObjects("q_l_projectList")
    Dim tblManualJobList        As ListObject:  Set tblManualJobList = wsProjectData.ListObjects("tbl_userDefinedProjectList")
    Dim tblProjectStartDates    As ListObject:  Set tblProjectStartDates = wsProjectData.ListObjects("tbl_startDates")
    
    Dim selectedProject         As String:      selectedProject = cmbProjects.Value
    Dim sArr()                  As String:      sArr = Split(selectedProject, " - ")
    Dim projectExists           As Boolean:     projectExists = False
    
'Set default textbox properties
    txtEditProjectName.Enabled = True
    txtEditProjectName.BackColor = &H80000005
    txtEditStartDate.Value = Null

'Check if project is not blank
'Check if project exists on the master list.  If it does update the name textbox, disable, and change background color to grey
'   set flag to indicate found
'If flag set not found then check manual project list.  return project name to name textbox (editable).
'Check start date list for start date and budget hours, update respective text boxes
    If cmbProjects.Value <> "" Then
        With tblImportedProjectList.DataBodyRange
            For i = 1 To .Rows.Count
                If .Cells(i, 1) = sArr(0) Then
                    projectExists = True
                    txtEditProjectName.Enabled = False
                    txtEditProjectName.BackColor = &H80000016
                    txtEditProjectName.Value = .Cells(i, 2)
                End If
            Next i
        End With
        
        With tblManualJobList.DataBodyRange
            If projectExists = False Then
                For i = 1 To .Rows.Count
                    If .Cells(i, 1) = sArr(0) Then
                        projectExists = True
                        txtEditProjectName.Value = .Cells(i, 2)
                    End If
                Next i
            End If
        End With
        
        If projectExists = False Then txtEditProjectName.Value = ""
        
        With tblProjectStartDates.DataBodyRange
            For i = 1 To .Rows.Count
                If .Cells(i, 1) = sArr(0) Then
                    txtEditStartDate.Value = .Cells(i, 2)
                    txtEditBudgetHours.Value = .Cells(i, 3)
                End If
            Next i
        End With
    End If
        
End Sub

'@Description("On click sets updates project name, start date, and budget hours tables with new values")
Private Sub cmbUpdateProject_Click()

'Generic Variables
    Dim i As Integer, j As Integer, k As Integer

    Dim wb                      As Workbook:    Set wb = ThisWorkbook
    Dim wsProjectData           As Worksheet:   Set wsProjectData = wb.Worksheets("project list")
    Dim wsActiveSheet           As Worksheet:   Set wsActiveSheet = wb.ActiveSheet
    
    Dim tblManualJobList        As ListObject:  Set tblManualJobList = wsProjectData.ListObjects("tbl_userDefinedProjectList")
    Dim tblProjectStartDates    As ListObject:  Set tblProjectStartDates = wsProjectData.ListObjects("tbl_startDates")
    Dim tblActiveWeek           As ListObject:  Set tblActiveWeek = Main.GetWeeklyTableListObjectFromWorksheet(wsActiveSheet)
    
    Dim selectedProject         As String:      selectedProject = cmbProjects.Value
    Dim sArr()                  As String:      sArr = Split(selectedProject, " - ")
    Dim projectExists           As Boolean:     projectExists = False
    Dim errorMessage            As String
    
'Check if project is blank, indicate error and exit
'Otherwise search resepective tables for job number and update fields as necessary.  Set flags to indicate an update was completed.
'If the update was not completed set errorMessage details
    If cmbProjects.Value = "" Then
        MsgBox "Please select a project", vbOKOnly + vbExclamation, "Input Error"
        Exit Sub
    Else
        With tblManualJobList.DataBodyRange
            For i = 1 To .Rows.Count
                If .Cells(i, 1) = sArr(0) Then
                    .Cells(i, 2) = txtEditProjectName.Value
                    projectExists = True
                End If
            Next i
        End With
        
        If projectExists = False Then errorMessage = errorMessage + "- Manual Job List Table" & vbCrLf
        
        projectExists = False
        With tblProjectStartDates.DataBodyRange
            For i = 1 To .Rows.Count
                If .Cells(i, 1) = sArr(0) Then
                    .Cells(i, 2) = txtEditStartDate.Value
                    .Cells(i, 3) = txtEditBudgetHours.Value
                    projectExists = True
                End If
            Next i
        End With
        
        If projectExists = False Then errorMessage = errorMessage + "- Start Date Table" & vbCrLf
        
    End If

'Display error message if any
    If errorMessage <> "" Then
        MsgBox "The following table were not updated" & vbCrLf & vbCrLf & errorMessage & vbCrLf & "Please ensure Job Number is listed and try again.", _
            vbOKOnly + vbCritical, "Error"
    End If
   
'Check that current worksheet is a weekly report.  If so update project list otherwise close the userform and display error message.
    If Not (tblActiveWeek Is Nothing) Then
        RefreshProjectList tblActiveWeek, sArr(0) & " - " & txtEditProjectName.Value
    Else
        MsgBox "Unable to refresh projects list.  Return to Weekly Report page and click edit/add project button to refresh.", vbOKOnly + vbInformation
        usfEditProject.Hide
    End If

End Sub

'@Description("Calls method to check that project number is unique when textbox is updated")
Private Sub txtAddProjectNumber_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = ValidateAddProjectNumber(txtAddProjectNumber.Text)
End Sub

'@Description("When trying to assign a new project number the following method will check relevant tables to see if the job number exists.  If it does an message is displayed indicating locations where was found")
Private Function ValidateAddProjectNumber(jobNumber As String) As Boolean

'Generic Variables
    Dim i As Integer, j As Integer, k As Integer

    Dim wb                      As Workbook:    Set wb = ThisWorkbook
    Dim wsProjectData           As Worksheet:   Set wsProjectData = wb.Worksheets("project list")
    Dim wsActiveSheet           As Worksheet:   Set wsActiveSheet = wb.ActiveSheet
    Dim tblImportedProjectList  As ListObject:  Set tblImportedProjectList = wsProjectData.ListObjects("q_l_projectList")
    Dim tblManualJobList        As ListObject:  Set tblManualJobList = wsProjectData.ListObjects("tbl_userDefinedProjectList")
    Dim tblProjectStartDates    As ListObject:  Set tblProjectStartDates = wsProjectData.ListObjects("tbl_startDates")
    Dim tblActiveWeek           As ListObject:  Set tblActiveWeek = Main.GetWeeklyTableListObjectFromWorksheet(wsActiveSheet)
    
    Dim errorMessage            As String
    
'Check if on weekly report worksheet.  If not display error and close form.
    If tblActiveWeek Is Nothing Then
        MsgBox "Unable to load projects list.  Return to Weekly Report page and click edit/add project button to try again.", vbOKOnly + vbInformation
        usfEditProject.Hide
        ValidateAddProjectNumber = True
        Exit Function
    End If

'each of the following blocks scans the tables for the job number.  If found it adds its details to the error message
    With tblImportedProjectList.DataBodyRange
        For i = 1 To .Rows.Count
            If .Cells(i, 1) = txtAddProjectNumber.Value Then
                errorMessage = errorMessage & "- Imported Project List" & vbCrLf
            End If
        Next i
    End With
    
    With tblProjectStartDates.DataBodyRange
        For i = 1 To .Rows.Count
            If .Cells(i, 1) = txtAddProjectNumber.Value Then
                errorMessage = errorMessage & "- Project Start Date" & vbCrLf
            End If
        Next i
    End With
    
    With tblActiveWeek.DataBodyRange
        For i = 1 To .Rows.Count
            If .Cells(i, 1) = txtAddProjectNumber.Value Then
                errorMessage = errorMessage & "- " & wsActiveSheet.Name & vbCrLf
            End If
        Next i
    End With

'Displays error message if any
    If errorMessage <> "" Then
        MsgBox "The project number currently exists in one or more of the following tables: " & _
        vbCrLf & vbCrLf & _
        errorMessage & _
        vbCrLf & _
        "Please enter a different number or remove existing references", _
        vbOKOnly + vbCritical, "Invalid Project Number"
        ValidateAddProjectNumber = True
        
        txtAddProjectNumber.Value = ""
    Else
        ValidateAddProjectNumber = False
    End If
    
    

End Function


