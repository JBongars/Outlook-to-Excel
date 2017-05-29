Option Explicit

Public Sub Export_Outlook()

''''''''''''''''''''''''''''''''''''''''''''''''
'                   Setup                      '
''''''''''''''''''''''''''''''''''''''''''''''''

    'Counter
    Dim i, j As Integer
        
    'xl$Item Declarations
    Dim xlApp As Excel.Application
    Dim xlwb As Excel.Workbook
    Dim xlws As Excel.Worksheet
    Dim MsgResponse As Variant
    
    'Create Excel Application
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    
    'Create New Workbook
    Set xlwb = xlApp.Workbooks.Add
    xlwb.Title = "Outlook-Report"
    
    
''''''''''''''''''''''''''''''''''''''''''''''''
'                  Routine                     '
''''''''''''''''''''''''''''''''''''''''''''''''
    
    Debug.Print "Opening Workbook"

    'With xlApp
        
        'Optimization
        xlApp.ScreenUpdating = False
        
        Call GenData(xlwb)
        xlwb.Sheets("Sheet1").Delete
        
        'Taking too long, try at your own risk.
        
        'MsgResponse = MsgBox("Do you want to export Inbox data?" & _
        '    Chr(10) & "This may significantly slow down this program...", _
        '    vbYesNo, "Export Inbox Data")
        '
        'If MsgResponse = vbYes Then Call GenInbox
        
        'Reset
        xlApp.ScreenUpdating = True

    'End With
    
    Set xlApp = Nothing
    Set xlwb = Nothing

End Sub


Public Sub GenData(wb As Workbook)

    Debug.Print "Starting Contacts..." & Chr(10)

    Dim fMAPI As MAPIFolder
    Dim fItem As Items 'Subject to change with class
    Dim fObj As Object
    Dim olNameSpace As NameSpace
    Dim i As Integer
    
    Set olNameSpace = Application.GetNamespace("MAPI")
    
    'Adds multiple stores to current namespace (PST Files)
    'olNameSpace.AddStore ("C:\Users\Owner\Documents\Outlook Files\stephen@honan.com.sg.pst")
    
    On Error GoTo 0

    'Avoids all the previous declarations from main class
    With wb
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''
        '                 CONTACTS                   '
        ''''''''''''''''''''''''''''''''''''''''''''''
        
        'Search for default folder in current class
        Set fMAPI = olNameSpace.GetDefaultFolder(olFolderContacts)
        Set fItem = fMAPI.Items
    
        Debug.Print "Styling Worksheet..." & Chr(10)
    
        'Declare new worksheet
        .Sheets.Add().Name = "Contacts"
        
        With .Sheets("Contacts")
            .Activate
                
            'Link below for all possible fields
            'https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.contactitem_properties.aspx
            
            .Cells(1, 1).Value = "Account"
            .Cells(1, 2).Value = "Name"
            .Cells(1, 3).Value = "Company"
            .Cells(1, 4).Value = "Email Address"
            .Cells(1, 5).Value = "Direct-No."
            .Cells(1, 6).Value = "Work-No."
            .Cells(1, 7).Value = "Home-No."
            .Cells(1, 8).Value = "Fax-No"
            .Cells(1, 9).Value = "Mobile-No"
            .Cells(1, 10).Value = "Mailing-Address"
            
            i = 2 'Start counter here
            
            For Each fObj In fItem
                If TypeName(fObj) = "ContactItem" Then 'Validation
                    
                    'Debug.Print "Oh NO!"
                    'Debug.Print TypeName(fObj)
                    

                    .Cells(i, 1).Value = fObj.Account
                    .Cells(i, 2).Value = fObj.FullName
                    .Cells(i, 3).Value = fObj.CompanyName
                    .Cells(i, 4).Value = fObj.Email1Address
                    .Cells(i, 5).Value = fObj.BusinessTelephoneNumber
                    .Cells(i, 6).Value = fObj.CompanyMainTelephoneNumber
                    .Cells(i, 7).Value = fObj.HomeTelephoneNumber
                    .Cells(i, 8).Value = fObj.BusinessFaxNumber
                    .Cells(i, 9).Value = fObj.MobileTelephoneNumber
                    .Cells(i, 10).Value = fObj.MailingAddress
                    
                    Debug.Print TypeName(fObj) & "@" & i - 1 & " - " & IsEmpty(.Cells(i, 4).Value) & ": " & fObj.Email1Address

    
                    '.ActiveSheet.Hyperlinks.Add _
                    '    Anchor:=Cells(i, 4), _
                    '    Address:="mailto:" & Cells(i, 4).Value, _
                    '    TextToDisplay:=Cells(i, 4).Value
    
                    i = i + 1 'counter
                End If
            Next fObj
            
            'Formatting
            .Cells.RowHeight = 18.6
            .Columns("A:J").EntireColumn.AutoFit
            .Range("A1:J1").Font.Bold = True
            
            With .Range("A1:J1").Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            
            Range("A1").Select
            
            'Debug
            Debug.Print "Cleaning up..." & Chr(10)
            
            'Null out the variables
            Set fItem = Nothing
            Set fMAPI = Nothing
            Set fObj = Nothing
            
        End With
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''
        '                 CALENDAR                   '
        ''''''''''''''''''''''''''''''''''''''''''''''
        
        'Search for default folder in current class
        Set fMAPI = olNameSpace.GetDefaultFolder(olFolderCalendar)
        Set fItem = fMAPI.Items
    
        Debug.Print "Styling Worksheet..." & Chr(10)
    
        'Declare new worksheet
        .Sheets.Add().Name = "Appointments"
           
        With .Sheets("Appointments")
            .Activate
               
            'Link below for all possible fields
            'https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.appointmentitem_properties.aspx
            
            .Cells(1, 1).Value = "Start"
            .Cells(1, 2).Value = "End"
            .Cells(1, 3).Value = "Reminder"
            .Cells(1, 4).Value = "Importance"
            .Cells(1, 5).Value = "Subject"
            .Cells(1, 6).Value = "Location"
            .Cells(1, 7).Value = "Organizer"
            .Cells(1, 8).Value = "Recipients"
            .Cells(1, 9).Value = "Created On"
            .Cells(1, 10).Value = "Remarks"
            
            i = 2 'Start counter here
            
            For Each fObj In fItem
                If TypeName(fObj) = "AppointmentItem" Then 'Validation

                    .Cells(i, 1).Value = fObj.Start
                    If fObj.AllDayEvent = False Then .Cells(i, 2).Value = fObj.End
                    If fObj.ReminderSet = True Then .Cells(i, 3).Value = fObj.ReminderMinutesBeforeStart
                    .Cells(i, 4).Value = fObj.Importance
                    .Cells(i, 5).Value = fObj.Subject
                    .Cells(i, 6).Value = fObj.Location
                    .Cells(i, 7).Value = fObj.Organizer
                    '.Cells(i, 8).Value = fObj.Recipients
                    .Cells(i, 9).Value = fObj.CreationTime
                    .Cells(i, 10).Value = fObj.Body
                    
                    Debug.Print TypeName(fObj) & "@" & i - 1 & " - " & IsEmpty(.Cells(i, 1).Value) & ": " & fObj.Start
    
                    i = i + 1 'counter
                End If
            Next fObj
            
            'Formatting
            .Cells.RowHeight = 18.6
            .Columns("A:J").EntireColumn.AutoFit
            .Range("A1:J1").Font.Bold = True
            
            With .Range("A1:J1").Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            
            Range("A1").Select
            
            'Debug
            Debug.Print "Cleaning up..." & Chr(10)
            
            'Null out the variables
            Set fItem = Nothing
            Set fMAPI = Nothing
            Set fObj = Nothing
                    
        End With
        
                
        ''''''''''''''''''''''''''''''''''''''''''''''
        '                   TASKS                    '
        ''''''''''''''''''''''''''''''''''''''''''''''
        
        'Search for default folder in current class
        Set fMAPI = olNameSpace.GetDefaultFolder(olFolderTasks)
        Set fItem = fMAPI.Items
    
        Debug.Print "Styling Worksheet..." & Chr(10)
    
        'Declare new worksheet
        .Sheets.Add().Name = "Tasks"
        
        With .Sheets("Tasks")
            .Activate
                
            'Link below for all possible fields
            'https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.taskitem_properties.aspx
            
            .Cells(1, 1).Value = "Completed on:"
            .Cells(1, 2).Value = "Start"
            .Cells(1, 3).Value = "Due"
            .Cells(1, 4).Value = "Reminder"
            .Cells(1, 5).Value = "Subject"
            .Cells(1, 6).Value = "Importance"
            .Cells(1, 7).Value = "IsReccuring"
            .Cells(1, 8).Value = "Status"
            '.Cells(1, 9).Value = "Recipents"
            .Cells(1, 10).Value = "Remarks"
            
            i = 2 'Start counter here
            
            For Each fObj In fItem
                'Debug.Print TypeName(fObj)
                If TypeName(fObj) = "TaskItem" Then 'Validation
                        
                    'For Checkmark use: ChrW(&H2713)
                    If fObj.Complete = True Then .Cells(i, 1).Value = fObj.DateCompleted
                    If fObj.StartDate < 200000 Then .Cells(i, 2).Value = fObj.StartDate
                    If fObj.DueDate < 200000 Then .Cells(i, 3).Value = fObj.DueDate
                    If fObj.ReminderSet = True Then .Cells(i, 4).Value = fObj.ReminderTime
                    .Cells(i, 5).Value = fObj.Subject
                    .Cells(i, 6).Value = fObj.Importance
                    .Cells(i, 7).Value = fObj.IsRecurring
                    .Cells(i, 8).Value = fObj.Status
                    '.Cells(i, 9).Value = fObj.Recipients
                    .Cells(i, 10).Value = fObj.Body
                        
                    Debug.Print TypeName(fObj) & "@" & i - 1 & " - " & IsEmpty(.Cells(i, 5).Value) & ": " & fObj.Subject
    
                    i = i + 1 'counter
                End If
            Next fObj
            
            'Formatting
            .Cells.RowHeight = 18.6
            '.Columns("A:J").EntireColumn.AutoFit
            .Range("A1:J1").Font.Bold = True
            
            With .Range("A1:J1").Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            
            Range("A1").Select
            
            'Debug
            Debug.Print "Cleaning up..." & Chr(10)
            
            'Null out the variables
            Set fItem = Nothing
            Set fMAPI = Nothing
            Set fObj = Nothing
        End With
            
    End With
    
    On Error Resume Next
    
    Set olNameSpace = Nothing
    
    Debug.Print "Done Contacts!" & Chr(10) & Chr(10)

End Sub

Public Sub GenInbox(wb As Workbook)

    Debug.Print "Starting Contacts..." & Chr(10)

    Dim fMAPI As MAPIFolder
    Dim fItem As Items 'Subject to change with class
    Dim fObj As Object
    Dim olNameSpace As NameSpace
    Dim i As Integer
    
    Set olNameSpace = Application.GetNamespace("MAPI")
    
    'Adds multiple stores to current namespace (PST Files)
    'olNameSpace.AddStore ("C:\Users\Owner\Documents\Outlook Files\stephen@honan.com.sg.pst")
    
    On Error GoTo 0

    'Avoids all the previous declarations from main class
    With wb
        
                
        ''''''''''''''''''''''''''''''''''''''''''''''
        '                   Inbox                    '
        ''''''''''''''''''''''''''''''''''''''''''''''
        
        'Search for default folder in current class
        Set fMAPI = olNameSpace.GetDefaultFolder(olFolderInbox)
        Set fItem = fMAPI.Items
    
        Debug.Print "Styling Worksheet..." & Chr(10)
    
        'Declare new worksheet
        .Sheets.Add().Name = "Inbox"
        
        With .Sheets("Inbox")
            .Activate
                
            'Link below for all possible fields
            'https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.taskitem_properties.aspx
            
            .Cells(1, 1).Value = "Date"
            .Cells(1, 2).Value = "Sender"
            .Cells(1, 3).Value = "Sender Email"
            .Cells(1, 4).Value = "Recipents"
            .Cells(1, 5).Value = "Unread?"
            .Cells(1, 6).Value = "Subject"
            .Cells(1, 7).Value = "isTask"
            .Cells(1, 8).Value = "TaskCompleted"
            .Cells(1, 9).Value = "Size"
            .Cells(1, 10).Value = "Saved"
            
            i = 2 'Start counter here
            
            For Each fObj In fItem
                'Debug.Print TypeName(fObj)
                If TypeName(fObj) = "MailItem" Then 'Validation
                        
                    'For Checkmark use: ChrW(&H2713)
                    .Cells(i, 1).Value = fObj.SentOn
                    .Cells(i, 2).Value = fObj.From
                    .Cells(i, 3).Value = fObj.To
                    .Cells(i, 4).Value = fObj.UnRead
                    .Cells(i, 5).Value = fObj.Subject
                    .Cells(i, 6).Value = fObj.IsMarkedAsTask
                    .Cells(i, 7).Value = fObj.TaskDueDate
                    .Cells(i, 8).Value = fObj.TaskCompletedDate
                    .Cells(i, 9).Value = fObj.Size
                    .Cells(i, 10).Value = fObj.Saved
                        
                    Debug.Print TypeName(fObj) & "@" & i - 1 & " - " & _
                    IsEmpty(.Cells(i, 5).Value) & ": " & fObj.Subject
    
                    i = i + 1 'counter
                End If
            Next fObj
            
            'Formatting
            .Cells.RowHeight = 18.6
            
            'Program will crash if this line is active
            '.Columns("A:J").EntireColumn.AutoFit
            .Range("A1:J1").Font.Bold = True
            
            With .Range("A1:J1").Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            
            Range("A1").Select
            
            'Debug
            Debug.Print "Cleaning up..." & Chr(10)
            
            'Null out the variables
            Set fItem = Nothing
            Set fMAPI = Nothing
            Set fObj = Nothing
        End With
            
    End With
    
    On Error Resume Next
    
    Set olNameSpace = Nothing
    
    Debug.Print "Done Contacts!" & Chr(10) & Chr(10)

End Sub

