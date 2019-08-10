Attribute VB_Name = "EmailSendoutModule"
Option Explicit
Public Sub EmailSendout()
'   This procedure will generate one new Outlook mail item
'   for each valid email address in Column B, populate the
'   From:, To:, and Cc: fields, then add any attachments.

'   Reference Needed: Microsoft Outlook 16.0 Object Library

    Dim OutApp As Outlook.Application
    Dim OutMail As Outlook.MailItem
    Dim Cell As Range
    Dim EmailAddr1 As String
    Dim EmailAddr2 As String
    Dim Attach1 As String
    Dim Attach2 As String
    
'   Userform support variables
    Dim From As String
    Dim TemplatePath As String
    Dim lrow As Long
    
    lrow = Cells(Rows.Count, "A").End(xlUp).Row
    ' Debug.Print lrow
    
    TemplatePath = GenerateDraftsForm.TemplateBox
    From = GenerateDraftsForm.FromBox

    Application.ScreenUpdating = False
    
'   Create Outlook object
    Set OutApp = New Outlook.Application

'   Loop through the rows
    ' For Each Cell In Columns("C").Cells.SpecialCells(xlCellTypeVisible)
    For Each Cell In Range("C2:C" & lrow).SpecialCells(xlCellTypeVisible)

'       Check that each cell in column E is populated with a valid email address
        If Cell.Value Like "*@*.??" Or Cell.Value Like "*@*.???" Then

'           Get the data
            EmailAddr1 = Cell.Value
            EmailAddr2 = Cell.Offset(0, 1).Value
            Attach1 = Cell.Offset(0, 2).Value
            Attach2 = Cell.Offset(0, 3).Value

'           Create new Mail Item from template, save as draft, and close before sending
            Set OutMail = OutApp.CreateItemFromTemplate(TemplatePath)
                With OutMail
                    .SentOnBehalfOfName = From
                    .To = EmailAddr1
                    .CC = EmailAddr2
                    ' .DeferredDeliveryTime = DateAdd("n", 5, Now())
                    .Display
                    
                    
'                   Use radio buttons to determine how many attachments to add
                    If GenerateDraftsForm.AttachButton1 = True Then
                        .Attachments.Add (Attach1)
                        Else
                        If GenerateDraftsForm.AttachButton2 = True Then
                            .Attachments.Add (Attach1)
                            .Attachments.Add (Attach2)
                            Else
                        End If
                    End If
                    
                    .Save
                    .Close olSave
                    ' .Send
                End With
        End If
    Next
    
    Application.ScreenUpdating = True
    
    ' MsgBox "Mamma Mia!"
    
End Sub
Public Sub ShowForm()

' Sets the template and from email address to the last saved settings then calls the user form

GenerateDraftsForm.TemplateBox = Worksheets("Saved").Range("A2")
GenerateDraftsForm.FromBox = Worksheets("Saved").Range("B2")
GenerateDraftsForm.AttachButton1 = Worksheets("Saved").Range("C2")
GenerateDraftsForm.AttachButton2 = Worksheets("Saved").Range("D2")
GenerateDraftsForm.AttachButtonNone = Worksheets("Saved").Range("E2")
' GenerateDraftsForm.TimePicker = Format(DateAdd("n", 5, Now()), "hh:mm")
GenerateDraftsForm.Show

End Sub

Public Sub EmailSchedule()
'   This procedure will generate one new Outlook mail item
'   for each valid email address in Column B, populate the
'   From:, To:, and Cc: fields, then add any attachments.

'   Reference Needed: Microsoft Outlook 16.0 Object Library

    Dim OutApp As Outlook.Application
    Dim OutMail As Outlook.MailItem
    Dim Cell As Range
    Dim EmailAddr1 As String
    Dim EmailAddr2 As String
    Dim Attach1 As String
    Dim Attach2 As String
    
'   Userform support variables
    Dim From As String
    Dim TemplatePath As String
    Dim lrow As Long
    
    lrow = Cells(Rows.Count, "A").End(xlUp).Row
    ' Debug.Print lrow
    
    TemplatePath = GenerateDraftsForm.TemplateBox
    From = GenerateDraftsForm.FromBox

    Application.ScreenUpdating = False
    
'   Create Outlook object
    Set OutApp = New Outlook.Application

'   Loop through the rows
    ' For Each Cell In Columns("C").Cells.SpecialCells(xlCellTypeVisible)
    For Each Cell In Range("C2:C" & lrow).SpecialCells(xlCellTypeVisible)

'       Check that each cell in column E is populated with a valid email address
        If Cell.Value Like "*@*.??" Or Cell.Value Like "*@*.???" Then

'           Get the data
            EmailAddr1 = Cell.Value
            EmailAddr2 = Cell.Offset(0, 1).Value
            Attach1 = Cell.Offset(0, 2).Value
            Attach2 = Cell.Offset(0, 3).Value

'           Create new Mail Item from template, save as draft, and close before sending
            Set OutMail = OutApp.CreateItemFromTemplate(TemplatePath)
                With OutMail
                    .SentOnBehalfOfName = From
                    .To = EmailAddr1
                    .CC = EmailAddr2
                    .DeferredDeliveryTime = SchedulerForm.MonthPicker + SchedulerForm.TimePicker
                    .Display
                    
'                   Use radio buttons to determine how many attachments to add
                    If GenerateDraftsForm.AttachButton1 = True Then
                        .Attachments.Add (Attach1)
                        Else
                        If GenerateDraftsForm.AttachButton2 = True Then
                            .Attachments.Add (Attach1)
                            .Attachments.Add (Attach2)
                            Else
                        End If
                    End If
                    
                    .Save
                    ' .Close olSave
                    .Send
                End With
        End If
    Next
    
    Application.ScreenUpdating = True
    
    ' MsgBox "Mamma Mia!"
    
End Sub
Public Sub TestSendout()
'   This procedure will generate one new Outlook mail item
'   for each valid email address in Column B, populate the
'   From:, To:, and Cc: fields, then add any attachments.

'   Reference Needed: Microsoft Outlook 16.0 Object Library

    Dim OutApp As Outlook.Application
    Dim OutMail As Outlook.MailItem
    Dim TestEmailAddr As String
    
    TestEmailAddr = InputBox("Enter the email address for testing", "Test Email Sendout", "gmail.com")

'   Userform support variables
    Dim From As String
    Dim TemplatePath As String

    TemplatePath = GenerateDraftsForm.TemplateBox
    From = GenerateDraftsForm.FromBox

    Application.ScreenUpdating = False
    
'   Create Outlook object
    Set OutApp = New Outlook.Application

'   Create new Mail Item from template, save as draft, and close before sending
        Set OutMail = OutApp.CreateItemFromTemplate(TemplatePath)
            With OutMail
                .SentOnBehalfOfName = From
                .To = TestEmailAddr

                ' .DeferredDeliveryTime = DateAdd("n", 5, Now())
                .Display
                                    
                .Save
                .Close olSave
                ' .Send
            End With

    Application.ScreenUpdating = True
    
End Sub
