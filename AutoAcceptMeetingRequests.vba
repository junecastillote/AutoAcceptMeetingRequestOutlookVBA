' Auto accept meeting requests
Sub AutoAcceptMeetingRequestsByOrganizer(Item As Outlook.MeetingItem)
    ' Declare variables
    Dim objAppointment As Outlook.AppointmentItem
    Dim objMeeting As Outlook.MeetingItem
    Dim organizer As Outlook.AddressEntry

    Debug.Print "-> Script: AutoAcceptRequestsByOrganizer"

    ' Get the associated appointment item For the meeting request
    Set objAppointment = Item.GetAssociatedAppointment(True)

    ' Check If the appointment item exists
    If Not objAppointment Is Nothing Then
        Set organizer = objAppointment.GetOrganizer
        If objAppointment.MeetingStatus <> 7 And objAppointment.MeetingStatus <> 5 And objAppointment.ResponseStatus = 5 Then
            ' Check If the organizer is an internal user
            If organizer.Type = "EX" Then
                Debug.Print "--> The meeting request [Subject: " & objAppointment.Subject & "] organizer" & vbCrLf & "    [Organizer: " & objAppointment.GetOrganizer.GetExchangeUser.PrimarySmtpAddress & "] matched the rule and will be automatically accepted."
            End If

            ' Check If the organizer is an external user
            If organizer.Type = "SMTP" Then
                Debug.Print "--> The meeting request [Subject: " & objAppointment.Subject & "] organizer" & vbCrLf & "    [Organizer: " & objAppointment.GetOrganizer.Address & "] matched the rule and will be automatically accepted."
            End If

            ' Accept the meeting request And send the response
            Set objMeeting = objAppointment.Respond(olMeetingAccepted, True)
            objMeeting.Send
        Else
            Debug.Print "--> No action."
        End If
    Else
        Debug.Print "--> Appointment does Not exist."
    End If

    ' Clean up objects
    Set objAppointment = Nothing
    Set objMeeting = Nothing
    Set organizer = Nothing
End Sub

' Auto accept all meeting requests from internal organizers
Sub AutoAcceptInternalMeetingRequests(Item As Outlook.MeetingItem)
    ' Declare variables
    Dim objAppointment As Outlook.AppointmentItem
    Dim objMeeting As Outlook.MeetingItem
    Dim organizer As Outlook.AddressEntry

    Debug.Print "-> Script: AutoAcceptInternalMeetingRequests"

    ' Get the associated appointment item For the meeting request
    Set objAppointment = Item.GetAssociatedAppointment(True)

    ' Check If the appointment item exists
    If Not objAppointment Is Nothing Then
        Set organizer = objAppointment.GetOrganizer
        ' Check If the organizer is an internal user
        If organizer.Type = "EX" And objAppointment.MeetingStatus <> 7 And objAppointment.MeetingStatus <> 5 And objAppointment.ResponseStatus = 5 Then
            Debug.Print "--> The meeting request [Subject: " & objAppointment.Subject & "] organizer " & vbCrLf & "    [Organizer: " & objAppointment.GetOrganizer.GetExchangeUser.PrimarySmtpAddress & "] is internal and will be automatically accepted."
            ' Accept the meeting request And send the response
            Set objMeeting = objAppointment.Respond(olMeetingAccepted, True)
            objMeeting.Send
        Else
            Debug.Print "--> No action."
        End If

    Else
        Debug.Print "--> Appointment does Not exist."
    End If

    ' Clean up objects
    Set objAppointment = Nothing
    Set objMeeting = Nothing
    Set organizer = Nothing
End Sub

' Auto accept meeting requests from external organizers If there's no conflict
Sub AutoAcceptExternalMeetingRequestsIfNoConflict(Item As Outlook.MeetingItem)
    ' Declare variables
    Dim objAppointment As Outlook.AppointmentItem
    Dim organizer As Outlook.AddressEntry
    Dim calendarFolder As Outlook.Folder
    Dim calendarItems As Outlook.Items
    Dim calendarItem As Outlook.AppointmentItem
    Dim objMeeting As Outlook.MeetingItem

    Debug.Print "-> Script: AutoAcceptExternalMeetingRequestsIfNoConflict"

    ' Get the associated appointment item For the meeting request
    Set objAppointment = Item.GetAssociatedAppointment(True)

    ' Check If the appointment item exists
    If Not objAppointment Is Nothing Then
        ' Get the organizer's information
        Set organizer = objAppointment.GetOrganizer

        ' Check If the organizer is an external sender And that the meeting request status is Not canceled.
        If organizer.Type = "SMTP" And objAppointment.MeetingStatus <> 7 And objAppointment.MeetingStatus <> 5 And objAppointment.ResponseStatus = 5 Then
            ' Get the default calendar folder
            Set calendarFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar)

            ' Filter calendar items For the same period As the New meeting request
            strFilter = ""
            strFilter = strFilter & "[Start] >= '" & Format(objAppointment.Start, "ddddd h:nn AMPM") & "'"
            strFilter = strFilter & " And [End] <= '" & Format(objAppointment.End, "ddddd h:nn AMPM") & "'"
            strFilter = strFilter & " And [MeetingStatus] <> 7" 'olMeetingReceivedAndCanceled
            strFilter = strFilter & " And [MeetingStatus] <> 5" 'olMeetingCanceled
            strFilter = strFilter & " And ([ResponseStatus] = 3" 'olResponseAccepted
            strFilter = strFilter & " Or [ResponseStatus] = 0)" 'olNonMeeting
            Set calendarItems = calendarFolder.Items.Restrict(strFilter)

            ' If the filtered calendar items count is Not 0, it means there's schedule conflict.
            If calendarItems.Count > 0 Then
                Debug.Print "--> The meeting request [Subject: " & objAppointment.Subject & "] from " & vbCrLf & "    [Organizer: " & objAppointment.GetOrganizer.Address & "] conflicts with the following appointment(s):"
                For Each calendarItem In calendarItems
                    If calendarItem.GetOrganizer.Type = "EX" Then
                        conflictOrganizerEmail = calendarItem.GetOrganizer.GetExchangeUser.PrimarySmtpAddress
                    ElseIf calendarItem.GetOrganizer.Type = "SMTP" Then
                        conflictOrganizerEmail = calendarItem.Address
                    End If
                    Debug.Print ""
                    Debug.Print "      Subject: " & calendarItem.Subject
                    Debug.Print "      Organizer: " & conflictOrganizerEmail
                    Debug.Print "      Start: " & calendarItem.Start
                    Debug.Print "      End: " & calendarItem.End
                    Debug.Print ""
                Next calendarItem
            End If

            ' If the filtered calendar items count is 0, it means there's no schedule conflict And the meeting request will be accepted
            If calendarItems.Count < 1 Then
                Debug.Print "--> The meeting request [Subject: " & objAppointment.Subject & "] from " & vbCrLf & "[    Organizer: " & objAppointment.GetOrganizer.Address & "] has no conflict and will be automatically accepted."
                Set objMeeting = objAppointment.Respond(olMeetingAccepted, True)
                objMeeting.Send
            End If

        Else
            Debug.Print "--> No action."
        End If

    Else
        Debug.Print "--> Appointment does Not exist."
    End If

    ' Clean up objects
    Set objAppointment = Nothing
    Set objMeeting = Nothing
    Set organizer = Nothing
    Set calendarFolder = Nothing
    Set calendarItems = Nothing
    Set calendarItem = Nothing
End Sub