Public Sub RemoveDuplicateMessages()
Const PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
Dim colItems As Items
Dim i As Integer
Dim j As Integer
Dim number_sequence As Integer
Dim strMsgID As String

'loop_flag = True
number_sequence = 1
Do While True
    Set colItems = ActiveExplorer.CurrentFolder.Items
    colItems.Sort "[RECEIVED]", True
    number_items = colItems.Count
    strMsgID = colItems(number_sequence).PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
    
    '
    '  Determine the end point of a search. It is assumed that the number of email duplication is less than 10.
    '
    range_search = 10
    If number_items < number_sequence + range_search Then
        number_end = number_items
    Else
        number_end = number_sequence + range_search
    End If
    
    '
    '  If Message ID is the same as that of the current email, remove the email.
    '
    Debug.Print number_sequence
    sequence_proceed_flag = True
    For i = number_sequence + 1 To number_end
        If colItems(i).PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID) = strMsgID Then
            If colItems(number_sequence).Body = colItems(i).Body Then
                colItems(i).Delete
                sequence_proceed_flag = False
                Debug.Print "Removed!!!", number_sequence
                Exit For
            End If
        End If
    Next
    
    '
    '  Once the sequence reaches the last email, quit this job.
    '
    If number_sequence = number_items Then
        Debug.Print "Finished!!!"
        Exit Do
    End If
    
    '
    '  If email is removed in this loop, leave this row position as is to find another duplication of email.
    '  If email is not removed in this loop, go to next step.
    '
    If sequence_proceed_flag = True Then
        number_sequence = number_sequence + 1
    End If
Loop

End Sub

