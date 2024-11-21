'================================================================================
' Module: RemoveDuplicateEmails
' Description: Removes duplicate emails from the current Outlook folder.
' Author: [Rahul Kumar Singh]
' Date: [21 Nov 2024]
' License: MIT
'================================================================================

Attribute VB_Name = "DeleteDuplicateEmails"
Sub RemoveDuplicateEmailsSafely()
    Dim OutApp As Outlook.Application
    Dim MailFolder As Outlook.folder
    Dim Item As Object
    Dim MailItem As Outlook.MailItem
    Dim UniqueItems As Object
    Dim EmailKey As String
    Dim i As Long
    Dim ReceivedTimeWithoutSeconds As String
    Dim TotalItems As Long
    Dim CurrentItem As Long
    Dim DeletedCount As Long ' Counter for deleted emails

    ' Initialize Outlook application and folder
    Set OutApp = Application
    Set MailFolder = OutApp.ActiveExplorer.CurrentFolder

    ' Create a Dictionary to store unique email keys
    Set UniqueItems = CreateObject("Scripting.Dictionary")

    ' Get total items in the folder for progress calculation
    TotalItems = MailFolder.Items.Count
    DeletedCount = 0 ' Initialize the deletion counter

    ' Show progress form
    Load ProgressForm
    ProgressForm.Show vbModeless

    ' Loop through items in reverse order to avoid skipping items
    For i = TotalItems To 1 Step -1
        Set Item = MailFolder.Items(i)

        ' Update progress
        CurrentItem = TotalItems - i + 1
        ProgressForm.UpdateProgress CurrentItem, TotalItems

        ' Process only if the item is a MailItem
        If Item.Class = olMail Then
            Set MailItem = Item

            ' Format ReceivedTime to exclude seconds
            ReceivedTimeWithoutSeconds = Format(MailItem.ReceivedTime, "yyyy-mm-dd hh:nn")

            ' Create a unique key for each email based on Subject, Sender, ReceivedTime (without seconds), and Body
            EmailKey = MailItem.Subject & "|" & MailItem.SenderEmailAddress & "|" & _
                       ReceivedTimeWithoutSeconds & "|" & MailItem.Body

            ' Check if the email key already exists in the dictionary
            If UniqueItems.Exists(EmailKey) Then
                ' Duplicate found, delete the email
                MailItem.Delete
                DeletedCount = DeletedCount + 1 ' Increment deletion counter
            Else
                ' Add the unique email key to the dictionary to preserve the first occurrence
                UniqueItems.Add EmailKey, True
            End If
        End If

        ' Yield control back to Outlook every 10 iterations
        If CurrentItem Mod 10 = 0 Then
            DoEvents
        End If
    Next i

    ' Unload progress form
    Unload ProgressForm

    ' Display the deletion count
    MsgBox DeletedCount & " duplicate emails were removed!", vbInformation

    ' Cleanup
    Set OutApp = Nothing
    Set MailFolder = Nothing
    Set MailItem = Nothing
    Set UniqueItems = Nothing
End Sub


