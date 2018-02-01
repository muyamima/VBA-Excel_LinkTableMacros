Attribute VB_Name = "Text"
Option Private Module

Function GetText(ByVal MsgNr As Integer, Optional ByVal Lang As String) As String 'Lang as ISO 639 double or triple letter code.
If Lang = "" Then
    Lang = GetLang()
Else
    Lang = LCase(Left(Trim(Lang), 3))
End If

Select Case Lang
    Case "de", "deu", "ger"
        Select Case MsgNr
            Case 1 'Select
                GetText = "Wählen"
            Case 2 'Select a file.
                GetText = "Datei auswählen."
            Case 3 'No file was selected.
                GetText = "Keine Datei ausgewählt."
            Case 20 'This item has no file extention.
                GetText = "Diese Datei hat keine Dateierweiterung."
            Case 50 'Testing link:
                GetText = "Prüft Verknüpfung:"
            Case 51 'Checking complete! & vbCrLf & Cells with broken or suspect links are highlighted in red.
                GetText = "Prüfen fertig!" & vbCrLf & "Zellen mit defekter oder verdächter Verknüpfungen werden in Rot angezeigt."
            Case 100 'Is the active sheet the sheet with Hyperlinks you would like to check?
                GetText = "Sind im aktiven Blatt Verknüpfungen enthalten, die Sie prüfen möchten?"
            Case 101 'File needs to be saved in order to check links.
                GetText = "Die Datei muss gespeichert werden damit die Verknüpfungen geprüft werden können."
            Case Else 'No message defined.
                GetText = "Keine Meldung definiert."
        End Select
    Case Else 'Default to English
        Select Case MsgNr
            Case 1 'Select
                GetText = "Select"
            Case 2 'Select a file.
                GetText = "Select a file."
            Case 3 'No file was selected.
                GetText = "No file was selected."
            Case 20 'This item has no file extention.
                GetText = "This item has no file extention."
            Case 50 'Testing link:
                GetText = "Testing link:"
            Case 51 'Checking complete! & vbCrLf & Cells with broken or suspect links are highlighted in red.
                GetText = "Checking complete!" & vbCrLf & "Cells with broken or suspect links are highlighted in red."
            Case 100 'Is the active sheet the sheet with Hyperlinks you would like to check?
                GetText = "Is the active sheet the sheet with Hyperlinks you would like to check?"
            Case 101 'File needs to be saved in order to check links.
                GetText = "File needs to be saved in order to check links."
            Case Else 'No message defined.
                GetText = "No message defined."
        End Select
    End Select
End Function

Private Function GetLang() As String
    Dim Lang As String: Lang = "de"
    
    GetLang = Lang
End Function

Private Sub TestText()
Dim MsgNr As Integer: MsgNr = 0
Dim Lang As String: Lang = ""

Lang = InputBox("Enter language code:" & vbCrLf & "(ISO 639 language codes, double or triple.)", "Test text messages.")
MsgNr = InputBox("Enter the message number to check:", "Test text messages.")

MsgBox "Message number " & MsgNr & " for " & Chr(34) & Lang & Chr(34) & " reads:" & vbCrLf & GetText(MsgNr, Lang), vbOKOnly, "Test text messages."
End Sub
