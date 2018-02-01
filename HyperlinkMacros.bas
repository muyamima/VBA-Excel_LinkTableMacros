Attribute VB_Name = "HyperlinkMacros"
Sub AuditHyperlinks()
    If MsgBox(Text.GetText(100), vbYesNo) = vbNo Then
        Exit Sub
    End If

    'Variable declaration.
    Dim WB As Workbook: Set WB = ThisWorkbook
    Dim WBPath As String: WBPath = CStr(WB.Path) & "\"
    Dim LinkPath As String: LinkPath = ""
    Dim TestOK As Boolean: TestOK = 0

    'Check save status of workbook. (Paths are only updated on save.)
    'If unsaved, save.
    If WB.Saved = False Then
        Dim FileSaveAs As Boolean: FileSaveAs = 0
        FileSaveAs = Application.Dialogs(xlDialogSaveAs).Show(WBPath)
        If Not FileSaveAs Then
                    MsgBox Text.GetText(101), vbInformation
            Exit Sub
        End If
    End If

    'Loop through each hyperlink in this workbook.
    On Error Resume Next
    For Each Link In Cells.Hyperlinks
        TestOK = 0
        LinkPath = Link.Address
        
        'Complete address for relative paths.
        If LCase(Left(LinkPath, 4)) <> "http" Then
            LinkPath = WBPath & LinkPath
        End If
      
        Application.StatusBar = Text.GetText(50) & LinkPath
    
        'Test links.
        If Left(LinkPath, 4) = "http" Then
            TestOK = TestWebAddress(LinkPath)
        Else
            TestOK = TestFilePath(LinkPath)
        End If
    
        'Item colouring according to testresult.
        If TestOK Then
            Link.Parent.Font.Color = RGB(0, 0, 255)
        Else
            Link.Parent.Font.Color = RGB(255, 0, 0)
        End If
    Next Link

    Application.StatusBar = False
    WB.Save
    MsgBox (Text.GetText(51))
End Sub

Function TestWebAddress(WebAddress As String) As Boolean
    'Variable declaration.
    Dim Website As New MSXML2.XMLHTTP: Set Website = Nothing

    Website.Open "HEAD", WebAddress, False
    Website.Send

    If Website.statustext = "OK" Then
        TestWebAddress = 1
    Else
        TestWebAddress = 0
    End If
End Function

Function TestFilePath(FilePath As String) As Boolean
    If Dir(FilePath) <> "" Then
        TestFilePath = 1
    Else
        TestFilePath = 0
    End If
End Function

Sub AddHyperlink()
    'Variable declaration.
    Dim WB As Workbook: Set WB = ThisWorkbook
    Dim WBPath As String: WBPath = CStr(WB.Path) & "\"
    Dim LinkPath, LinkName As String: LinkPath = "": LinkName = ""
    Dim PathEnd, FileNameEnd As Integer: PathEnd = 0: FileNameEnd = 0

    'Converts each text hyperlink selected into a working hyperlink.
    For Each xCell In Selection
        'Prompt filepicker if current cell is empty.
        If xCell.Formula = "" Then
            LinkPath = SelectFile
            'Catch invalid selection.
            If LinkPath = "" Then
                Exit Sub
            End If
        Else
            LinkPath = Trim(CStr(xCell.Formula))
        End If
        
        'Extract path and filename from complete path.
        PathEnd = InStrRev(LinkPath, "\", , 1)
        LinkName = Mid(LinkPath, PathEnd + 1)
        LinkPath = Left(LinkPath, PathEnd)
        
        'Account for Relative Path.
        If LinkPath = WBPath Then
            LinkPath = LinkName
        Else
            LinkPath = LinkPath & LinkName
        End If
                
        'Add http:// in front of webaddress.
        If LCase(Left(LinkPath, 4)) = "www." Then
            LinkPath = "http://" & LinkPath
            
        'Remove http/https from front of webaddressname.
        ElseIf LCase(Left(LinkPath, 4)) = "http" Then
            If LCase(Left(LinkPath, 7)) = "http://" Then
                LinkName = Mid(LinkPath, 8)
            End If
            If LCase(Left(LinkPath, 8)) = "https://" Then
                LinkName = Mid(LinkPath, 9)
            End If

        'Check for file extention.
        Else
            FileNameEnd = InStrRev(LinkName, ".", , 1) - 1
            If FileNameEnd < 1 Then
                MsgBox Text.GetText(20), vbCritical
                Exit Sub
            End If
'            LinkName = Mid(LinkName, 1, FileNameEnd)
        End If
        
    'Create hyperlink.
    ActiveSheet.Hyperlinks.Add Anchor:=xCell, Address:=LinkPath, ScreenTip:=LinkPath, TextToDisplay:=LinkName
    Next xCell
    
    'Save workbook. (Paths are only updated on save.)
    WB.Save    
End Sub

Function SelectFile() As String 'Open explorer to make file selection.
    'Variable declaration.
    Dim WB As Workbook: Set WB = ThisWorkbook
    Dim WBPath As String: WBPath = CStr(WB.Path) & "\"
    
    'File Selection Dialog Box.
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = 0   'Only one file at a time!
        .ButtonName = Text.GetText(1)
        .InitialFileName = WBPath   'Start in workbook folder.
        .Title = Text.GetText(2)
        .Show
        If .SelectedItems.Count > 0 Then    'Return complete filepath if a file was selected.
            SelectFile = .SelectedItems(1)
        Else
            MsgBox Text.GetText(3), vbCritical
        End If
    End With
End Function

Private Sub ReplacePath()   'Fix relative links broken by opening/saving in the wrong place.
    'Variable declaration.
    Dim WB As Workbook: Set WB = ThisWorkbook
    Dim WBPath As String: WBPath = CStr(WB.Path) & "\"
    Dim WrongPath As String: WrongPath = "C:\" 'Part of the path that is wrong.
    Dim LinkPath As String: LinkPath = ""
    Dim FileNameEnd As Integer: FileNameEnd = 0
    
   'Loop through each hyperlink in this workbook.
    For Each Link In Cells.Hyperlinks
        LinkPath = Link.Address
        
        'Extract path and filename from complete path.
        PathEnd = InStrRev(LinkPath, "\", , 1)
        LinkName = Mid(LinkPath, PathEnd + 1)
        LinkPath = Left(LinkPath, PathEnd)
        
        'Check for links where the path is wrong.
        If LinkPath = WrongPath Then
            LinkPath = LinkName
        Else
            LinkPath = LinkPath & LinkName
        End If
        
        'Check for file extention.
        FileNameEnd = InStrRev(LinkName, ".", , 1) - 1
            If FileNameEnd < 1 Then
                        MsgBox Text.GetText(20), vbCritical
                Exit Sub
            End If
'        LinkName = Mid(LinkName, 1, FileNameEnd)

        'Update hyperlink.
        With Link
            .Address = LinkPath
            .ScreenTip = LinkPath
            .TextToDisplay = LinkName
        End With
    Next Link
End Sub
