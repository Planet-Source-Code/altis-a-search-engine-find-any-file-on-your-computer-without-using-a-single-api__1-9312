Attribute VB_Name = "SearchEngine"
Option Explicit

Private Searching As Boolean


Public Function FileLength(FileNamePath As String) As String
    Dim nFileNum As Integer
    
    Dim TempNum As Single

    ' Get file's length
    FileLength = FileLen(FileNamePath) / 1024
    
    ' Round number
    TempNum = FileLength - Int(FileLength)
    
    ' Use different scale according to the size of the file
    Select Case Val(FileLength)
        Case Is > 1024 ' Use Mega Byte
            FileLength = Int(FileLength / 1000) & "MB"
        Case Else  ' Use Kilo Byte
            ' All values are to round up
            FileLength = Format(FileLength + (1 - TempNum), "@@KB")
    End Select

' ********************************************************************************
' ********************************************************************************
' All files that are found must pass this section.  Therefore, it is a good idea
' to place codes here
' ********************************************************************************
' ********************************************************************************
End Function

Public Function FileDate(FileNamePath As String) As String
    On Error Resume Next
    ' Get file's date and time
    FileDate = FileDateTime(FileNamePath)
End Function

Private Sub StartSearch(FileToSearch As String, SearchDirectory As String, _
    CurrentForm As Form, ListDisplay As ListView, _
    Optional LabelCurrentDirectory As Label, _
    Optional StatusBarCurrentDirectory As StatusBar)
    
    Dim CurrentCharacter As Integer ' A pointer of a character in a String
    
    Dim CurrentDirectory As Integer ' The Index of the CurrentDirectory
    
    Dim Directories() As String ' Holds the name of the directory
    ReDim Directories(0)
    
    Dim DirectoryCount As Integer   ' Number of dierctory in a directory
    
    Dim DirectoryName As String     ' Holds the name of the directory

    Dim FileName As String  ' Holds the name of the file
    
    Dim ListViewItemX As ListItem
    
    Dim ListViewItemCount As Integer
    
    
    ' Stop the search if Searching suddenly turn to False
    If Searching = False Then Exit Sub
    
    
    
    On Error Resume Next    ' Error Handling
    
    
    
    ' Set the second column's alighment to right justify
    ListDisplay.ColumnHeaders(3).Alignment = lvwColumnRight
    
    ' If the user didn't enter anythingn in SearchDirectory, assume it is "C:\"
    If Len(SearchDirectory) <= 0 Then SearchDirectory = "C:\"
    
    ' Make sure the last character of SearchDirectory is not "\"
    CurrentCharacter = Len(SearchDirectory)
    Do Until Mid(SearchDirectory, CurrentCharacter, 1) <> "\"
        ' Remove last character
        SearchDirectory = Left(SearchDirectory, CurrentCharacter - 1)
        ' Continue to go backward
        CurrentCharacter = CurrentCharacter - 1
    Loop
    
    ' If the user didn't enter the extension of the file, place while card
    ' character at the end of the String
    CurrentCharacter = Len(FileToSearch)
    ' Go into the loop ONLY when user enter something
    If Len(FileToSearch) > 0 Then
        Do Until Mid(FileToSearch, CurrentCharacter, 1) = "." Or _
            CurrentCharacter <= 1
            CurrentCharacter = CurrentCharacter - 1
        Loop
    End If
    ' Check teh value Current Charcater.  If it is 1, that's mean the user
    ' didn't enter any extension
    If CurrentCharacter <= 1 Then
        ' Insert wild card character at the end of the string
        FileToSearch = FileToSearch & "*"
    End If

' Find all directories first
    ' Retrieve the first entry.
    DirectoryName = Dir(SearchDirectory & "\", vbDirectory)
    
    ' Start the loop
    Do While DirectoryName <> ""
        
        ' Ignore the current directory and the encompassing directory.
        If DirectoryName <> "." And DirectoryName <> ".." Then
            ' Use bitwise comparison to make sure Directory is a directory.
            Select Case GetAttr(SearchDirectory & "\" & DirectoryName)
                Case vbDirectory To vbDirectory + vbSystem
                    ' Save Directory
                    ReDim Preserve Directories(DirectoryCount + 1)
                    Directories(DirectoryCount) = DirectoryName
                    DirectoryCount = DirectoryCount + 1
            End Select
        End If

        ' Get next entry
        DirectoryName = Dir
    Loop
    
' Find file(s)
    ' Retrieve the first entry.
    DirectoryName = Dir(SearchDirectory & "\", vbDirectory)
    
    ' Start the loop
    Do While DirectoryName <> ""
        
        ' Ignore the current directory and the encompassing directory.
        If DirectoryName <> "." And DirectoryName <> ".." Then
            ' Use bitwise comparison to make sure Directory is a directory.
            Select Case GetAttr(SearchDirectory & "\" & DirectoryName)
                Case vbDirectory To vbDirectory + vbSystem
                    ' Search all contents in this direcotry
                    StartSearch FileToSearch, SearchDirectory & "\" & DirectoryName, _
                        CurrentForm, ListDisplay, _
                        LabelCurrentDirectory, StatusBarCurrentDirectory
            End Select
        End If

        ' Get next entry
        DirectoryName = Dir(SearchDirectory & "\", vbDirectory)
        
        Do
            ' Loop until we find the next directory name
            DirectoryName = Dir
        Loop Until DirectoryName = Directories(CurrentDirectory) Or _
            Directories(CurrentDirectory) = "" Or _
            DirectoryName = ""
            
        ' Show current directory
        LabelCurrentDirectory.Caption = SearchDirectory & "\" & Directories(CurrentDirectory)
        StatusBarCurrentDirectory.Panels(1).Text = _
            SearchDirectory & "\" & Directories(CurrentDirectory)
        
        If Directories(CurrentDirectory) = "" Then Exit Do
        
        ' Change pointer
        CurrentDirectory = CurrentDirectory + 1
        
        DoEvents
    
    
        ' Stop the search if Searching suddenly turn to False
        If Searching = False Then Exit Sub
    
    
    Loop
    
' ********************************************************************************
' ********************************************************************************
' NOTE:  TO MAKE THIS CODE SEARCH THE DIRECTORY ALSO, REPLACE:
'           FileName = Dir(SearchDirectory & "\" & FileToSearch)
'        WITH:
'           FileName = Dir(SearchDirectory & "\" & FileToSearch, vbDirectory)
' ********************************************************************************
' ********************************************************************************
' Find file now
    ' Retrieve the first entry
    FileName = Dir(SearchDirectory & "\" & FileToSearch)
    
    ' Start the loop
    Do While FileName <> ""
        ' Ingore the current directory and the encompassing directory.
        If FileName <> "." And FileName <> ".." Then
            ' Put this filename into the list view
            Set ListViewItemX = ListDisplay.ListItems.Add(, , FileName)
            ListViewItemCount = ListViewItemCount + 1
            ListViewItemX.Tag = ListViewItemCount
            
            ' Set sub item
            ' Show File's Directory
            ListViewItemX.SubItems(1) = SearchDirectory
                            
            ' Show File's Size
            ListViewItemX.SubItems(2) = _
                FileLength(SearchDirectory & "\" & FileName)
            ' Show last access date and time
            ListViewItemX.SubItems(3) = FileDate(SearchDirectory & "\" & FileName)
        End If
        
        ' Get next entry
        FileName = Dir
    Loop
End Sub

Public Function FindFile(FileToSearch As String, SearchDirectory As String, _
    CurrentForm As Form, ListDisplay As ListView, _
    Optional LabelCurrentDirectory As Label, _
    Optional StatusBarCurrentDirectory As StatusBar) As String
    
    ' Check current status
    If Searching = True Then
        ' Reset value
        Searching = False
        
        Exit Function
    Else
        ' Set value to True
        Searching = True
    End If
        
    ' Start the search
    StartSearch FileToSearch, SearchDirectory, CurrentForm, ListDisplay, _
        LabelCurrentDirectory, StatusBarCurrentDirectory
    
    ' Reset value
    Searching = False
    
    ' Display results
    FindFile = (ListDisplay.ListItems.Count) & " item(s) found."
End Function

