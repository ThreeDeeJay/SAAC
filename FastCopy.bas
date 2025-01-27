Attribute VB_Name = "FastCopy"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) _
    As Long
   
Const CB_RESETCONTENT = &H14B
Const CB_GETCOUNT = &H146
Const CB_GETITEMDATA = &H150
Const CB_SETITEMDATA = &H151
Const CB_GETLBTEXT = &H148
Const CB_ADDSTRING = &H143

' Duplicate the contents of a ComboBox control to another ComboBox control
'
' Pass False to the last argument to append contents
' (the target control isn't cleared before adding elements)
'
' uses API functions for the fastest code

Sub DuplicateComboBox(Source As ComboBox, Target As ComboBox, _
    Optional AppendMode As Boolean)
    Dim index As Long
    Dim itmData As Long
    Dim numItems As Long
    Dim sItemText As String
    
    ' prepare the receiving buffer
    sItemText = Space$(512)
    
    ' temporarily prevent updating
    LockWindowUpdate Target.hWnd
    
    ' reset target contents, if not in append mode
    If Not AppendMode Then
        SendMessage Target.hWnd, CB_RESETCONTENT, 0, ByVal 0&
    End If
    
    ' get the number of items in the source control
    numItems = SendMessage(Source.hWnd, CB_GETCOUNT, 0&, ByVal 0&)
    
    For index = 0 To numItems - 1
        ' get the item text
        SendMessage Source.hWnd, CB_GETLBTEXT, index, ByVal sItemText
        ' get the item data
        itmData = SendMessage(Source.hWnd, CB_GETITEMDATA, index, ByVal 0&)
        ' add the item text to the target list
        SendMessage Target.hWnd, CB_ADDSTRING, 0&, ByVal sItemText
        ' add the item data to the target list
        SendMessage Target.hWnd, CB_SETITEMDATA, index, ByVal itmData
    Next
    
    ' allow redrawing
    LockWindowUpdate 0
    
End Sub
