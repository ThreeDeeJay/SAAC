Attribute VB_Name = "ShellMod"
Const STARTF_USESHOWWINDOW = &H1&
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long




      Function StartDoc(DocName As String) As Long
          Dim Scr_hDC As Long
          Scr_hDC = GetDesktopWindow()
          StartDoc = ShellExecute(Scr_hDC, "Open", DocName, _
          "", "", STARTF_USESHOWWINDOW)
      End Function





