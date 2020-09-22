Attribute VB_Name = "stayontop"
' This is the code line to use to keep our form
' on top of everything else:
' StayOnTop hwnd, True
' Put that line in the Form_Load Sub
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2

 Public Declare Function SetWindowPos Lib _
    "user32" (ByVal hwnd As Long, ByVal _
    hWndInsertAfter As Long, ByVal x As Long, _
    ByVal y As Long, ByVal cx As Long, ByVal _
    cy As Long, ByVal wFlags As Long) As Long

