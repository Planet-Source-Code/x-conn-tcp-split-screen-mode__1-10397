VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6075
   ClientLeft      =   1980
   ClientTop       =   3210
   ClientWidth     =   6855
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   6855
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   25
      Top             =   120
      Width           =   285
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      BackColor       =   0
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Splash"
      TabPicture(0)   =   "frmMain.frx":0705
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Exchange"
      TabPicture(1)   =   "frmMain.frx":0721
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "RTFOut"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "RTFIn"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Setup"
      TabPicture(2)   =   "frmMain.frx":073D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame4"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame6"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Help"
      TabPicture(3)   =   "frmMain.frx":0759
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame7 
         Caption         =   "Help"
         Height          =   4455
         Left            =   240
         TabIndex        =   34
         Top             =   480
         Width           =   5895
         Begin VB.Label Label13 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmMain.frx":0775
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   735
            Left            =   120
            TabIndex        =   40
            Top             =   3120
            Width           =   5655
         End
         Begin VB.Label Label12 
            Caption         =   $"frmMain.frx":082F
            Height          =   615
            Left            =   120
            TabIndex        =   39
            Top             =   2280
            Width           =   5655
         End
         Begin VB.Label Label11 
            Caption         =   "If you're having problems using X-Conn, make sure that your Internet or LAN connection is working.  "
            Height          =   495
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   5655
         End
         Begin VB.Label Label10 
            Caption         =   $"frmMain.frx":0914
            Height          =   735
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   5655
         End
         Begin VB.Label Label7 
            Caption         =   $"frmMain.frx":09EA
            Height          =   495
            Left            =   120
            TabIndex        =   36
            Top             =   1680
            Width           =   5655
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "fordpref@home.com"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   3960
            Width           =   5655
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Stay On Top"
         Height          =   615
         Left            =   -74880
         TabIndex        =   31
         Top             =   1680
         Width           =   6135
         Begin VB.OptionButton Option6 
            Caption         =   "Do Not Stay On Top"
            Height          =   255
            Left            =   3960
            TabIndex        =   33
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Stay On Top Of All Other Programs  (Default)"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Popup Messaging"
         Height          =   735
         Left            =   -74880
         TabIndex        =   28
         Top             =   4320
         Width           =   6135
         Begin VB.TextBox txtPopup 
            BackColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   4575
         End
         Begin VB.CommandButton cmdPopup 
            Caption         =   "Send Popup"
            Height          =   375
            Left            =   4800
            TabIndex        =   29
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   4320
         Left            =   -74640
         Picture         =   "frmMain.frx":0A73
         ScaleHeight     =   4260
         ScaleWidth      =   5595
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "X-Conn TCP Split Screen - By FordPref Programming"
         Top             =   600
         Width           =   5655
      End
      Begin VB.Frame Frame4 
         Caption         =   "Connection Information"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   14
         Top             =   2520
         Width           =   6135
         Begin VB.TextBox txtPort 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   3375
         End
         Begin VB.TextBox txtRemoteIP 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtLocalIP 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Port To Listen/Connect"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3600
            TabIndex        =   20
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Remote IP Address"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3600
            TabIndex        =   19
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Local IP Address"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3600
            TabIndex        =   18
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Border"
         Height          =   1095
         Left            =   -73200
         TabIndex        =   11
         Top             =   480
         Width           =   1575
         Begin VB.OptionButton Option4 
            Caption         =   "No Border"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Show Border"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sounds"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   1575
         Begin VB.OptionButton Option1 
            Caption         =   "Play Sounds"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "No Sounds"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Winsock"
         Height          =   1095
         Left            =   -71520
         TabIndex        =   3
         Top             =   480
         Width           =   2775
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Session"
            Height          =   375
            Left            =   1440
            TabIndex        =   7
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdListen 
            Caption         =   "Listen"
            Height          =   375
            Left            =   1440
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdDisconnect 
            Caption         =   "Disconnect"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdConnect 
            Caption         =   "Connect"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
      End
      Begin RichTextLib.RichTextBox RTFIn 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3201
         _Version        =   393217
         BackColor       =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":156FB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTFOut 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2640
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2990
         _Version        =   393217
         BackColor       =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":157EF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "Local Text"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Remote Text"
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Written By FordPref Programming"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   -74880
         TabIndex        =   21
         Top             =   4440
         Width           =   6135
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "X-Conn TCP Split - Screen"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   720
      TabIndex        =   26
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Status Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   5760
      Width           =   6375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Leave these settings alone to play sounds within a VB program
Private Declare Function mciSendString Lib "winmm.dll" Alias _
        "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
        lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
        hwndCallback As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub cmdClear_Click()
    On Error GoTo err ' If there is an error, go to our error trap at the bottom of the subroutine
    mbox = MsgBox("Clear the current chat session?", vbOKCancel, "Clear Session?") ' Ask the user if they're sure they want to erase the incoming message box.
    If mbox = vbOK Then ' If the user click OK to clear the message box, then...
        RTFOut.Text = "" ' Set the incoming message box (RTFOut.Text) to "" (nothing).
        RTFOut.SetFocus ' Set the focus back on the outgoing message box to send another message to the other computer.
        Exit Sub ' Exit the subroutine
    End If

err: ' If the user pressed Cancel on the message box above, we end up here, since this produces an error in Visual Basic
    RTFOut.SetFocus ' The user pressed Cancel, so we do nothing but reset the focus back to the outgoing message box
    Exit Sub ' Exit the subroutine
End Sub

Private Sub cmdConnect_Click()
    On Error Resume Next ' If there's an error, resume the next command.
    Winsock1.Close ' Close any open ports (just in case).
    Winsock1.Connect txtRemoteIP.Text, txtPort.Text ' Try to connect to the computer IP address specified in the txtRemoteIP text box, on the port specified in the txtPort text box.
    lblStatus.Caption = "Connecting to " + txtRemoteIP.Text ' Inform the user we are trying to connect to the specified IP address.
    cmdDisconnect.Enabled = True ' Enable the Disconnect button since we may want to disconnect or stop trying to connect.
    cmdListen.Enabled = False ' Disable the Listen button and the Connect button since we are already trying to connect.
    cmdConnect.Enabled = False ' We are trying to connect, so hide the connect button.
    If err Then lblStatus.Caption = err.Description ' If there are any errors, inform the user by showing it on the lblStatus bar.
End Sub

Private Sub cmdDisconnect_Click()
    On Error Resume Next ' If there's an error, resume next command.
    Dim playsound As Long ' Declare the variable to hold the sound to be played if "Play Sounds" box is checked.
    If Option1.Value = True Then
        playsound = sndPlaySound("xcdiscon.wav", 1) ' If the "Play Sounds" box is checked, play the sound.
    End If
    Winsock1.Close ' We want to disconnect or stop listening for a connection request, so close the connected or listening port.
    cmdConnect.Enabled = True ' Enable the Connect button so we can connect to another computer.
    cmdListen.Enabled = True ' Enable the Listen button so we can listen for a connection request.
    cmdDisconnect.Enabled = False ' We are not connected to anything, so disable the Disconnect button.
    lblStatus.Caption = "Disconnected - Not Listening For Request." ' Show the user we are disconnected, and that we are not listening for a connection request.
End Sub

Private Sub cmdListen_Click()
    On Error Resume Next ' If there's an error, resume next command.
    cmdConnect.Enabled = False ' We are listening for a connection, so disable the Connect button.
    cmdListen.Enabled = False ' We are already listening for a connection, so disable the Listen button.
    cmdDisconnect.Enabled = True ' Enable the Disconnect button in case you want to stop listening for connection request.
    Winsock1.LocalPort = txtPort.Text ' Set the local port to listen on by getting the value from the txtPort text box.
    Winsock1.Listen ' Listen for the connection request by the other computer.
    lblStatus.Caption = "Listening For Connection Request" ' Inform the user that we are listening for a connection request.
End Sub

Public Sub stayontop(hwnd As Long, Stay As Boolean) ' This procedure/module keeps our form on top
       If Stay Then
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    Else
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    End If
End Sub

Private Sub cmdPopup_Click()
    On Error Resume Next ' If there's an error, continue with next command
    If txtPopup.Text = "" Then Exit Sub ' If the txtPopup text box is empty, don't send any data, exit subroutine.
    Winsock1.SendData ("|" & txtPopup.Text) ' Send the data with a pipe, |, as first character.  The pipe tells our program that this message is for a popup message box.  See DATA_ARRIVAL subroutine.
    txtPopup.Text = "" ' Set the txtPopup box to blank for another popup message.
End Sub

Private Sub Form_Load()
    On Error Resume Next ' Resume next command if there is an error.
    SSTab1.Tab = 0 ' This ensures that Tab 0 (splash screen tab) is shown first.
    Option1.Value = True ' Turns sounds on by default.
    Option4.Value = True ' Turns the program border off by default.
    Option5.Value = True ' Makes the program stay on top of all others by default.
    cmdDisconnect.Enabled = False
    txtLocalIP.Text = Winsock1.LocalIP ' This shows the local IP address in the txtlocalip box in Setup.
    RTFOut.SelColor = &H80FF80 ' Make sure our text is green on the outgoing message box.
    Winsock1.Close ' Make sure that Winsock1 (our connection port) is closed on startup - just to be sure.
    txtPort.Text = GetSetting("X-Conn Split-Screen", "Startup", "Port", 1989) ' Set the default port to 1989 if nothing is stored in the registry
    txtRemoteIP.Text = GetSetting("X-Conn Split-Screen", "Startup", "Remote", "") ' Clear the txtRemoteIP box if no remote IP is stored in the registry
    ' Display startup text on statusbar.
    lblStatus.Caption = "Status Bar - Watch Here For Important Information "
    If Option1.Value = True Then
        playsound = sndPlaySound("xcstartup.wav", 1) ' If the "Play Sounds" box is selected, play the sound.
    End If
End Sub

Private Sub Option3_Click()
    If Option3.Value = True Then Me.Caption = "X-Conn TCP Split-Screen" ' Turn on the form border (must be on to reposition the form)
End Sub

Private Sub Option4_Click()
    If Option4.Value = True Then Me.Caption = "" ' Turn off the form border
End Sub

Private Sub Option5_Click()
    If Option5.Value = True Then stayontop hwnd, True  ' This is calls the module that keeps our form on top.
End Sub

Private Sub Option6_Click()
    If Option6.Value = True Then stayontop hwnd, False ' Do not stay on top if the user has this selected.
End Sub

Private Sub Picture1_Click()
    SaveSetting "X-Conn Split-Screen", "Startup", "Port", txtPort.Text ' Saves the current port to connect and listen on.
    SaveSetting "X-Conn Split-Screen", "Startup", "Remote", txtRemoteIP.Text ' Saves the current computer's remote IP address for quick access next time you run the program.
    Winsock1.Close ' Just to make sure the connection or listening port is closed - for security reasons.
    End ' end everything associated with this program - do not use UNLOAD ME - use END!
End Sub

Private Sub RTFOut_KeyPress(KeyAscii As Integer)
    On Error GoTo err ' If there is an error in this subroutine, go to "err" code at bottom.
        Dim playsound As Long ' Declare the variable to hold the sound to be played if "Play Sounds" box is selected.
    If Option1.Value = True Then
        playsound = sndPlaySound("xctype.wav", 1) ' If the "Play Sounds" box is selected, play the sound.
    End If
    RTFOut.SelStart = Len(RTFOut.Text) ' Set cursor to end of outgoing message box. This keeps the last message on the screen.
    RTFOut.SelColor = &H80FF80 ' Make sure our text is green on the outgoing message box.
    Winsock1.SendData Chr(KeyAscii) ' Send each character (as it is typed to the other) computer.
    Exit Sub
err:
    lblStatus.Caption = err.Description ' Show the error to the user on the status bar.
    Resume Next ' Resume with next command after showing the error.
End Sub

Private Sub Winsock1_Close()
    Dim playsound As Long ' Declare the variable to hold the sound to be played if "Play Sounds" box is selected.
    If Option1.Value = True Then
        playsound = sndPlaySound("xcdiscon.wav", 1) ' If the "Play Sounds" box is selected, play the sound.
    End If
    lblStatus.Caption = "Connection Has Been Closed." ' Show the user that the connection is closed.
    cmdConnect.Enabled = True ' Reset the command buttons.
    cmdListen.Enabled = True  ' Connect and listen need to be enabled.
    cmdDisconnect.Enabled = False ' Disable Disconnect since we're not connected or listening for connection.
    cmdConnect.SetFocus ' Set the focus back to the cmdConnect button.
End Sub

Private Sub Winsock1_Connect()
    On Error Resume Next ' If there's an error, continue with next command.
    Dim playsound As Long ' Declare variable to hold the sound to be played if "Play Sounds" box is checked.
    lblStatus.Caption = "Connection Has Been Established!" ' Show the user we have a connection.
    txtRemoteIP.Text = Winsock1.RemoteHostIP ' Put the remote computer's IP in the remoteIP box.
    cmdConnect.Enabled = False ' Disable the Connect and Listen buttons.
    cmdListen.Enabled = False  ' We don't need these buttons enabled, and it prevents possible errors.
    cmdDisconnect.Enabled = True ' We are connected, so enable the Disconnect button.
    If Option1.Value = True Then
        playsound = sndPlaySound("xcestab.wav", 1) ' If the "Play Sounds" selected, play the default sound.
    End If
    RTFOut.SetFocus ' Set the focus on the box to enter messages to send to the other computer.
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next ' Just in case there's an error, continue with next command.
    Dim playsound As Long ' Declare variable to hold the sound to be played if "Play Sounds" box is checked.
    Winsock1.Close ' Close any open socket (just in case).
    Winsock1.Accept requestID ' Accept the other computer's connection request.
    lblStatus.Caption = "Connection Has Been Established!" ' Show the user we have accepted the connection request, and are connected.
    txtRemoteIP.Text = Winsock1.RemoteHostIP ' Show the remote computer's IP in the txtRemoteIP text box.
    cmdConnect.Enabled = False ' We are connected, so disable the Connect and Listen buttons.
    cmdListen.Enabled = False ' This helps to prevent anyone from clicking them and causing errors.
    cmdDisconnect.Enabled = True ' Enable the Disconnect button since we're connected.
    If Option1.Value = True Then
        playsound = sndPlaySound("xcestab.wav", 1) ' If the "Play Sounds" is selected, play the default sound.
    End If
    RTFOut.SetFocus ' Set the focus on the box to enter messages to send to the other computer.
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim currenttext As String ' String to hold contents of RTFIn if needed.
    Dim ndata As String ' Declare a variable to hold the incoming data.
    On Error Resume Next ' If there's an error, resume next command.
    Winsock1.GetData ndata ' Get the incoming data and store it in variable "ndata".
        If ndata = Chr(8) Then
            RTFIn.Text = Mid(RTFIn.Text, 1, Len(RTFIn.Text) - 1) ' Remove last character received if the BACKSPACE key was press.
            RTFIn.SelStart = Len(RTFIn.Text) ' Set the cursor to the end of the text box.
        Exit Sub
        End If
                If Mid(ndata, 1, 1) = "|" Then GoTo boxcode ' Check to see if this is a message box, if it is go to subroutine "boxcode".
    RTFIn.SelStart = Len(RTFIn.Text) ' Set the cursor to the end of the text box to hold the incoming messages.
    RTFIn.SelColor = &HC0E0FF ' Set the color of the incoming message to pale orange.
    RTFIn.SelText = Mid(ndata, InStr(1, ndata, "^") + 1) ' Insert the text to the last of the RTFIn message box.
    RTFIn.SelStart = Len(RTFIn.Text) ' Set the cursor to the end of the text box.
    RTFIn.SelColor = &H80FF80 ' Change the color of the letters back to pale green.
    Exit Sub ' Exit the subroutine.
boxcode: ' If the incoming data's first character was a pipe , |, then the program jumps here.
    MsgBox Mid(ndata, 2, Len(ndata) - 1), vbInformation, "Incoming Message" ' Display the incoming data as a message box.
    RTFOut.SetFocus ' Put the focus back on the RTFOut box to send another message.
    Exit Sub ' Exit the subroutine.
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lblStatus.Caption = Description ' If there was a winsock error, show the user.
    RTFOut.SetFocus ' Set the focus back on the message box to send another message.
End Sub
