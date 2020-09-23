VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FF8080&
   Caption         =   "Port Sweep"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Interval        =   2000
      Left            =   120
      Top             =   1080
   End
   Begin MSWinsockLib.Winsock wnsConnection 
      Left            =   600
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CmdAction 
      Caption         =   " "
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox txtMessage 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox StopIP4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "1"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtPort 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   10
      Text            =   "27374"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox StopIP1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "127"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox StopIP2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "0"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox StopIP3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "0"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox StartIP2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   2
      Text            =   " 0"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox StartIP3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox StartIP4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "1"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox StartIP1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "127"
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Port to Scan:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   1095
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Stop:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Start:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////
'/////////      Port Scanning Tool    //////////////
'///////////////////////////////////////////////////
'///  Created By : eMBit                         ///
'///  Created    : July 9, 2001                  ///
'///////////////////////////////////////////////////
' Created as a tool for portscanning a subnet for  '
' open ports, also created so that a newbie to VB  '
' can see the use of the winsock command.  The     '
' scanner will start a given IP and increment by 1 '
' until the stop IP is reached.                    '
'///////////////////////////////////////////////////
' OK so all works well... Now for some modifications
' I woke up this morning after letting it scan all
' night... The computer was turned off... hmmm...
' well that defeated the purpose, so I am adding
' some code to write to a file so as to not lose a
' scan again.  Also somewhat of a bother is clicking
' on an ip address string and having to delete whats
' already in the text box... so i will put in a got
' focus to select all text
'///////////////////////////////////////////////////
' As well... so that we dont scan the same computers
' 2 times we will write to the file the starting
' computer that we scanned, and when we click on
' stop or we end the scan we will write the last
' computer scanned
'///////////////////////////////////////////////////

'variable that we'll be using
Dim Action As Integer
Dim StartIP As String, EndIP As String
Dim Seconds

Private Sub CmdAction_Click()
'if command button already clicked then action would
'equal 1
If Action = 1 Then
    'change the command button to say start
    CmdAction.Caption = "Start"
    'close the winsock connection
    wnsConnection.Close
    'and reset the action variable to 0
    Action = 0
    Exit Sub
'else you are starting the scan and action equals 0
Else
    'set the action variable to 1
    Action = 1
    'set the command button to say stop
    CmdAction.Caption = "Stop"
    'call the routine to scan the ports
    Call ScanPorts
End If
End Sub

Private Sub Form_Load()
'initialize the action variable to 0
Action = 0
'set the command button to say start
CmdAction.Caption = "Start"
'we will lock the message box so that we dont write
'over the ip addresses accidently
txtMessage.Locked = True
End Sub

Private Sub ScanPorts()
'variables we'll use in this routine
Dim Start As Integer
'if we get an error that we dont like we'll just
'continue since we dont have any error handling
On Error Resume Next
'initialize the variables to 0
Start = 0
'concatenate the ip string to stop the scanner
EndIP = StopIP1 & "." & StopIP2 & "." & StopIP3 & "." & StopIP4
Do While Not StartIP = EndIP
    If Action = 0 Then
        Exit Do
    End If
    'concatenate the ip starting string
    StartIP = Trim(StartIP1.Text) & "." & Trim(StartIP2.Text) & "." & Trim(StartIP3.Text) & "." & Trim(StartIP4.Text)
    If Start = 0 Then
        Open App.Path & "\" & "Scan.txt" For Append As #1
        Write #1, "Scan started with "; StartIP
        Close #1
    End If
    'close the winsock control
    wnsConnection.Close
    'connect to the ip and port
    wnsConnection.Connect StartIP, txtPort
    'a little loop to wait to see if the connection
    'can be made
    Seconds = 0
    Do While Seconds = 0
        DoEvents
    Loop
    'if the computers are connected...
    If wnsConnection.State = 7 Then
        txtMessage.Text = txtMessage.Text & StartIP & vbNewLine
'///////////////////////////////////////////////////
        'this will open a file for output
        Open App.Path & "\" & "Scan.txt" For Append As #1
           'and will write the IP and the port number
           'being scanned
           Write #1, StartIP, txtPort.Text
        'and close the file
        Close #1
'///////////////////////////////////////////////////
    End If
    'step the ip address 1
    StartIP4.Text = StartIP4.Text + 1
        'if the ip address has reached the 255 limit then
        'count the next range 1 and reset to 1
        If StartIP4.Text = "256" Then
            StartIP4.Text = "1"
            StartIP3.Text = StartIP3.Text + 1
            If StartIP3.Text = "256" Then
                StartIP3.Text = "1"
                StartIP2.Text = StartIP2.Text + 1
                If StartIP2.Text = "256" Then
                    StartIP2.Text = "1"
                    StartIP1.Text = StartIP1.Text + 1
                    If StartIP1.Text = "256" Then
                        Exit Sub
                    End If
                End If
            End If
        End If
    'set the start variable to 1 so that when we loop
    'we dont write the starting ip again
    Start = 1
Loop
'then write to the file the last ip that was scanned
Open App.Path & "\" & "Scan.txt" For Append As #1
    Write #1, "The last ip scanned was "; StartIP
Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
'///////////////////////////////////////////////////
'ok so these next subroutines are for when a text
'box has the focus to select everything in the box
Private Sub StartIP1_GotFocus()
StartIP1.SelStart = 0
StartIP1.SelLength = Len(StartIP1.Text)
End Sub

Private Sub StartIP2_GotFocus()
StartIP2.SelStart = 0
StartIP2.SelLength = Len(StartIP1.Text)
End Sub

Private Sub StartIP3_GotFocus()
StartIP3.SelStart = 0
StartIP3.SelLength = Len(StartIP1.Text)
End Sub

Private Sub StartIP4_GotFocus()
StartIP4.SelStart = 0
StartIP4.SelLength = Len(StartIP1.Text)
End Sub

Private Sub StopIP1_gotfocus()
StopIP1.SelStart = 0
StopIP1.SelLength = Len(StopIP1.Text)
End Sub

Private Sub StopIP2_gotfocus()
StopIP2.SelStart = 0
StopIP2.SelLength = Len(StopIP1.Text)
End Sub

Private Sub StopIP3_gotfocus()
StopIP3.SelStart = 0
StopIP3.SelLength = Len(StopIP1.Text)
End Sub

Private Sub StopIP4_gotfocus()
StopIP4.SelStart = 0
StopIP4.SelLength = Len(StopIP1.Text)
End Sub

Private Sub txtPort_gotfocus()
txtPort.SelStart = 0
txtPort.SelLength = Len(txtPort.Text)
End Sub
'///////////////////////////////////////////////////

Private Sub Timer_Timer()
Seconds = 1
End Sub



