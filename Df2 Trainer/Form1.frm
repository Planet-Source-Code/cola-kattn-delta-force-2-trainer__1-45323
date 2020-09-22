VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "DF2 Trainer"
   ClientHeight    =   7845
   ClientLeft      =   5940
   ClientTop       =   540
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   5445
   Begin VB.Timer tmrHealth 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2280
      Top             =   3000
   End
   Begin VB.CommandButton cmdHealth 
      Caption         =   "Enable"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Trainer"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   7320
      Width           =   5175
   End
   Begin VB.Timer tmrGrenades 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1800
      Top             =   3000
   End
   Begin VB.Timer tmrSecondary 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1320
      Top             =   3000
   End
   Begin VB.CommandButton cmdGrenades 
      Caption         =   "Enable"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdSecondary 
      Caption         =   "Enable"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   5280
      Width           =   855
   End
   Begin VB.Timer tmrSidearm 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   840
      Top             =   3000
   End
   Begin VB.Timer tmrPrimary 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   360
      Top             =   3000
   End
   Begin VB.CommandButton cmdSidearm 
      Caption         =   "Enable"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdPrimary 
      Caption         =   "Enable"
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Unlimited Health (SP or MP host)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   6720
      Width           =   3975
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H8000000A&
      Height          =   495
      Left            =   120
      Top             =   6600
      Width           =   5175
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000A&
      Height          =   495
      Left            =   120
      Top             =   5880
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Unlimited Grenades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000A&
      Height          =   495
      Left            =   120
      Top             =   5160
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Unlimited secondary weapon ammo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Unlimited sidearm weapon ammo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000A&
      Height          =   495
      Left            =   120
      Top             =   4440
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Unlimited primary weapon ammo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000A&
      Height          =   495
      Left            =   120
      Top             =   3720
      Width           =   5175
   End
   Begin VB.Image imgLogo 
      Height          =   3480
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   5475
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'|---
'|   Trainer for 'Delta Force 2 V1. 06.15 '
'|---
'|  Description:
'|  This program is used to get infinite ammunation
'|  in the game DF2 by manipulating the ammo values
'|  in the memory for each weapon.
'|  Timers are used to change the value to max when
'|  you are using the ammo in the game, so the ammo
'|  meter will not decrease and you do not need to reload.
'|---
'|  Please feel free to use this code for educating purpose
'|  and for learning to write your own trainers of this type.
'|  I spent some time looking up the location offsets for the
'|  ammo values in the game and coding this trainer
'|  so copying the whole code and putting your name on it
'|  would be considered too easy (for not to say lame) so
'|  please prove that your own name is not the only thing you can write!
'|---
'|  Code written by Cola-Kattn AKA [Tech]
'|---

Dim Program As New Class1
Dim ExePath As String
Dim WindowName As String

Private Sub Form_Load()
    'Set the window name:
    WindowName = "Delta Force 2,  V1.06.15"
    'Find the executable path for Df2
    ExePath = ReadKey("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths\df2.exe\Path")
    '** I got an error when trying to execute df2 with the shell command, but I'll keep this anyway **
End Sub

'Command buttons for enabling\disabling the different timers:

Private Sub cmdPrimary_Click()

    Dim TestFindWindow As Integer
    'Test to see if the df2 game window can be found
    TestFindWindow = Program.Test(WindowName)
    If TestFindWindow = 0 Then
        MsgBox "The game is not running or the version number differs.", vbCritical, "Error"
        Exit Sub
    End If
    
    If cmdPrimary.Caption = "Enable" Then
        tmrPrimary.Enabled = True
        cmdPrimary.Caption = "Disable"
    Else
        tmrPrimary.Enabled = False
        cmdPrimary.Caption = "Enable"
    End If
    
End Sub

Private Sub cmdSidearm_Click()

    Dim TestFindWindow As Integer
    
    TestFindWindow = Program.Test(WindowName)
    If TestFindWindow = 0 Then
        MsgBox "The game is not running or the version number differs.", vbCritical, "Error"
        Exit Sub
    End If

    If cmdSidearm.Caption = "Enable" Then
        tmrSidearm.Enabled = True
        cmdSidearm.Caption = "Disable"
    Else
        tmrSidearm.Enabled = False
        cmdSidearm.Caption = "Enable"
    End If
    
End Sub

Private Sub cmdSecondary_Click()

    Dim TestFindWindow As Integer
    
    TestFindWindow = Program.Test(WindowName)
    If TestFindWindow = 0 Then
        MsgBox "The game is not running or the version number differs.", vbCritical, "Error"
        Exit Sub
    End If

    If cmdSecondary.Caption = "Enable" Then
        tmrSecondary.Enabled = True
        cmdSecondary.Caption = "Disable"
    Else
        tmrSecondary.Enabled = False
        cmdSecondary.Caption = "Enable"
    End If
    
End Sub

Private Sub cmdGrenades_Click()

    Dim TestFindWindow As Integer
    
    TestFindWindow = Program.Test(WindowName)
    If TestFindWindow = 0 Then
        MsgBox "The game is not running or the version number differs.", vbCritical, "Error"
        Exit Sub
    End If

    If cmdGrenades.Caption = "Enable" Then
        tmrGrenades.Enabled = True
        cmdGrenades.Caption = "Disable"
    Else
        tmrGrenades.Enabled = False
        cmdGrenades.Caption = "Enable"
    End If
    
End Sub

Private Sub cmdHealth_Click()

    Dim TestFindWindow As Integer
    
    TestFindWindow = Program.Test(WindowName)
    If TestFindWindow = 0 Then
        MsgBox "The game is not running or the version number differs.", vbCritical, "Error"
        Exit Sub
    End If
        
    If cmdHealth.Caption = "Enable" Then
        tmrHealth.Enabled = True
        cmdHealth.Caption = "Disable"
    Else
        tmrHealth.Enabled = False
        cmdHealth.Caption = "Enable"
    End If
End Sub


'End of command button list.

'The following timer events are for every type of weapons (Primary, Sidearm, Secondary and Grenades)
'Each time the events is run, the current ammo value for each weapon is tested,
'if the value is under the max for the specific weapon, the value will be reset to max.
Private Sub tmrPrimary_Timer()

    Dim value As Byte
    
    'M249
    Program.ReadByte WindowName, &H79A40C, value
    If value < 200 Then Program.WriteByte WindowName, &H79A40C, 200
    
    'Barret .50
    Program.ReadByte WindowName, &H79A410, value
    If value < 10 Then Program.WriteByte WindowName, &H79A410, 10

    'MP5
    Program.ReadByte WindowName, &H79A3E4, value
    If value < 30 Then Program.WriteByte WindowName, &H79A3E4, 30
    
    'M40
    Program.ReadByte WindowName, &H79A42C, value
    If value < 5 Then Program.WriteByte WindowName, &H79A42C, 5

    'UAR
    Program.ReadByte WindowName, &H79A448, value
    If value < 30 Then Program.WriteByte WindowName, &H79A448, 30

    'M4, burst
    Program.ReadByte WindowName, &H79A3F0, value
    If value < 30 Then Program.WriteByte WindowName, &H79A3F0, 30

End Sub

Private Sub tmrSidearm_Timer()
 
    Dim value As Byte
 
    'SOCOM .45
    Program.ReadByte WindowName, &H79A400, value
    If value < 12 Then Program.WriteByte WindowName, &H79A400, 12

    'SOCOM .45 SUPPRESSED
    Program.ReadByte WindowName, &H79A418, value
    If value < 12 Then Program.WriteByte WindowName, &H79A418, 12

    'P11
    Program.ReadByte WindowName, &H79A444, value
    If value < 6 Then Program.WriteByte WindowName, &H79A444, 6

End Sub

Private Sub tmrSecondary_Timer()

    Dim value As Byte

    'LAW
    Program.ReadByte WindowName, &H79A540, value
    If value < 5 Then Program.WriteByte WindowName, &H79A540, 5

    'Satchel Charge
    Program.ReadByte WindowName, &H79A4DC, value
    If value < 5 Then Program.WriteByte WindowName, &H79A4DC, 5

    'Claymore
    Program.ReadByte WindowName, &H79A4FC, value
    If value < 5 Then Program.WriteByte WindowName, &H79A4FC, 5

    'Camera
    Program.ReadByte WindowName, &H79A56C, value
    If value < 5 Then Program.WriteByte WindowName, &H79A56C, 5

End Sub

Private Sub tmrGrenades_Timer()

    Dim value As Byte

    'Impact frag grenade
    Program.ReadByte WindowName, &H79A504, value
    If value < 5 Then Program.WriteByte WindowName, &H79A504, 5

    'Delay frag grenade
    Program.ReadByte WindowName, &H79A558, value
    If value < 5 Then Program.WriteByte WindowName, &H79A558, 5

End Sub

Private Sub tmrHealth_Timer()
    
    Dim value As Byte
    
    'Health in multiplayer
    Program.ReadByte WindowName, &HC7214C, value
    If value < 255 Then Program.WriteByte WindowName, &HC7214C, 255
    
    'Health in single player
    Program.ReadByte WindowName, &HC73904, value
    If value < 255 Then Program.WriteByte WindowName, &HC73904, 255
    
End Sub

'End of timers.

Private Sub cmdExit_Click()

    'Exit program.
    Unload Form1
    
End Sub

'The subs below is just misc stuff used to make the trainer more functional

Private Sub cmdStartDf2_Click()
    'This does not work, df2 will give you an error.
    'if you want to try, just make a command button named cmdStartDf2 and
    'please notify me if it worked for you.
    If ExePath = "" Then
        MsgBox "Unable to locate Delta Force 2 on you harddrive", vbCritical, "Error"
        Exit Sub
    End If
    Shell (ExePath & "\df2.exe")
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Move the form if the mouse button is hold down
    Program.MoveForm Me, Button
    
End Sub

Private Sub imgLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Move the form if the mouse button is hold down on the logo
    Program.MoveForm Form1, Button
    
End Sub

'Read a key value from the windows registry.
'(I use this to find the Delta Force 2 executable file)
Private Property Get ReadKey(value As String) As String

    Dim b As Object
    On Error Resume Next
    Set b = CreateObject("wscript.shell")
    r = b.RegRead(value)
    ReadKey = r
    
End Property
