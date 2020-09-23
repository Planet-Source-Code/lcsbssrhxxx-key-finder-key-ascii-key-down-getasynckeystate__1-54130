VERSION 5.00
Begin VB.Form frmKeyFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Key Finder 2005"
   ClientHeight    =   1815
   ClientLeft      =   1080
   ClientTop       =   1350
   ClientWidth     =   3495
   Icon            =   "frmKeyFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3495
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Key Down (Key Code)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label KeyDown 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label KeyAsync 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label letter 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Character "
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Key Press(Key ASCII)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "GetAsyncKeyState"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label KeyASCI 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   1
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu mini 
         Caption         =   "Minimize"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmKeyFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Integer
Private Sub exit_Click()
End
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call clear
A = KeyCode
KeyDown.Caption = KeyCode
KeyAsync.Caption = KeyCode
Call FindLetter
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
KeyASCI.Caption = KeyAscii
Call FindLetter
End Sub
Private Sub mini_Click()
frmKeyFind.WindowState = 1
End Sub
Private Sub tmrclock_Timer()
clock.Caption = Time
End Sub
Private Sub clear()
KeyDown.Caption = ""
KeyASCI.Caption = ""
End Sub
Private Sub FindLetter()
If A = 65 Then letter.Caption = "a / A"
If A = 66 Then letter.Caption = "b / B"
If A = 67 Then letter.Caption = "c / C"
If A = 68 Then letter.Caption = "d / D"
If A = 69 Then letter.Caption = "e / E"
If A = 70 Then letter.Caption = "f / F"
If A = 71 Then letter.Caption = "g / G"
If A = 72 Then letter.Caption = "h / H"
If A = 73 Then letter.Caption = "i / I"
If A = 74 Then letter.Caption = "j / J"
If A = 75 Then letter.Caption = "k / K"
If A = 76 Then letter.Caption = "l / L"
If A = 77 Then letter.Caption = "m / M"
If A = 78 Then letter.Caption = "n / N"
If A = 79 Then letter.Caption = "o / O"
If A = 80 Then letter.Caption = "p / P"
If A = 81 Then letter.Caption = "q / Q"
If A = 82 Then letter.Caption = "r / R"
If A = 83 Then letter.Caption = "s / S"
If A = 84 Then letter.Caption = "t / T"
If A = 85 Then letter.Caption = "u / U"
If A = 86 Then letter.Caption = "v / V"
If A = 87 Then letter.Caption = "w / W"
If A = 88 Then letter.Caption = "x / X"
If A = 89 Then letter.Caption = "y / Y"
If A = 90 Then letter.Caption = "z / Z"

If A = 48 Then letter.Caption = "0"
If A = 49 Then letter.Caption = "1"
If A = 50 Then letter.Caption = "2"
If A = 51 Then letter.Caption = "3"
If A = 52 Then letter.Caption = "4"
If A = 53 Then letter.Caption = "5"
If A = 54 Then letter.Caption = "6"
If A = 56 Then letter.Caption = "8"
If A = 57 Then letter.Caption = "9"

If A = 92 Then letter.Caption = "{WIN}"
'num lock
If A = 144 Then letter.Caption = "{NUMLOCK}"
If A = 96 Then letter.Caption = "0"
If A = 97 Then letter.Caption = "1"
If A = 98 Then letter.Caption = "2"
If A = 99 Then letter.Caption = "3"
If A = 100 Then letter.Caption = "4"
If A = 101 Then letter.Caption = "5"
If A = 102 Then letter.Caption = "6"
If A = 103 Then letter.Caption = "7"
If A = 104 Then letter.Caption = "8"
If A = 105 Then letter.Caption = "9"
',<
If A = 188 Then letter.Caption = ","
'.>
If A = 190 Then letter.Caption = "."
'/?
If A = 191 Then letter.Caption = "/"
';:
If A = 186 Then letter.Caption = ";"
'/?
If A = 222 Then letter.Caption = "'"
'[{
If A = 219 Then letter.Caption = "["
']}
If A = 221 Then letter.Caption = "]"
'\|
If A = 220 Then letter.Caption = "\"
'`~
If A = 192 Then letter.Caption = "`"
'ESC
If A = 27 Then letter.Caption = "{ESC}"
'= +
If A = 187 Then letter.Caption = "="
'- _                            .
If A = 189 Then letter.Caption = "-"
'shift
If A = 16 Then letter.Caption = "{SHIFT}"
'space
If A = 32 Then letter.Caption = "{SPACE}"
'ctrl
If A = 17 Then letter.Caption = "{CTRL}"
'alt
If A = 18 Then letter.Caption = "{ALT}"
'enter
If A = 13 Then letter.Caption = "{ENTER}"
'back space
If A = 8 Then letter.Caption = "{BACKSPACE}"
'arrow keys
If A = 37 Then letter.Caption = "{LEFT}"
If A = 38 Then letter.Caption = "{UP}"
If A = 39 Then letter.Caption = "{RIGHT}"
If A = 40 Then letter.Caption = "{DOWN}"
'above arrow keys
If A = 45 Then letter.Caption = "{INS}"
If A = 46 Then letter.Caption = "{DEL}"
If A = 36 Then letter.Caption = "{HOME}"
If A = 33 Then letter.Caption = "{PGUP}"
If A = 34 Then letter.Caption = "{PGDOWN}"
If A = 35 Then letter.Caption = "{END}"
'f1 to f12
If A = 112 Then letter.Caption = "F1"
If A = 113 Then letter.Caption = "F2"
If A = 114 Then letter.Caption = "F3"
If A = 115 Then letter.Caption = "F5"
If A = 116 Then letter.Caption = "F5"
If A = 117 Then letter.Caption = "F6"
If A = 118 Then letter.Caption = "F7"
If A = 119 Then letter.Caption = "F8"
If A = 120 Then letter.Caption = "F9"
If A = 121 Then letter.Caption = "F10"
If A = 122 Then letter.Caption = "F11"
If A = 123 Then letter.Caption = "F12"
End Sub

