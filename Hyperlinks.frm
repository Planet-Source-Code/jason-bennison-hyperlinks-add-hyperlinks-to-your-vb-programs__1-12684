VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hyperlinks in a VB Program"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "www.codeguru.earthweb.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "Hyperlinks.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "www.vb-world.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "Hyperlinks.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "www.vbcode.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "Hyperlinks.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "www.vbexplorer.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "Hyperlinks.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Please register your vote on PSC :-)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      MouseIcon       =   "Hyperlinks.frx":0C28
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "www.codearchive.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "Hyperlinks.frx":0F32
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "www.planet-source-code.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "Hyperlinks.frx":123C
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Command1_Click()
End
End Sub

'To Put a hyperlink in your program:

'1. In Declarations, declare shell32.dll (and remember to distribute this with your app_
'   Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'2. Put a label control on your VB form

'3. Set the following properties to the label control
'   Font color = blue
'   Font Unterlined
'   Font Bold
'   Allignment = 2 - Centre
'   Caption = What ever you like, but something that describes the webpage you are taking to user to
'   Mousepointer = 99 (custom)
'   MouseIcon = "C:\Program Files\Microsoft Visual Studio\Common\graphics\Cursors\H_point.cur"
'       ** there may be some variation on where this .cur file is on your system, use the Win search to locate it ***

'4 In the click event of the label, enter...
'   ShellExecute hWnd, "open", "http://www.YourWebSite.com", vbNullString, vbNullString, conSwNormal
'   substituting "http://www.YourWebSite.com" with the URL of the webpage you want the user to go

'5. Job Done! - Congrats! you have a hyperlink in your VB program!

'6. You can gimme a vote, big 5 would be nice.

'Regards,  Jason Bennison   jasonbennison@hotmail.com








Private Sub Label1_Click()
ShellExecute hWnd, "open", "http://www.planet-source-code.com", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Label3_Click()
ShellExecute hWnd, "open", "http://www.planet-source-code.com/xq/ASP/txtCodeId.12684/lngWId.1/qx/vb/scripts/ShowCode.htm", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Label4_Click()
ShellExecute hWnd, "open", "http://www.vbexplorer.com", vbNullString, vbNullString, conSwNormal

End Sub

Private Sub Label5_Click()
ShellExecute hWnd, "open", "http://www.vbcode.com", vbNullString, vbNullString, conSwNormal

End Sub

Private Sub Label6_Click()
ShellExecute hWnd, "open", "http://www.vb-world.net", vbNullString, vbNullString, conSwNormal

End Sub

Private Sub Label7_Click()
ShellExecute hWnd, "open", "http://codeguru.earthweb.com", vbNullString, vbNullString, conSwNormal

End Sub
