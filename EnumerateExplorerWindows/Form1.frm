VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GogoX@Lycos.com"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Enumerate"
      Height          =   510
      Left            =   9315
      TabIndex        =   1
      Top             =   4950
      Width           =   1590
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10860
   End
   Begin VB.Label Label1 
      Caption         =   "Open several Windows Explorer and/or Internet Explorer windows and press button 'Enumerate'"
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   4590
      Width           =   10815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

 ' You must add a reference to 'Microsoft Shell Controls And Automation'
 ' %SYSTEM_PATH%\Shell32.dll
Dim objShell32 As New Shell32.Shell
Dim objWindows As Object
Dim lngCounter As Long
Const strPause = "     "
 ' Get collection of open folder windows
Set objWindows = objShell32.Windows

 ' Prepare list box
List1.Clear
List1.AddItem "Explorer windows count =  " & objWindows.Count


 ' Iterate thru all objects if present
For Each objWindows In objShell32.Windows
    List1.AddItem ""
    lngCounter = lngCounter + 1
    List1.AddItem "Number " & lngCounter
    List1.AddItem strPause & "Path = " & objWindows.Path
    List1.AddItem strPause & "Full name = " & objWindows.FullName
    List1.AddItem strPause & "Name = " & objWindows.Name
    List1.AddItem strPause & "hWnd = " & objWindows.hWnd
    List1.AddItem strPause & "Location = " & objWindows.LocationName
    List1.AddItem strPause & "URL = " & objWindows.LocationURL
    List1.AddItem strPause & "Status text = " & objWindows.StatusText
Next

 ' Clear objects
Set objWindows = Nothing
Set objShell32 = Nothing

End Sub


