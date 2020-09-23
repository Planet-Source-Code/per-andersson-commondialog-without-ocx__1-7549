VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Commondialog without OCX"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9705
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   615
      Left            =   0
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "www.FireStormEntertainment.Cjb.Net"
      Height          =   195
      Left            =   7080
      TabIndex        =   3
      Top             =   0
      Width           =   2610
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Filename:"
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   675
   End
   Begin VB.Label Label1 
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   8760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Dim f As Boolean, f32Options As Long
    Dim sFileTitle As String, sFilter As String
    sFilter = "Executable |*.exe;*.dll;*.vbx;*.ocx;*.fon|"
    sFilter = sFilter & "Program |*.exe|DLL |*.dll|"
    sFilter = sFilter & "Control |*.vbx;*.ocx|Font | *.fon"
    f32Options = cdlOFNHideReadOnly Or cdlOFNFileMustExist
    f = VBGetOpenFileName( _
        vFileName:=sExe, _
        vFileTitle:=sFileTitle, _
        vFlags:=f32Options, _
        vOwner:=Me.hWnd, _
        vFilter:=sFilter)
        
        'Put the filename from the commondialogbox to Label1
        Label1.Caption = sExe
End Sub

