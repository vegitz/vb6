VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >>"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<< Prev"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private xlapp As Object
Private xlwb As Object
Private xlws As Object

Private xldir As String
Private xlpath As String

Private total_rows As Long
Private row As Integer
Private col As Integer
Private starting_row As Integer


Private Sub cmdNext_Click()
    If row < total_rows Then
        row = row + 1
        display_record_at row
    End If
End Sub

Private Sub cmdPrev_Click()
    If row > starting_row Then
        row = row - 1
        display_record_at row
    End If
End Sub

Private Sub display_record_at(ByVal where As Integer)
     Const COLUMN_NAME As Integer = 2
     Const COLUMN_ADDRESS As Integer = 7
     
     Me.txtName.Text = xlws.cells(where, COLUMN_NAME)
     Me.txtAddress.Text = xlws.cells(where, COLUMN_ADDRESS)
End Sub

Private Sub Form_Load()
    Dim xcell As Object ' Excel.Range
    
    xldir = App.Path
    If Right(xldir, 1) <> "\" Then
        xldir = xldir & "\"
    End If
    
    xlpath = xldir & "sampledata2.xls"
    
    Set xlapp = CreateObject("Excel.Application")
    
    Set xlwb = xlapp.workbooks.open(xlpath)
    Set xlws = xlwb.worksheets("Sheet1")
    
    total_rows = xlws.usedrange.rows.Count
    starting_row = 2            ' if without header, set to 1
    
    row = starting_row
    display_record_at row
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xlwb.Close
    xlapp.Quit
End Sub
