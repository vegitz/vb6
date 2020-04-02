VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4895
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim logs As String
    Dim logfn As String
    
    logfn = App.Path & "\sys.log"
    logs = load_logs(logfn)
    
    display_logs logs, False
End Sub

Private Sub display_logs(ByVal logs As String, ByVal include_first_row As Boolean)
    Dim rows() As String
    Dim cols() As String
    
    Dim row_index As Integer
    Dim start_index As Integer
    
    
    rows = Split(logs, vbCrLf)
    
    If include_first_row Then
        start_index = 0
    Else
        start_index = 1
    End If
    
    With Me.ListView1
        .View = lvwReport
        With .ColumnHeaders
            .Add , , "Date", 1000
            .Add , , "Time", 1000
            .Add , , "Description", 2000
        End With
        With .ListItems
            .Clear
            For row_index = start_index To UBound(rows)
                cols = Split(rows(row_index), ",", 3)
                With .Add(, , cols(0))
                    .ListSubItems.Add , , cols(1)
                    .ListSubItems.Add , , cols(2)
                End With
            Next
        End With
    End With
End Sub


Private Function load_logs(ByVal log_fn As String) As String
    Dim content As String
    
    Open log_fn For Input As #1
    content = Input(LOF(1), 1)
    Close #1
    
    load_logs = content
End Function
