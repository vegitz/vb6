VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   5
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   480
      Width           =   5655
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtStudentId 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================
' phVB6project | 2020-Mar
'............................................................
'
' demo app that uses an Excel (97 to 2000) file as database
' and the worksheet as table using VB6
'
' this demo contains code for:
' -> opening the excel file as database
' -> opening the worksheet as table
' -> displaying the fields
' -> move forward/next
' -> move backward/previous
' -> search entries
'
' please make sure to name your controls accordingly
' or re-name the variables used here to fit your project
'
'------------------------------------------------------------
Option Explicit

Private xl_connstr As String
Private cnn As ADODB.Connection
Private rst As ADODB.Recordset

Private Sub cmdNext_Click()
    move_next rst
    display_current
End Sub

Private Sub display_current()
    Dim data As ADODB.Fields
    Dim prop_name As String
    Dim prop_value As String
    Dim idx As Integer
    
    Set data = get_record(rst)
    If data Is Nothing Then
        MsgBox "No more data", vbInformation
    Else
        Me.txtStudentId.Text = data.Item("student no").Value
        Me.txtName.Text = data.Item("student name").Value
    End If
End Sub

Private Sub cmdPrev_Click()
    move_prev rst
    display_current
End Sub

Private Sub cmdSearch_Click()
    search_for Me.txtSearch.Text, rst
    display_current
End Sub

Private Sub Form_Load()
    ' define connection string (replace the "data source" with your excel file)
    xl_connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\sampledata2.xls;Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
    ' create and open the connection object
    Set cnn = New ADODB.Connection
    cnn.Open (xl_connstr)
    ' open the worksheet as table
    Set rst = New ADODB.Recordset
    ' place the content of worksheet on recordset
    rst.Open "select * from [sheet1$]", cnn, adOpenKeyset, adLockPessimistic    ' note the need to add dollar sign ($)
    rst.MoveFirst
    
    display_current
End Sub

Private Function get_record(ByVal rs As ADODB.Recordset) As ADODB.Fields
    If rs.EOF Or rs.BOF Then
        Debug.Print "no record at this position"
        Set get_record = Nothing
    Else
        Debug.Print "getting record at this position"
        Set get_record = rs.Fields
    End If
End Function

Private Sub search_for(ByVal criteria As String, ByVal rs As ADODB.Recordset)
    rs.MoveFirst
    rs.Filter = criteria
End Sub

Private Sub move_next(ByVal rs As ADODB.Recordset)
    If Not rs.EOF Then
        Debug.Print "moving to next record..."
        rs.MoveNext
        ' if we moved past the last, go back
        If rs.EOF Then
            Debug.Print "going back to last record"
            rs.MovePrevious
        End If
    End If
End Sub

Private Sub move_prev(ByVal rs As ADODB.Recordset)
    If Not rs.BOF Then
        Debug.Print "moving to previous record..."
        rs.MovePrevious
        ' if we moved past the first, go forward
        If rs.BOF Then
            Debug.Print "going forward to first record"
            rs.MoveNext
        End If
    End If
End Sub

