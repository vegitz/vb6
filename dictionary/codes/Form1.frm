VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim dctConn As Scripting.Dictionary
    
    Set dctConn = New Scripting.Dictionary
    dctConn("dsn") = "mysimpledb"
    dctConn("uid") = "myself"
    dctConn("pwd") = "secret"
    
    connect_to_db dctConn
End Sub

Private Sub connect_to_db(ByVal conn As Scripting.Dictionary)
    For Each Key In conn
        Debug.Print (Key & " = " & conn(Key))
    Next
End Sub
