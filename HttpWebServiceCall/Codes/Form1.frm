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
    Dim url As String
    
    url = "http://api.myideaph.com/echo/?id=123&name=john+doe"
    resp = getResponse(url)
    
    Debug.Print ("Delimited: " & resp)
    
    url = "http://api.myideaph.com/echo/js/?id=123&name=john+doe"
    resp = getResponse(url)
    
    Debug.Print ("JSON: " & resp)
End Sub

Private Function getResponse(ByVal uriPath As String) As String
    Dim objHttp As WinHttp.WinHttpRequest
    
    Set objHttp = New WinHttp.WinHttpRequest
    Call objHttp.Open("get", uriPath)
    objHttp.Send
    
    getResponse = objHttp.ResponseText
End Function
