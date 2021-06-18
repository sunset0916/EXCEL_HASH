VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Ganarate SHA256"
   ClientHeight    =   1400
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4510
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    
    Dim hash As String
    
    If TextBox1.Text = "" Then
        MsgBox "文字を入力してください。"
        Exit Sub
    End If
    
    hash = SHA256(TextBox1.Text)
    
    MsgBox (hash)
    
End Sub

Function SHA256(str As String) As String
    Dim objSHA256 As Object
    Dim objUTF8 As Object
    Dim bytes() As Byte
    Dim hash() As Byte
    Dim strSHA As String

    Set objSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    Set objUTF8 = CreateObject("System.Text.UTF8Encoding")

    bytes = objUTF8.GetBytes_4(str)

    hash = objSHA256.ComputeHash_2((bytes))
        
    Dim i, wk
    
    For i = 1 To UBound(hash) + 1
        wk = wk & Right("0" & Hex(AscB(MidB(hash, i, 1))), 2)
    Next i
    
    SHA256 = wk
    

End Function

Private Sub CommandButton2_Click()

    End

End Sub
