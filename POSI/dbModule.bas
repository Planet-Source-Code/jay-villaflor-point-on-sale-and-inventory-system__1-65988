Attribute VB_Name = "dbModule"
Public db As DAO.Database
Public userLog As String

Public Sub openDB()
Set db = OpenDatabase(App.Path & "\POSI.mdb")
End Sub

Public Function execQuery(ByVal sqlStr As String) As String
db.Execute (sqlStr)
End Function

Public Function txtCaps(ByVal strData As String) As String
cr$ = Chr$(13) + Chr$(10)
t$ = strData
If t$ <> "" Then
 Mid$(t$, 1, 1) = UCase$(Mid$(t$, 1, 1))
 For i = 1 To Len(t$) - 1
   If Mid$(t$, i, 2) = cr$ Then Mid$(t$, i + 2, 1) = UCase$(Mid$(t$, i + 2, 1))
   If Mid$(t$, i, 1) = " " Then Mid$(t$, i + 1, 1) = UCase$(Mid$(t$, i + 1, 1))
 Next
 txtCaps = t$
End If
End Function

Public Function Marquee(ByVal MyText As String, ByVal Num As Integer) As String
Dim Tx As String
 Static n As Integer
 Static n2 As Integer
 Tx = Space(Num)
 n = n + n2
  If n > Num - Len(MyText) Then n2 = -1
   If n < 1 Then n = 2: n2 = 1
   Mid$(Tx, n, Len(MyText)) = MyText
   Marquee = Tx
End Function

