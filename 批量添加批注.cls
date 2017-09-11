VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub 批量增加批注()
On Error Resume Next '忽略错误继续执行VBA代码,避免出现错误消息
Application.ScreenUpdating = False '关闭屏幕刷新
Application.Calculation = xlCalculationManual '手动重算
Dim i As Integer, j As Integer
For i = 1 To 142
    y = 2
    With Cells(2, i)
    .AddComment
    .Comment.Visible = False
    .Comment.Text Text:="ts6_A" & i
    End With
Next
iApplication.Calculation = xlCalculationAutomatic '自动重算
Application.ScreenUpdating = True '打开屏幕刷新
On Error GoTo 0 '恢复正常的错误提示
End Sub
