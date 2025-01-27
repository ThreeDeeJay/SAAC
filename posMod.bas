Attribute VB_Name = "PosMod"
Private Const D_CONV As Double = 3.14159265358979 / 180
Private Const longLen As Integer = 1500
Private Const shorLen As Integer = 1000
Private cenX As Integer
Private cenY As Integer
Public Function SetPos(pctName As Object, pos As Long, Optional mode As Long, Optional wheelM As Long)
On Error Resume Next
If mode = 1 Then
If pos = 0 Then
pctName(1).Visible = False
ElseIf pctName(1).Visible = False Then
pctName(1).Visible = True
End If
pctName(1).Height = pctName(0).Height
pctName(1).Top = 0
pctName(1).Left = 0
pctName(1).Width = (pctName(0).Width / 510) * (pos)
Exit Function
End If
If pos = 510 Then
pctName(1).Visible = False
ElseIf pctName(1).Visible = False Then
pctName(1).Visible = True
End If

pctName(1).Height = pctName(0).Height
pctName(1).Top = 0
If wheelM = 1 Then
pctName(1).Left = 0
pctName(1).Width = (pctName(0).Width / 510) * (510 - pos)
Exit Function
End If
If pos = 255 Then
pctName(1).Visible = False
ElseIf pctName(1).Visible = False Then
pctName(1).Visible = True
End If
If pos < 255 Then
pctName(1).Width = (pctName(0).Width / 2) - ((pctName(0).Width / 2) - ((pctName(0).Width / 510) * (255 - pos)))
pctName(1).Left = (pctName(0).Width / 2) - ((pctName(0).Width / 510) * (255 - pos))
Else
pctName(1).Left = (pctName(0).Width / 2)
pctName(1).Width = (pctName(0).Width / 510) * (pos - 255)
End If
End Function

Public Function SetPOV(lnName As Object, pos2 As Long)
Dim pos As Long
If pos2 = -1 Then lnName.X1 = lnName.X2: lnName.Y1 = lnName.Y2: Exit Function
pos = pos2 / 100
cenY = lnName.Y2
cenX = lnName.X2
    lnName.X1 = cenX + Sin(pos * D_CONV) * 250
    lnName.Y1 = cenY - Cos(pos * D_CONV) * 250
End Function


