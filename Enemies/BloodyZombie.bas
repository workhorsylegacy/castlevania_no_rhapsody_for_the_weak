Attribute VB_Name = "Module4"

Option Explicit


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' GENERAL DECLARATIONS
' -------------------------------------------------------------------
' BLOODY ZOMBIE ANIMATIONS
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/

Function BloodyZombieWalk(W As Long, D As Long, S As Long, L As Boolean) As Long
'Walk Right 8 to 9, 1 to 6
    With Enemy(CH)
        If W < 0 Or W > 6 Then W = 1
        If .pLeft = True Then
            EnemyMoveLeft 1, 1
        Else
            EnemyMoveRight 1, 1
        End If
        If D = 10 Then
            D = 0
            If W >= 0 And W < 6 Then
                S = IIf(L = True, WALKLEFT, WALKRIGHT)
                W = W + 1
            ElseIf W = 6 Then
                W = 1
            End If
        Else
            D = D + 1
        End If
    End With
End Function
