Public Class AA

    Public Function check_afm(ByVal M_AFM As String) As Integer

        '<EhHeader>
        On Error GoTo check_afm_Err
        M_AFM = M_AFM.Trim
        '</EhHeader>
        Dim SUMA, k As Long
        Dim l As Integer = Len(M_AFM)
100:    SUMA = 0
110:    check_afm = 1
120:    k = 1

130:    For k = 1 To 8
            If Not IsNumeric(Mid(M_AFM, k, 1)) Then
                check_afm = 0
                Exit Function
            End If
140:        SUMA = SUMA + Val(Mid(M_AFM, k, 1)) * 2 ^ (9 - k)

        Next

150:    If SUMA Mod 11 <> Val(Mid(M_AFM, l, 1)) Then
160:        If SUMA Mod 11 = 10 And Val(Mid(M_AFM, l, 1)) = 0 Then
            Else
170:            MsgBox("Λάθος στο ΑΦΜ " + M_AFM)
180:            check_afm = 0
            End If
        End If
        If l <> 9 Then
            check_afm = 0
        End If
        '<EhFooter>
        Exit Function

check_afm_Err:

        Resume Next

        '</EhFooter>

    End Function

End Class
