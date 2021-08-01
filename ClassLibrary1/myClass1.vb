

Imports System.Runtime.CompilerServices

Namespace UtilityLibraries

    Public Class string_comp
        Public Function comp_string(str As String)

            Dim bool1 As Boolean = False
            Dim chr1 As Char
            For Each chr1 In str
                If Char.IsUpper(chr1) Then
                    bool1 = True
                End If
            Next
            Return bool1

        End Function
    End Class

End Namespace
