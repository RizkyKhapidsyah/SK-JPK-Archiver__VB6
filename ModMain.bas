Attribute VB_Name = "ModMain"
Function GetFileName(Path As String) As String

    For Findsep = 1 To Len(Path)
        If Mid(Path, Len(Path) - (Findsep - 1), 1) = "\" Or Mid(Path, Len(Path) - (Findsep - 1), 1) = "/" Then
            GetFileName = Right(Path, Findsep - 1)
            Exit Function
        End If
    Next Findsep

End Function

Function GetPath(FullPath As String) As String
    
    Dim C As Integer
    Dim S As Integer
    Dim J As Integer

    C = 0: S = 0: J = 0
    
    For M = 1 To Len(FullPath)
        GetChr0 = Right(FullPath, M): GetChr1 = Left(GetChr0, 1)
        If GetChr1 = "\" Or GetChr1 = "/" Then C = C + 1
    Next M
    For M = 1 To Len(FullPath)
        GetChr0 = Left(FullPath, M): GetChr1 = Right(GetChr0, 1)
        J = J + 1
        If GetChr1 = "\" Or GetChr1 = "/" Then
            J = 0: S = S + 1
            If S = C Then GetPath = Right(GetChr0, M - J): Exit Function
        End If
    Next M

End Function
