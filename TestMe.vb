''' <summary>
''' For testing parser class.
''' </summary>
Module TestMe

    Public Sub ShowParts(ByVal thisAddress As String)
        Dim p As New Parser()
        p.Parse(thisAddress)
        Debug.Print(p.Result.ToString)
        Debug.Print("HOUSE:    " & p.HouseNum)
        Debug.Print("HOUSESUF: " & p.HouseSuf)
        Debug.Print("PREDIR:   " & p.PreDir)
        Debug.Print("STREET:   " & p.Street)
        Debug.Print("STSUF:    " & p.StSuf)
        Debug.Print("POSTDIR:  " & p.PostDir)
        If p.SUD.Count > 0 Then
            Debug.Print("SUD:      " & p.SUD(0).SUD)
            Debug.Print("RANGE:    " & p.SUD(0).Range)
        End If
    End Sub

End Module
