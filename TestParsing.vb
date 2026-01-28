Module TestParsing

    ''' <summary>
    ''' Public routine to call and test parser class.
    ''' </summary>
    ''' <param name="thisAddress">Whole address to be parsed.</param>
    ''' <returns>Code for success.</returns>
    Public Function ShowParts(ByVal thisAddress As String) As Integer
        Dim p As New Parser()
        p.Parse(thisAddress)
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
        Debug.Print("CLEAN:  " & p.CleanStreet)
        Return p.Result
    End Function

End Module
