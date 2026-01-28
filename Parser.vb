Imports System.Text.RegularExpressions

''' <summary>
''' Class to separate address into component parts.
''' </summary>
''' <remarks>
''' Components are exposed as public properties after running Parse method.
''' </remarks>
Public Class Parser

#Region "-- public enum and structures --"

    ''' <summary>
    ''' Return values for Parse method to show if it could not run.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum BadAddress As Integer
        ''' <summary>
        ''' Address was not parsed.
        ''' </summary>
        ''' <remarks></remarks>
        NOT_DONE = -1
        ''' <summary>
        ''' Success.
        ''' </summary>
        ''' <remarks></remarks>
        NO_PROBLEM = 0
        ''' <summary>
        ''' Address does not start with a digit from 1 to 9.
        ''' </summary>
        ''' <remarks></remarks>
        NO_NUMBER = 1
        ''' <summary>
        ''' Address is only a number (or number and letter).
        ''' </summary>
        ''' <remarks></remarks>
        ONLY_NUMBER = 2
        ''' <summary>
        ''' Address is a post office box.
        ''' </summary>
        ''' <remarks></remarks>
        PO_BOX = 3
        ''' <summary>
        ''' Address was intersection without any valid street address
        ''' </summary>
        ''' <remarks></remarks>
        INTERSECTION = 4
        ''' <summary>
        ''' No street name.
        ''' </summary>
        ''' <remarks></remarks>
        NO_STREET = 5
        ''' <summary>
        ''' More than one house number.
        ''' </summary>
        ''' <remarks></remarks>
        TWO_NUMBERS = 6
        ''' <summary>
        ''' Any other error.
        ''' </summary>
        ''' <remarks></remarks>
        UNKNOWN_ERROR = 9
    End Enum

    ''' <summary>
    ''' SUD and associated range.
    ''' </summary>
    ''' <remarks>Used in list of all SUDs and ranges found in address.</remarks>
    Public Structure SUDtype
        ''' <summary>
        ''' Secondary unit designator.
        ''' </summary>
        ''' <remarks>USPS official abbreviation, e.g. "Apt", "Unit".</remarks>
        Public SUD As String
        ''' <summary>
        ''' ID for unit.
        ''' </summary>
        ''' <remarks>Usually a number, but not always, e.g., "Unit 2C" or "2nd" in "2nd floor".</remarks>
        Public Range As String
        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj IsNot SUDtype Then
                Return False
            Else
                With CType(obj, SUDtype)
                    Return .Range = Me.Range And .SUD = Me.SUD
                End With
            End If
        End Function
        Public Overrides Function ToString() As String
            Dim s As String
            If Me.SUD = "FL" And IsNumeric(Me.Range) Then  ' add ordinal suffix before "FLOOR"
                s = Me.Range
                If Len(s) >= 2 AndAlso Mid(s, Len(s) - 1, 1) = "1" Then
                    s &= "TH"
                ElseIf Right(s, 1) = "1" Then
                    s &= "ST"
                ElseIf Right(s, 1) = "2" Then
                    s &= "ND"
                ElseIf Right(s, 1) = "3" Then
                    s &= "RD"
                Else
                    s &= "TH"
                End If
                s &= " FLOOR"
            Else
                If Len(Me.SUD) > 0 Then s = Me.SUD & " " Else s = "#"
                If Len(Me.Range) > 0 Then s &= Range
                s = RTrim(s)
            End If
            Return s
        End Function
    End Structure

#End Region

#Region "-- public properties --"

    ''' <summary>
    ''' Result for parsing.
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property Result As BadAddress
        Get
            Return _rlst
        End Get
    End Property

    ''' <summary>
    ''' Private field for public property.
    ''' </summary>
    ''' <remarks></remarks>
    Private _rlst As BadAddress = BadAddress.NOT_DONE

    ''' <summary>
    ''' House number on address.
    ''' </summary>
    ''' <value>Returns string in case number is too long for integer.</value>
    ''' <returns></returns>
    ''' <remarks>
    ''' Cannot be missing if parsing is successful.
    ''' </remarks>
    Public ReadOnly Property HouseNum() As String
        Get
            HouseNum = _HouseNum
        End Get
    End Property

    ''' <summary>
    ''' Private field for public property.
    ''' </summary>
    ''' <remarks></remarks>
    Private _HouseNum As String

    ''' <summary>
    ''' Letter after house number.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Zero-byte string if none found.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property HouseSuf() As String
        Get
            HouseSuf = _HouseSuf
        End Get
    End Property

    ''' <summary>
    ''' Private field for public property.
    ''' </summary>
    ''' <remarks></remarks>
    Private _HouseSuf As String

    ''' <summary>
    ''' Predirectional abbreviation.
    ''' </summary>
    ''' <value>Abbreviation for direction, one or two letters.</value>
    ''' <returns>Zero-byte string if none found.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PreDir() As String
        Get
            PreDir = _PreDir
        End Get
    End Property

    ''' <summary>
    ''' Private field for public property.
    ''' </summary>
    ''' <remarks></remarks>
    Private _PreDir As String

    ''' <summary>
    ''' Street name.
    ''' </summary>
    ''' <value>Only property that can be more than one word.</value>
    ''' <returns></returns>
    ''' <remarks>
    ''' Cannot be missing if parsing is successful.
    ''' </remarks>
    Public ReadOnly Property Street() As String
        Get
            Street = _Street
        End Get
    End Property

    ''' <summary>
    ''' Private field for public property.
    ''' </summary>
    ''' <remarks></remarks>
    Private _Street As String

    ''' <summary>
    ''' Street suffix abbrevation.
    ''' </summary>
    ''' <value>USPS official abbreviation.</value>
    ''' <returns>Zero-byte string if none found.</returns>
    ''' <remarks>E.g., "St", "Ave", "Dr".</remarks>
    Public ReadOnly Property StSuf() As String
        Get
            StSuf = _StSuf
        End Get
    End Property

    ''' <summary>
    ''' Private field for public property.
    ''' </summary>
    ''' <remarks></remarks>
    Private _StSuf As String

    ''' <summary>
    ''' Postdirectional abbreviation.
    ''' </summary>
    ''' <value>Abbreviation for direction, one or two letters.</value>
    ''' <returns>Zero-byte string if none found.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PostDir() As String
        Get
            PostDir = _PostDir
        End Get
    End Property

    ''' <summary>
    ''' Private field for public property.
    ''' </summary>
    ''' <remarks></remarks>
    Private _PostDir As String

    ''' <summary>
    ''' All SUDs found in address.
    ''' </summary>
    ''' <value>Count equals zero if none fond in address.</value>
    ''' <returns></returns>
    ''' <remarks>NB: there can be more than one, e.g., Building 3 Room 21.</remarks>
    Public ReadOnly Property SUD() As List(Of SUDtype)
        Get
            SUD = _SUD
        End Get
    End Property

    ''' <summary>
    ''' Private field for public property.
    ''' </summary>
    ''' <remarks></remarks>
    Private _SUD As List(Of SUDtype)

    ''' <summary>
    ''' Address text to parse.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RawAddress() As String
        Get
            RawAddress = _RawAddress
        End Get
        Set(value As String)
            _RawAddress = value
            Call Parse(_RawAddress)
        End Set
    End Property

    ''' <summary>
    ''' Private field for public property.
    ''' </summary>
    ''' <remarks></remarks>
    Private _RawAddress As String

    ''' <summary>
    ''' Components put together in one string.
    ''' </summary>
    ''' <returns>String with complete address.</returns>
    Public ReadOnly Property CleanStreet As String

        Get

            Dim dir As String

            Dim s As String = " " & Street

            If Len(StSuf) > 0 Then s &= " " & StSuf

            dir = DirName(PreDir)
            If Len(dir) > 0 Then s = " " & dir & s

            If HouseSuf <> "" Then
                s = HouseSuf + s
                If HouseSuf = "1/2" Then s = " " + s
            End If

            s = HouseNum + s

            dir = DirName(PostDir)
            If Len(dir) > 0 Then s &= " " & dir

            For Each sd In SUD : s &= " " & sd.ToString : Next

            Return s

        End Get

    End Property

#End Region

#Region "-- public methods --"

    ''' <summary>
    ''' Find all address parts.
    ''' </summary>
    ''' <param name="address">Raw address to be parsed.</param>
    ''' <remarks>This method is called recursively to find addresses at the end of the address specified if no address found at beginning.</remarks>
    Public Sub Parse(ByVal address As String)

        Dim OK As Boolean

        ' for parsing with regular expressons
        Dim m As Match
        Dim mc As MatchCollection
        Dim rg As Regex, rg2 As Regex

        Dim hasSpace As Boolean
        Dim fnd As Integer                  ' match found to use
        Dim thisSUD As SUDtype              ' SUD to add
        Dim rng As String                   ' range to add

        Dim StSufmatch As List(Of FoundPatttern)    ' street suffixes found
        Dim foPat As FoundPatttern                  ' for saving patterns found
        Dim fiPat As FindPatttern                   ' for looping thorugh patterns to find
        Dim SUDmatch As New List(Of FoundPatttern)

        ' to save rax suffixes and predirectionals in case they are really part of the street name
        Dim rawSTsuf As String = ""
        Dim rawPredir As String = ""
        Dim rawSUD As String = ""
        Dim stIsSfx As Boolean

        ' save value passed
        _RawAddress = address

        ' string being parsed and cut up
        Dim prs As String = address

        Call Reset()

        _rlst = BadAddress.NO_PROBLEM

#Region "Initial clean up"

        ' remove any phone numbers 

        rg = GetReg("\(?\d{3}(\)|\.|-|\s)\d{3}(\.|-|\s)\d{4}")
        If rg.IsMatch(prs) Then prs = rg.Replace(prs, "")

        ' remove any parentheticals
        ' e.g., "123 MAIN ST (DO NOT SEND MAIL)" becomes "123 MAIN ST"

        rg = GetReg("\s?\(.*\)\s?")
        prs = rg.Replace(prs, " ")

        ' treat consecutive apostrophes as one
        ' e.g, "123 O''FARRELL" becomes "123 O'FARRELL"

        prs = prs.Replace("''", "'")

        ' fix trailing f SUD ranges with single quotes
        ' e.g, "123 MAIN ST APT 'A'" becomes "123 MAIN ST APT A"

        rg = GetReg("'.'$")
        If rg.IsMatch(prs) Then prs = Replace(prs, rg.Match(prs).ToString, Replace(rg.Match(prs).ToString, "'", ""))

        ' omit any date and anything after it

        rg = GetReg("\d{1,2}/\d{1,2}/\d{2,4}")
        prs = rg.Replace(prs, "")

        ' add apostrophe in place of space for "O"
        ' e.g., "123 O FARRELL ST" becomes "123 O'FARRELL ST" 

        rg = GetReg("\bO \w*")
        If rg.IsMatch(prs) Then prs = Replace(prs, rg.Match(prs).Value, Replace(rg.Match(prs).Value, " ", "'"))

        ' treat "0.5" and " .5" as "1/2"

        rg = GetReg("(\s0)?\.5\s")
        m = rg.Match(prs)
        If m.Success Then prs = rg.Replace(prs, " 1/2 ")

        ' fix capital "O" that should be zero: find "O" following any digit
        ' e.g., "12O MAIN ST" becomes "120 MAIN ST"

        rg = GetReg("(\s0)?\dO.+")
        m = rg.Match(prs)
        While m.Success AndAlso UCase(Mid(m.Value, 2, 4)) <> "OFAR"
            Mid(prs, m.Index + 2, 1) = "0"
            m = rg.Match(prs)
        End While

        ' replace periods and commas with spaces

        If Len(prs) > 0 Then
            prs = Replace(prs, ",", " ")
            prs = Replace(prs, ".", " ")
        End If

        ' cut off at "&" or "@"
        ' "101 Grove St @ Polk" is OK, "24th & Mission" is not

        rg = GetReg("[&@]|\sAND\s")
        m = rg.Match(prs)
        If m.Success Then
            prs = Left(prs, m.Index)
            ' count as intersection if it cannot parse
            _rlst = BadAddress.INTERSECTION
        End If

        ' remove all punctuation but hyphen, forward slash, pound sign, and ampersand

        rg = GetReg("[^'&\/\-#\w\s]")
        prs = rg.Replace(prs, "")

        ' no leading zeros

        While Left(prs, 1) = "0" : prs = Mid(prs, 2) : End While

        ' fix missing space on APT, e.g., "STAPT"

        rg = GetReg("\w{2}APT\s?")
        m = rg.Match(prs)
        If m.Success Then
            If m.Value <> "CHAPT" Then prs = rg.Replace(prs, Left(m.Value, 2) & " " & Mid(m.Value, 3))
        End If

        ' remove PO Box and number

        rg = GetReg("\b[p]*(ost)*\.*\s*[o|0]*(ffice)*\.*\s*b[o|0]x((\s?)\d+|$)")
        If rg.IsMatch(prs) Then
            _rlst = BadAddress.PO_BOX
            prs = rg.Replace(prs, "")
        End If

        ' no double spaces

        While InStr(prs, "  ") > 0 : prs = Replace(prs, "  ", " ") : End While

        ' upper-case only, no spaces at beginning or end

        prs = UCase(Trim(prs))

        ' deal with C/O

        If InStr(prs, "C/O") = 1 Then
            ' remove "C/O", parse the rest
            prs = LTrim(Mid(prs, 4))
        ElseIf InStr(prs, "C/O") > 1 Then
            ' remove anything after "C/O"
            prs = RTrim(Left(prs, InStr(prs, "C/O") - 1))
        End If

        ' space between street number and ordinal suffix, e.g. "1 st ave" is "1st ave"

        rg = GetReg("\d+(\s(TH|RD|ND))\s")
        m = rg.Match(prs)
        If m.Success Then
            prs = Replace(prs, m.Groups(1).Value, LTrim(m.Groups(1).Value))
        End If

#End Region

        ' remove "homeless" and "transient" from address, see if it will parse without it
        rg = GetReg("HOME?LE?SS?|TRANS\w*NT")
        m = rg.Match(prs)
        If m.Success Then prs = Trim(rg.Replace(prs, ""))

        ' look for number at beginning NOT part of street name
        rg = GetReg("\A[1-9]")
        rg2 = GetReg("\A\d{1,3}(ST|ND|RD|TH)")

        If Not rg.IsMatch(prs) Or rg2.IsMatch(prs) Then

            ' no house number, do not parse 
            If _rlst = BadAddress.NO_PROBLEM Then _rlst = BadAddress.NO_NUMBER

        Else

            ' omit before hyphen that is the cross street from "Queens-style" hyphenated numbers
            ' e.g., "31-35 55th Street" is "35 55th Street"

            rg = GetReg("\A(\d{2,3})-(\d{2,3})\s")
            m = rg.Match(prs)
            If m.Success Then prs = Mid(prs, InStr(prs, "-") + 1)

            ' handle "1/2" and other fractions: use as house number suffix and remove
            rg = GetReg("\s\d\W+\d\s")
            m = rg.Match(prs)
            If m.Success Then
                _HouseSuf = Trim(m.Value)
                prs = rg.Replace(prs, " ")
            End If

            ' ignore all other slashes
            If InStr(prs, "/") > 0 Then prs = Left(prs, InStr(prs, "/") - 1)

            '   +---------------------------+
            '   |   save house number       |
            '   +---------------------------+

            ' find end of number
            rg = GetReg("[^\d]")
            m = rg.Match(prs)

            If Not m.Success Then
                If _rlst = BadAddress.NO_PROBLEM Then _rlst = BadAddress.ONLY_NUMBER

            Else

                ' cut off house number
                _HouseNum = Left(prs, m.Index)
                prs = Mid(prs, m.Index + 1)

                ' remove hyphen after number
                If Left(prs, 1) = "-" Then prs = Mid(prs, 2)

                If Left(prs, 1) = " " Then
                    hasSpace = True
                    prs = Mid(prs, 2)
                End If

                If Len(prs) <= 1 Then
                    If _rlst = BadAddress.NO_PROBLEM Then _rlst = BadAddress.ONLY_NUMBER
                Else

                    '   +-----------------------------------+
                    '   |   save letter for house suffix    |
                    '   +-----------------------------------+

                    ' NB: suffix or predirectional for N, S, E, W:

                    ' 123E Main St      house suffix
                    ' 123E-Main St      house suffix
                    ' 123 E-Main St     predirectional
                    ' 123-E Main St     house suffix
                    ' 123 E Main St     predirectional

                    ' find space or hyphen before house suffix

                    prs = LTrim(Replace(prs, "-", " "))
                    prs = Replace(prs, "  ", " ")

                    rg = GetReg("\A[A-Z][ -]")
                    If rg.IsMatch(prs) Then

                        rg = GetReg("\A[NSEW] ")
                        rg2 = GetReg("\A(E\sEAST|N\sNORTH|W\wWEST|S\sSOUTH)\b")

                        If Not (hasSpace And rg.IsMatch(prs)) Or rg2.IsMatch(prs) Then
                            _HouseSuf = Left(prs, 1)
                            prs = Mid(prs, 3)
                        End If
                    End If

                    '   +-------------------------------------------+
                    '   |   save SUD and range, remove from address |
                    '   +-------------------------------------------+

                    Call GetSUD(prs, rawSUD)

                    '   +------------------------+
                    '   |   save range alone     |
                    '   +------------------------+

                    If InStr(prs, "#") > 0 Then

                        rng = LTrim(Mid(prs, InStr(prs, "#") + 1))
                        If InStr(rng, " ") > 0 Then rng = Left(rng, InStr(rng, " ") - 1)

                        thisSUD = New SUDtype With {.Range = Replace(rng, "#", "")}

                        _SUD.Add(thisSUD)

                        prs = Left(prs, InStr(prs, "#") - 1) + LTrim(Mid(prs, InStr(prs, "#") + 1 + Len(rng) + 1))

                    End If

                    prs = Trim(prs)

                    '   +------------------------+
                    '   |   save street type     |
                    '   +------------------------+

                    ' look for "Avenue B"

                    rg = GetReg("(?:^|\A)AV[^I](\w+)?\s(?<ST>[A-Z])(?:\W)?(?:\s|$)")
                    m = rg.Match(prs)

                    If m.Success Then

                        rawSTsuf = m.Value
                        _StSuf = "AVE"
                        _Street = m.Groups("ST").Value
                        prs = rg.Replace(prs, "")

                    ElseIf prs Like "CALLE*" Then
                        ' ignore: this is the street type, the rest following it is the name, not the street suffix

                    Else

                        ' find all pattern matches

                        StSufmatch = New List(Of FoundPatttern)
                        For Each fiPat In _lstStSuf
                            rg = GetReg(fiPat.Pattern)
                            mc = rg.Matches(prs)
                            For Each m In mc
                                If m.Success Then

                                    If fiPat.Abrv = "AVE" And Mid(prs, m.Index + 1) Like "AVENIDA*" Then
                                        ' do not count as "AVE"

                                    ElseIf fiPat.Abrv = "HTS" And prs Like "DI*MO*" Then
                                        ' do not count "DIAMOND HTS" as "HTS"

                                    ElseIf fiPat.Abrv = "PARK" And (prs Like "CLINTON *" Or prs Like "ELGIN *") Then
                                        ' "park" is part of the street name

                                    ElseIf Len(_HouseSuf) = 1 Or m.Index > 0 Or fiPat.Abrv = "HWY" Then
                                        foPat = New FoundPatttern With {.Abrv = fiPat.Abrv, .Match = m, .Flag = fiPat.Flag}
                                        StSufmatch.Add(foPat)
                                    End If

                                End If
                            Next

                        Next

                        If StSufmatch.Count = 0 Then
                            For Each fiPat In _lstStSuf
                                If Not fiPat.NoAbrv Then
                                    rg = GetReg("(?:^|\b)" & fiPat.Abrv & "(?:\b|$)")
                                    m = rg.Match(prs)
                                    If m.Success Then
                                        If fiPat.Abrv = "AVE" And Mid(prs, m.Index + 1) Like "AVENIDA*" Then
                                            ' do not count as "AVE"

                                        ElseIf fiPat.Abrv = "HTS" And prs Like "DI*MO*" Then
                                            ' do not count "DIAMOND HTS" as "HTS"
                                        ElseIf fiPat.Abrv = "PARK" And (prs Like "CLINTON *" Or prs Like "ELGIN *") Then
                                            ' "park" is part of the street name
                                        Else
                                            foPat = New FoundPatttern With {.Abrv = fiPat.Abrv, .Match = m, .Flag = fiPat.Flag}
                                            StSufmatch.Add(foPat)
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next

                        End If

                        If StSufmatch.Count > 0 Then  ' some type found

                            If StSufmatch.Count > 1 Then
                                StSufmatch = StSufmatch.OrderBy(Function(x) x.Match.Index).ToList
                                If StSufmatch(0).Abrv = "ST" And StSufmatch(0).Match.Index > 0 Then
                                    If IsNumeric(RTrim(Left(prs, StSufmatch(0).Match.Index))) Then
                                        ' e.g., 1915 81 ST AVE
                                        StSufmatch.RemoveAt(0)
                                    End If


                                End If
                            End If

                            If (StSufmatch(0).Abrv = "ST" Or StSufmatch(0).Abrv = "AVE") And StSufmatch(0).Match.Index > 0 Then
                                ' use St or Ave if first
                                fnd = 0
                            Else
                                ' use last one
                                fnd = StSufmatch.Count - 1
                            End If

                            If StSufmatch(fnd).Abrv = "HWY" Then
                                rg = GetReg("((CA|STATE|U\.?S\.?)\s)?" & StSufmatch(fnd).Match.Value & "\s\d+")
                                m = rg.Match(prs)
                                If m.Success Then
                                    _Street = Left(prs, m.Index + Len(m.Value))
                                    prs = Mid(prs, m.Index + Len(m.Value) + 1)
                                End If
                            End If

                            If StSufmatch(fnd).Match.Index = 0 And prs <> StSufmatch(fnd).Match.Value And StSufmatch(fnd).Match.Value <> "AVE" Then
                                ' ignore: this includes "St" for "Saint", not "Street", and any street name beginning with a suffix like "PT VICENTE"
                                stIsSfx = True

                            ElseIf _Street = "" Then

                                ' save type by Abrv
                                rawSTsuf = StSufmatch(fnd).Match.Value
                                _StSuf = StSufmatch(fnd).Abrv

                                ' street name is everything before
                                _Street = RTrim(Left(prs, StSufmatch(fnd).Match.Index))

                                ' look at the rest
                                If Right(StSufmatch(fnd).Match.Value, 1) = "#" Then
                                    prs = Mid(prs, StSufmatch(fnd).Match.Index + Len(StSufmatch(fnd).Match.Value))
                                Else
                                    prs = Trim(Mid(prs, StSufmatch(fnd).Match.Index + Len(StSufmatch(fnd).Match.Value) + 1))
                                End If

                            End If  ' begins with "ST" 

                        Else  ' none found

                            ' look for really bad spellings and abreviations for "street", including run-ons with SUD

                            rg = GetReg("\bST((E[I|P])|(R[A|O|I])|A|ETS)")  ' none like this
                            If Not rg.IsMatch(prs) Then

                                rg = GetReg("\bST[^POUV]")
                                m = rg.Match(prs)

                                If m.Success And m.Index > 0 Then

                                    foPat = New FoundPatttern With {.Abrv = "ST", .Match = m}
                                    StSufmatch.Add(foPat)

                                    _StSuf = "ST"

                                    If SUD.Count = 0 Then

                                        ' look for SUD as run-on after "ST"
                                        prs = Left(prs, m.Index) & " " & Mid(prs, m.Index + 3)
                                        Call GetSUD(prs, rawSUD)

                                    End If

                                    prs = RTrim(Left(prs, m.Index))

                                End If
                            End If

                        End If

                        ' treat "AT" and "SR" as typo for "ST" if no other street type found

                        If StSufmatch.Count = 0 Then
                            rg = GetReg("(?:^|\b)\sAT(?:\s|\b|$)")
                            m = rg.Match(prs)
                            If Not m.Success Then
                                rg = GetReg("(?:^|\b)\sSR(?:\s|\b|$)")
                                m = rg.Match(prs)
                            End If
                            If m.Success Then
                                rawSTsuf = m.Value
                                _StSuf = "ST"
                                _Street = RTrim(Left(prs, m.Index))
                                prs = Mid(prs, m.Index + Len(m.Value) + 1)
                            End If
                        End If

                    End If  ' like "Avenue B"

                    If Len(prs) > 0 Then  ' anything after street suffix (or none found)

                        ' have to check street for "#" as well

                        If InStr(_Street, "#") > 0 Then

                            rng = LTrim(Mid(_Street, InStr(_Street, "#") + 1))
                            If InStr(rng, " ") > 0 Then rng = Left(rng, InStr(rng, " ") - 1)

                            thisSUD = New SUDtype With {.Range = Replace(rng, "#", "")}
                            _SUD.Add(thisSUD)

                            _Street = RTrim(Left(_Street, InStr(_Street, "#") - 1))
                            prs = ""

                        End If

                    End If  ' anything after street suffix

                    ' APO, FPO
                    rg = GetReg("\b(ap*o|fpo)(\b|$)")
                    If Len(prs) > 0 AndAlso rg.IsMatch(prs) And Len(_Street) = 0 Then
                        _rlst = BadAddress.PO_BOX

                    Else

                        If Len(_Street) = 0 And Len(prs) > 0 Then  ' no street type, no SUD

                            ' look for number at end of line

                            rg = GetReg("\s\d+($|\w{1})$")
                            m = rg.Match(prs)
                            If m.Success Then

                                thisSUD = New SUDtype With {.Range = LTrim(m.Value)}
                                _SUD.Add(thisSUD)

                                prs = RTrim(Left(prs, m.Index))

                            End If

                            _Street = prs
                            prs = ""

                        End If

                        If Len(prs) > 0 And SUD.Count = 0 And IsNumeric(prs) Then
                            thisSUD = New SUDtype With {.Range = prs}
                            _SUD.Add(thisSUD)
                            prs = ""
                        End If

                        If Len(_Street) = 0 Then
                            If Len(_HouseSuf) = 1 Then
                                _Street = _HouseSuf
                                _HouseSuf = ""
                            ElseIf Len(_PreDir) > 0 Then
                                _Street = _PreDir
                                _PreDir = ""
                            End If
                        End If

                    End If  ' APO, FPO

                    '   +-------------------+
                    '   |   predirectional  |
                    '   +-------------------+

                    rg = GetReg("\A(?:north|south)(?:\s*(?:east|west))?\b|\A(?:east|west|[NS]\s?[WE]?|[EW]|NO|SO)\b")
                    m = rg.Match(_Street)

                    ' do not use street name as predirectional, e.g., "West Ave"

                    If m.Success Then

                        If m.Value <> _Street Then
                            OK = True
                        ElseIf SUD.Count = 1 AndAlso (SUD(0).SUD = "" And IsNumeric(SUD(0).Range) And _StSuf = "") Then
                            OK = True
                        Else
                            OK = False
                        End If

                        If OK Then

                            rawPredir = m.Value

                            If m.Value = "SO" Then
                                _PreDir = "S"
                            ElseIf m.Value = "NO" And m.Index = 0 Then
                                _PreDir = "N"
                            ElseIf Len(m.Value) = 2 Then
                                _PreDir = m.Value
                            Else
                                _PreDir = m.Value
                                _PreDir = Replace(_PreDir, "NORTH", "N")
                                _PreDir = Replace(_PreDir, "SOUTH", "S")
                                _PreDir = Replace(_PreDir, "EAST", "E")
                                _PreDir = Replace(_PreDir, "WEST", "W")
                                _PreDir = Replace(_PreDir, " ", "")
                            End If

                            If m.Value <> _Street Then

                                ' remove predirectional
                                _Street = LTrim(Mid(_Street, Len(m.Value) + 1))
                            Else

                                ' "17 S 45" is "17 South 45th", not "17 South #45"
                                _Street = SUD(0).Range.ToString
                                SUD.RemoveAt(0)

                            End If

                        End If

                    End If

                    '   +-------------------+
                    '   |   postdirectional |
                    '   +-------------------+

                    prs = Trim(prs)
                    If Len(prs) > 0 Then  ' this is what was after street suffix and after SUD removed

                        rg = GetReg("\A(?:north|south)(?:\s*(?:east|west))?\b|\A(?:east|west|[NS][WE]?|[EW])$")
                        If rg.IsMatch(prs) Then
                            _PostDir = Trim(rg.Match(prs).Value)
                        ElseIf SUD.Count = 0 And Len(prs) <= 4 Then
                            thisSUD = New SUDtype With {.Range = prs}
                            _SUD.Add(thisSUD)
                            prs = ""
                        End If

                    End If

                    If Len(_PostDir) = 0 Then

                        rg = GetReg("\s(?:north|south)(?:\s*(?:east|west))?\b|\b(?:east|west|[NS][WE]?|[EW])(\b|$)")
                        m = rg.Match(_Street)
                        ' do not use street name as postdirectional, e.g., "West Ave"
                        If m.Success And m.Value <> _Street Then
                            _PostDir = LTrim(m.Value)
                            _Street = Left(_Street, m.Index)
                        End If

                    End If  ' and street suffix

                    If Len(_PostDir) > 0 Then
                        _PostDir = Replace(_PostDir, "NORTH", "N")
                        _PostDir = Replace(_PostDir, "SOUTH", "S")
                        _PostDir = Replace(_PostDir, "EAST", "E")
                        _PostDir = Replace(_PostDir, "WEST", "W")
                        _PostDir = Replace(_PostDir, " ", "")
                    End If

                End If  ' only one character after number

                If Left(_Street, 1) = "-" Then _Street = LTrim(Mid(_Street, 2))

                ' check for "7 STREET" and "9 AVENUE": use the number as the street name, not the house number

                If _HouseSuf = "" And _StSuf = "" And _PreDir = "" And _PostDir = "" And _SUD.Count = 0 Then

                    If _Street = "ST" Then
                        _Street = _HouseNum.ToString
                        rawSTsuf = rg.Match(_Street).Value
                        _StSuf = "ST"
                        _HouseNum = 0
                        If _rlst = BadAddress.NO_PROBLEM Then _rlst = BadAddress.NO_NUMBER

                    ElseIf _Street = "AVE" Then
                        _Street = _HouseNum.ToString
                        rawSTsuf = rg2.Match(_Street).NextMatch.Value
                        _StSuf = "AVE"
                        _HouseNum = 0
                        If _rlst = BadAddress.NO_PROBLEM Then _rlst = BadAddress.NO_NUMBER
                    End If

                End If

                If IsNumeric(_Street) Then
                    _Street = RTrim(_Street)
                    If Len(_Street) > 1 AndAlso Right(_Street, 2) Like "1?" Then
                        _Street &= "TH"
                    Else
                        Select Case Right(_Street, 1)
                            Case "1"
                                _Street &= "ST"
                            Case "2"
                                _Street &= "ND"
                            Case "3"
                                _Street &= "RD"
                            Case Else
                                _Street &= "TH"
                        End Select
                    End If
                End If

                ' count as successful if box number included but parsed anyway

                If Len(_Street) > 0 And _rlst = BadAddress.PO_BOX Then _rlst = BadAddress.NO_PROBLEM

                ' fix "B St"

                If Len(_HouseSuf) = 1 And (_Street = "ST" Or _Street = "AVE") Then
                    _StSuf = _Street
                    _Street = _HouseSuf
                    _HouseSuf = ""
                End If

            End If  ' not just number

        End If  ' has number

        ' remove "SFGH" and "T I" for Treasure Island"

        rg = GetReg("\b(?:(SFGH|PSF|SHELTER|HOMELESS|T ?I|)$|TREASURE ISLAND\w*)")
        m = rg.Match(_Street)
        If m.Success And _Street <> m.Value Then _Street = RTrim(Left(_Street, m.Index))

        rg = GetReg("[^\w]")

        If Trim(_Street) = "" Then

            If Len(_PreDir) > 0 Then
                _Street = rawPredir
                _PreDir = ""

            ElseIf _StSuf <> "" Then
                ' use raw suffix as street name
                _Street = rawSTsuf
                _StSuf = ""

            ElseIf SUD.Count = 1 Then
                If SUD(0).SUD = "PIER" Then
                    ' use pier as street name instead of SUD
                    _Street = SUD(0).SUD & " " & SUD(0).Range
                    SUD.RemoveAt(0)
                Else
                    If _lstStSuf.Exists(Function(p) p.Abrv = SUD(0).Range) Then
                        _Street = rawSUD
                        _StSuf = SUD(0).Range
                        SUD.RemoveAt(0)
                    End If
                End If
            End If

        ElseIf rg.IsMatch(Right(_Street, 1)) Then  ' not a letter
            _Street = RTrim(Left(_Street, Len(_Street) - 1))

        Else  ' street not missing

            _rlst = BadAddress.NO_PROBLEM

            ' look for single character at end of street, it cannot belong there

            If InStr(_Street, "#") > 0 Then _Street = Left(_Street, InStr(_Street, "#") - 1)

            If Not (stIsSfx Or Regex.IsMatch(_Street, "(?:^|\b)H(?:I|Y)?(?:GH?)?WA?Y(?:\b|$)")) Then

                rg = GetReg("\s\w$")
                m = rg.Match(_Street)
                While m.Success
                    If m.Value = " T" Then  ' treat as "ST"
                        If _StSuf = "" Then _StSuf = "ST"
                    Else  ' treat as range
                        thisSUD = New SUDtype With {.Range = LTrim(m.Value)}
                        _SUD.Add(thisSUD)
                    End If
                    _Street = RTrim(Left(_Street, Len(_Street) - 2))
                    m = rg.Match(_Street)
                End While

            End If

        End If

        If (_Street = "ST" Or _Street = "RD" Or _Street Like "AVE*") And _StSuf = "" And Len(_HouseSuf) = 1 Then
            _StSuf = Left(_Street, 3)
            _Street = _HouseSuf
            _HouseSuf = ""
        End If

        If _Street Like "* ST" Then _Street = Replace(_Street, " ST", "ST")

        If Left(_Street, 1) = "0" Then
            If Not IsNumeric(Mid(_Street, 2, 1)) Then
                _Street = "O" & Mid(_Street, 2)
            Else
                _Street = Mid(_Street, 2)
            End If
        End If

        If Trim(_Street) = "" And _rlst = BadAddress.NO_PROBLEM Then _rlst = BadAddress.NO_STREET

        ' look for number at beginning of street
        If UBound(Split(_Street, " ")) > 0 And IsNumeric(Split(_Street, " ")(0)) Then

            ' the only valid streets where the first word is a number are like "8 MILE RD"
            If Split(_Street, " ")(1) <> "MILE" Then
                If _rlst = BadAddress.NO_PROBLEM Then _rlst = BadAddress.TWO_NUMBERS
            End If

        End If

        ' try parsing after slash unless there is a date -- only up to 3 times
        rg = GetReg("\d{1,2}/\d{1,2}/\d{2,4}")
        If _rlst <> BadAddress.NO_PROBLEM And InStr(Replace(address, "C/O", "C O"), "/") > 0 And Not rg.IsMatch(address) And _recur < 3 Then
            _recur += 1
            Call Parse(LTrim(Mid(address, InStr(Replace(address, "C/O", "C O"), "/") + 1)))  ' ingore slash in C/O
        End If

        ' try parsing after comma -- only up to 3 times
        If _rlst <> BadAddress.NO_PROBLEM And InStr(address, ",") > 0 And _recur < 3 Then
            _recur += 1
            Call Parse(Mid(address, InStr(address, ",") + 1))
        End If

    End Sub

    ''' <summary>
    ''' Parser objects are equal if all componenets are the same, even if raw address is different
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    Public Overrides Function Equals(obj As Object) As Boolean
        Dim i As Integer
        If TypeOf obj Is AddressParser.Parser Then
            With CType(obj, AddressParser.Parser)
                If .HouseNum = Me.HouseNum And .HouseSuf = Me.HouseSuf And .PreDir = Me.PreDir And .Street = Me.Street And .StSuf = Me.StSuf And .PostDir = Me.PostDir Then
                    If .SUD.Count = Me.SUD.Count Then
                        For i = 1 To .SUD.Count
                            If Not .SUD(i - 1).Equals(Me.SUD(i - 1)) Then
                                Return False  ' different SUD
                            End If
                        Next
                    Else  ' different number of SUDs
                        Return False
                    End If
                Else  ' different component parts
                    Return False
                End If
            End With
        Else  ' not parser
            Return False
        End If
        Return True
    End Function

#End Region

#Region "-- private enum and structures --"

    ''' <summary>
    ''' Abbreviations to use for patterns to find.
    ''' </summary>
    ''' <remarks></remarks>
    Private Structure FindPatttern
        ''' <summary>
        ''' To identify the regex when re-calling it.
        ''' </summary>
        Public Abrv As String
        ''' <summary>
        ''' To identify th eregex by it pattern.
        ''' </summary>
        Public Pattern As String
        ''' <summary>
        ''' To indicate extra data, like if range applies.
        ''' </summary>
        ''' <remarks></remarks>
        Public Flag As Integer
        ''' <summary>
        ''' Exclude abbreviations that can be part of a name.
        ''' </summary>
        ''' <remarks></remarks>
        Public NoAbrv As Boolean
    End Structure

    ''' <summary>
    ''' Things found and their abbreviations.
    ''' </summary>
    ''' <remarks>Used in saving street types and SUDs out of address.</remarks>
    Private Class FoundPatttern
        ''' <summary>
        ''' Abbreviation to use for thing found.
        ''' </summary>
        ''' <remarks></remarks>
        Public Abrv As String
        ''' <summary>
        ''' Match object from RegEx that found this item.
        ''' </summary>
        ''' <remarks>Used to get string that matched and its position.</remarks>
        Public Match As Match
        ''' <summary>
        ''' Flag for refining match.
        ''' </summary>
        ''' <remarks>Used to identify whether SUD needs a range.</remarks>
        Public Flag As Integer
        ''' <summary>
        ''' Track whether range captured in 2nd group when finding SUDs.
        ''' </summary>
        ''' <remarks>You need the RegEx object itself (not the match object) to examine the group names.</remarks>
        Public HasRng As Boolean
    End Class

#End Region

#Region "-- private properties --"

    ''' <summary>
    ''' End of regular expression for SUDs that need a range.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const _patsfx_rng = "\s?#?\s?(?<RNG>\d*?[A-Z]?(?: |$)|\s)"   ' any number of digits (or none), followed by a single letter (or not)

    ''' <summary>
    ''' End of regular expression for SUDs that do not need a range.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const _patsfx_norng = ".?(?:\b|$)"

    ''' <summary>
    ''' Patterns for SUDs.
    ''' </summary>
    ''' <remarks></remarks>
    Private _lstSUD As List(Of FindPatttern)

    ''' <summary>
    ''' Patterns for street types.
    ''' </summary>
    ''' <remarks></remarks>
    Private _lstStSuf As List(Of FindPatttern)

    ''' <summary>
    ''' Dictionary of all RegEx objects already created.
    ''' </summary>
    ''' <remarks>So that we do not have to compile them repeatedly.</remarks>
    Private ReadOnly _dicReg As New Dictionary(Of String, Regex)

    ''' <summary>
    ''' To set option to compile all RegEx objects.
    ''' </summary>
    ''' <remarks></remarks>
    Private ReadOnly _regOpt As RegexOptions = RegexOptions.IgnoreCase

    ''' <summary>
    ''' Used to limit recursion in Parse method.
    ''' </summary>
    Private _recur As Integer

#End Region

#Region "-- private methods --"

    ''' <summary>
    ''' Set up SUD patterns.
    ''' </summary>
    ''' <remarks>For flag, 1 = range, 0 = no range.</remarks>
    Private Sub Setup_lstSUD()

        ' use consistent patterns for SUD that can have a range

        ' OK if followed by number of any length, single letter, or space after optional "#"
        ' save what follws as RNG group so we can use it as the range, e.g., "RM123"

        Dim f As FindPatttern
        Dim fg As Integer

        _lstSUD = New List(Of FindPatttern)

        fg = 1  ' SUDs that have range

        f = New FindPatttern
        With f : .Abrv = "APT" : .Pattern = "(?:^|\b)AP(?:P|A?R)?(?:T|ARTMENT)" : .Flag = fg : End With  ' was: "(?:^|\b)AP(P|AR)?(?:T|ARTMENT)(\b|^S)" 
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BLDG" : .Pattern = "((?:^|\b)(?:B(?:UI)?LD(?:IN)?G?)|BLDG)" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "DEPT" : .Pattern = "(?:^|\b)DEP(?:ARTMEN)?T" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FL" : .Pattern = "(?:^|\b)FL(?:OOR)?" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "HNGR" : .Pattern = "(?:^|\b)HA?NGE?R" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        ' this is tricky, need to handle "LOTUS" not as "LOT S"
        With f : .Abrv = "LOT" : .Pattern = "(?:^|\b)LOT(\s?)" : .Flag = fg : End With  '' & "#?(?<RNG>\d+|[A-Z]( |$)|\s)" 
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PIER" : .Pattern = "(?:^|\b)PIER(?:[^\w]|\s|$)" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RM" : .Pattern = "(?:^|\b)R(?:OO)?M" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SLIP" : .Pattern = "(?:^|\b)SLIP" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "STOP" : .Pattern = "(?:^|\b)STOP" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "STE" : .Pattern = "(?:^|\b)(?:S(?:UI)?T+E|SUIT)\b" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "TRLR" : .Pattern = "(?:^|\b)TR(?:AI)?LE?R" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "UNIT" : .Pattern = "(?:^|\b)UN(?:IT[^EY]|\b)" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PH  " : .Pattern = "(?:^|\b)P(?:ENT)?H(?:OUSE)?" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BED" : .Pattern = "(?:^|\b)BED[^\w]\b?" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "#" : .Pattern = "(?:^|\b)NO\.?\s?\d" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "DORM" : .Pattern = "(?:^|\b)DORM\b" : .Flag = fg : End With
        _lstSUD.Add(f)

        f = New FindPatttern
        With f : .Abrv = "FLAT" : .Pattern = "(?:^|\b)FLA?T[^S](?:\b|$)" : .Flag = fg : End With
        _lstSUD.Add(f)

        fg = 0  ' SUDs that have no range

        ' NB: "Floor" has no range because we do not expect "Floor 3", and we handle "3rd floor" separately
        f = New FindPatttern
        With f : .Abrv = "FL" : .Pattern = "(?:^|\b)(?<RNG>\d+(?<SFX>ST|ND|RD|TH))\sFL(?:OOR)?" : .Flag = fg : End With
        _lstSUD.Add(f)

        f = New FindPatttern
        With f : .Abrv = "BSMT" : .Pattern = "(?:^|\b)BA?SE?M(?:EN)?T" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LBBY" : .Pattern = "(?:^|\b)LO?BBY" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LOWR" : .Pattern = "(?:^|\b)LOWE?R(\b|$|[^\w])" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "OFC" : .Pattern = "(?:^|\b)OF(?:FI)?CE?" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "REAR" : .Pattern = "(?:^|\b)REAR" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "UPPR" : .Pattern = "(?:^|\b)UPPE?R" : .Flag = fg : End With

        Exit Sub

        ' NB: these do NOT get used, they work but they cause problems with San Francisco street names:

        f = New FindPatttern
        With f : .Abrv = "SIDE" : .Pattern = "(?:^|\b)SIDE" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "KEY" : .Pattern = "(?:^|\b)KEY" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FRNT" : .Pattern = "(?:^|\b)FRO?NT" : .Flag = fg : End With
        _lstSUD.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SPC" : .Pattern = "(?:^|\b)SPA?CE?" : .Flag = fg : End With
        _lstSUD.Add(f)

    End Sub

    ''' <summary>
    ''' Set up patterns to find street types.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Setup_lstStSuf()

        Dim f As FindPatttern

        _lstStSuf = New List(Of FindPatttern)

        f = New FindPatttern
        With f : .Abrv = "ALY" : .Pattern = "(?:^|\b)AL+E*Y?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "ANX" : .Pattern = "(?:^|\b)AN+E*X\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "ARC" : .Pattern = "(?:^|\b)ARC(?:ADE)?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "AVE" : .Pattern = "(?:^|\b)(?:AVE?(?:N(?:U(?:E)?)?)?\.?)|\b(?:A(?:\w)?VE)(?:\b|$)" : End With
        ''_lstStSuf.Add(f)  do not use, it matches "avalon"
        f = New FindPatttern
        With f : .Abrv = "BYU" : .Pattern = "(?:^|\b)(?:BYU|BA?YO*[OU])\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BCH" : .Pattern = "(?:^|\b)B(?:EA)?CH\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BND" : .Pattern = "(?:^|\b)BE?ND\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BLF" : .Pattern = "(?:^|\b)BLU?F+[^S]?(?:\b|$)" : End With  'was: .Pattern = "(?:^|\b)BLU?F+[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BLFS" : .Pattern = "(?:^|\b)BLU?F+S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BTM" : .Pattern = "(?:^|\b)B(?:O?T+O?M|OT)\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BLVD" : .Pattern = "(?:^|\b)(?:B(?:(?:OU)?LE?V(?:AR)?D|OULV?)\.?|(BLDV|BVLD|\wVLD|BLV?E?|BLUD(?:\w+)?))(?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)B(?:(?:OU)?LE?V(?:AR)?D|OULV?)\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BR" : .Pattern = "(?:^|\b)BR(?:(?:A?NCH)|\.?)(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BRG" : .Pattern = "(?:^|\b)BRI?D?GE?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BRK" : .Pattern = "(?:^|\b)BRO*K[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)BRO*K[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BRKS" : .Pattern = "(?:^|\b)BRO*KS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BG" : .Pattern = "(?:^|\b)B(?:UR)?G[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)B(?:UR)?G[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BGS" : .Pattern = "(?:^|\b)B(?:UR)?GS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "BYP" : .Pattern = "(?:^|\b)BYP(?:A?S*)?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CP" : .Pattern = "(?:^|\b)CA?MP[^EO](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)CA?M?P[^E]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CYN" : .Pattern = "(?:^|\b)CA?N?YO?N\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CPE" : .Pattern = "(?:^|\b)CA?PE\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CSWY" : .Pattern = "(?:^|\b)C(?:AU)?SE?WA?Y\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CTR" : .Pattern = "(?:^|\b)CTR(?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)C(?:E?N?TE?RE?|ENT)[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CTRS" : .Pattern = "(?:^|\b)C(?:E?N?TE?RE?|ENT)S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CIR" : .Pattern = "(?:^|\b)C(?:RCLE?|IRC?L?E?)[^S]?(?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)C(?:RCLE?|IRC?L?E?)[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CIRS" : .Pattern = "(?:^|\b)C(?:RCLE?|IRC?L?E?)S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CLF" : .Pattern = "(?:^|\b)CLI?F+[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)CLI?F+[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CLFS" : .Pattern = "(?:^|\b)CLI?F+S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CMN" : .Pattern = "(?:^|\b)CO?M+O?N\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "COR" : .Pattern = "^(?!.*CORP)(?:^|\b)COR(?:NER)?[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)COR(?:NER)?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CORS" : .Pattern = "^(?!.*CORP)(?:^|\b)COR(?:NER)?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CRSE" : .Pattern = "(?:^|\b)C(?:OU)?RSE\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CT" : .Pattern = "(?:^|\b)(?:CORT|C(?:OU)?R?T[^RS]?)(?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)C(?:OU)?R?T[^RS]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CTS" : .Pattern = "(?:^|\b)C(?:OU)?R?TS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CV" : .Pattern = "^(?!.*COVER)(?:^|\b)CO?VE?[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)CO?VE?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CVS" : .Pattern = "^(?!.*COVER)(?:^|\b)CO?VE?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CRK" : .Pattern = "(?:^|\b)C(?:RE*K|[RK])\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CRES" : .Pattern = "(?:^|\b)CR(?:ES?C?E?(?:NT)?|SC?E?NT|R[ES])\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CRST" : .Pattern = "(?:^|\b)CRE?ST\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "XING" : .Pattern = "(?:^|\b)(?:CRO?S+I?NG|XING)\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "XRD" : .Pattern = "(?:^|\b)(?:CRO?S+R(?:OA)?D|XR(?:OA)?D)\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "CURV" : .Pattern = "(?:^|\b)CURVE?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "DL" : .Pattern = "(?:^|\b)DA?LE?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "DM" : .Pattern = "(?:^|\b)DA?M\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "DV" : .Pattern = "(?:^|\b)DI?V(?:I?DE?)?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "DR" : .Pattern = "(?:^|\b)DR(?:I?VE?)?[^S]?(?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)DR(?:I?VE?)?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "DRS" : .Pattern = "(?:^|\b)DR(?:I?VE?)?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "EST" : .Pattern = "(?:^|\b)EST(?:ATE)?[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)EST(?:ATE)?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "ESTS" : .Pattern = "(?:^|\b)EST(?:ATE)?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "EXPY" : .Pattern = "(?:^|\b)EXP(?:R(?:ES+(?:WAY)?)?|[WY])?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "EXT" : .Pattern = "(?:^|\b)EXT(?:E?N(?:S(?:IO)?N)?)?[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)EXT(?:E?N(?:S(?:IO)?N)?)?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "EXTS" : .Pattern = "(?:^|\b)EXT(?:E?N(?:S(?:IO)?N)?)?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FALL" : .Pattern = "(?:^|\b)FALL[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)FALL[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FLS" : .Pattern = "(?:^|\b)FA?L+S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FRY" : .Pattern = "(?:^|\b)FE?R+Y\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FLD" : .Pattern = "(?:^|\b)F(?:IE)?LD[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)F(?:IE)?LD[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FLDS" : .Pattern = "(?:^|\b)F(?:IE)?LDS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FLTS" : .Pattern = "(?:^|\b)FLA?TS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FRD" : .Pattern = "(?:^|\b)FO?RD[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)FO?RD[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FRDS" : .Pattern = "(?:^|\b)FO?RDS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FRG" : .Pattern = "(?:^|\b)FO?RGE?[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)FO?RGE?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FRGS" : .Pattern = "(?:^|\b)FO?RGE?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FRK" : .Pattern = "(?:^|\b)FO?RK[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)FO?RK[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FRKS" : .Pattern = "(?:^|\b)FO?RKS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FWY" : .Pattern = "(?:^|\b)F(?:RE*)?WA?Y\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "GDN" : .Pattern = "(?:^|\b)G(?:A?R)?DE?N[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)G(?:A?R)?DE?N[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "GDNS" : .Pattern = "(?:^|\b)G(?:A?R)?DE?NS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "GTWY" : .Pattern = "(?:^|\b)GA?TE?WA?Y\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "GLN" : .Pattern = "(?:^|\b)GLE?N[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)GLE?N[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "GLNS" : .Pattern = "(?:^|\b)GLE?NS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "GRN" : .Pattern = "(?:^|\b)GRE*N[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)GRE*N[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "GRNS" : .Pattern = "(?:^|\b)GRE*NS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "GRV" : .Pattern = "(?:^|\b)GRO?VE?[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)GRO?VE?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "GRVS" : .Pattern = "(?:^|\b)GRO?VE?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "HBR" : .Pattern = "(?:^|\b)H(?:(?:A?R)?BO?R|ARB)[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)H(?:(?:A?R)?BO?R|ARB)[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "HBRS" : .Pattern = "(?:^|\b)H(?:A?R)?BO?RS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "HVN" : .Pattern = "(?:^|\b)HA?VE?N\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "HTS" : .Pattern = "(?:^|\b)H(?:(?:EI)?GH?)?TS?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "HWY" : .Pattern = "(?:^|\b)H(?:I|Y)?(?:GH?)?WA?Y\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "HL" : .Pattern = "(?:^|\b)HI?L+[^A-Z](?:\b|$)" : End With  ' .Pattern = "(?:^|\b)HI?L+[^A-Z]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "HLS" : .Pattern = "(?:^|\b)HI?L+S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "HOLW" : .Pattern = "(?:^|\b)HO?L+O?WS?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "INLT" : .Pattern = "(?:^|\b)INLE?T\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "ISLE" : .Pattern = "(?:^|\b)ISLES?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "JCT" : .Pattern = "(?:^|\b)JU?CT(?:(?:IO)?N)?[^S](?:\b|$)" : End With  ' .Pattern = "(?:^|\b)JU?CT(?:(?:IO)?N)?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "JCTS" : .Pattern = "(?:^|\b)JU?CT(?:(?:IO)?N)?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "KY" : .Pattern = "(?:^|\b)KE?Y[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)KE?Y[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "KYS" : .Pattern = "(?:^|\b)KE?YS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "KNL" : .Pattern = "(?:^|\b)KNO?L+[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)KNO?L+[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "KNLS" : .Pattern = "(?:^|\b)KNO?L+S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LAND" : .Pattern = "(?:^|\b)LAND[^A-Z](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)LAND[^A-Z]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LNDG" : .Pattern = "(?:^|\b)LA?ND(?:I?N)?G\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LN" : .Pattern = "(?:^|\b)L(?:A?NES?|N)\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LGT" : .Pattern = "(?:^|\b)LI?GH?T\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LGTS" : .Pattern = "(?:^|\b)LI?GH?TS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LF" : .Pattern = "(?:^|\b)L(?:OA)?F\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LCK" : .Pattern = "(?:^|\b)LO?CK[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)LO?CK[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LCKS" : .Pattern = "(?:^|\b)LO?CKS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LDG" : .Pattern = "(?:^|\b)LO?DGE?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LOOP" : .Pattern = "(?:^|\b)LOOPS?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "MALL" : .Pattern = "(?:^|\b)MALL\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "MNR" : .Pattern = "(?:^|\b)MA?NO?R[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)MA?NO?R[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "MNRS" : .Pattern = "(?:^|\b)MA?NO?RS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "MDW" : .Pattern = "(?:^|\b)M(?:EA?)?DO?W[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)M(?:EA?)?DO?W[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "MDWS" : .Pattern = "(?:^|\b)M(?:EA?)?DO?WS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "MEWS" : .Pattern = "(?:^|\b)MEWS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "ML" : .Pattern = "(?:^|\b)MI?L+[^KS](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)MI?L+[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "MLS" : .Pattern = "(?:^|\b)MI?L+S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "MTWY" : .Pattern = "(?:^|\b)MO?T(?:OR)?WA?Y\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "MTN" : .Pattern = "(?:^|\b)M(?:OU)?N?T(?:AI?|I)?N[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)M(?:OU)?N?T(?:AI?|I)?N[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "MTNS" : .Pattern = "(?:^|\b)M(?:OU)?N?T(?:AI?|I)?NS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "NCK" : .Pattern = "(?:^|\b)NE?CK\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "ORCH" : .Pattern = "(?:^|\b)ORCH(?:A?RD)?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "OVAL" : .Pattern = "(?:^|\b)OVA?L\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "OPAS" : .Pattern = "(?:^|\b)O(?:VER)?PAS+\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PARK" : .Pattern = "(?:^|\b)(?:PA?R?K[^A-Z]|PK)(?:\s|$)" : End With  ' was: .Pattern = "(?:^|\b)PA?R?K[^A-RT-Z]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PKWY" : .Pattern = "(?:^|\b)PA?R?K-?W?A?YS?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PASS" : .Pattern = "(?:^|\b)PASS[^A-Z](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)PASS[^A-Z]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PSGE" : .Pattern = "(?:^|\b)PA?S+A?GE\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PATH" : .Pattern = "(?:^|\b)PATHS?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PIKE" : .Pattern = "(?:^|\b)PIKES?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PNE" : .Pattern = "(?:^|\b)PI?NE[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)PI?NE[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PNES" : .Pattern = "(?:^|\b)PI?NES\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PL" : .Pattern = "(?:^|\b)PL(?:ACE)?[^A-Z]?(?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)PL(?:ACE)?[^A-Z]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PLN" : .Pattern = "(?:^|\b)PL(?:AI)?N[^ES](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)PL(?:AI)?N[^ES]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PLNS" : .Pattern = "(?:^|\b)PL(?:AI)?NE?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PLZ" : .Pattern = "(?:^|\b)(?<!LA )PLA?ZA?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PT" : .Pattern = "(?:^|\b)P(?:OI)?N?T[^SE]*\.?(?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)P(?:OI)?N?T[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PTS" : .Pattern = "(?:^|\b)P(?:OI)?N?TS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PRT" : .Pattern = "(?:^|\b)PO?RT(?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)PO?RT[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PRTS" : .Pattern = "(?:^|\b)PO?RTS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "PR" : .Pattern = "(?:^|\b)PR(?:(?:AI?)?R(?:IE)?|[^KT]?)?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RADL" : .Pattern = "(?:^|\b)RAD(?:I[AE]?)?L?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RAMP" : .Pattern = "(?:^|\b)RAMP\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RNCH" : .Pattern = "(?:^|\b)RA?NCH(?:E?S)?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RPD" : .Pattern = "(?:^|\b)RA?PI?D[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)RA?PI?D[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RPDS" : .Pattern = "(?:^|\b)RA?PI?DS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RST" : .Pattern = "(?:^|\b)RE?ST\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RDG" : .Pattern = "(?:^|\b)RI?DGE?[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)RI?DGE?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RDGS" : .Pattern = "(?:^|\b)RI?DGE?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RIV" : .Pattern = "(?:^|\b)RI?VE?R?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RD" : .Pattern = "(?:^|\b)R(?:OA)?D(?:\b|[^A-Z])(?:\b|$)(?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)R(?:OA)?D[^A-Z]*\.?(?:\b|$)(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RDS" : .Pattern = "(?:^|\b)R(?:OA)?DS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RTE" : .Pattern = "(?:^|\b)R(?:OU)?TE\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "ROW" : .Pattern = "(?:^|\b)ROW\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RUE" : .Pattern = "(?:^|\b)RUE\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "RUN" : .Pattern = "(?:^|\b)RUN\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SHL" : .Pattern = "(?:^|\b)SH(?:OA)?L[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)SH(?:OA)?L[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SHLS" : .Pattern = "(?:^|\b)SH(?:OA)?LS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SHRS" : .Pattern = "(?:^|\b)SH(?:OA?)?RE?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SKWY" : .Pattern = "(?:^|\b)SKY?W?A?YS?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SPG" : .Pattern = "(?:^|\b)SP(?:RI?)?N?G[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)SP(?:RI?)?N?G[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SPGS" : .Pattern = "(?:^|\b)SP(?:RI?)?N?GS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SPUR" : .Pattern = "(?:^|\b)SPURS?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SQ" : .Pattern = "(?:^|\b)SQU?A?R?E?[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)SQU?A?R?E?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SQS" : .Pattern = "(?:^|\b)SQU?A?R?E?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "STA" : .Pattern = "(?:^|\b)ST(?:N|AT?(?:IO)?N?)\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "STRA" : .Pattern = "(?:^|\b)STR(?:VN|AV?E?N?U?E?)\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "STRM" : .Pattern = "(?:^|\b)STRE?A?ME?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "ST" : .Pattern = "(?:^|\b)ST(?:\.|R(?:E+)?T?\.?)?(?<RNG>\d+)?(?:\b|$)" : .NoAbrv = True : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "STS" : .Pattern = "(?:^|\b)STR?E*T?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SMT" : .Pattern = "(?:^|\b)SU?M+I?T+\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "TER" : .Pattern = "(?:^|\b)TER(?:R?(?:(A|E)(?:N)?CE)?)?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "TRWY" : .Pattern = "(?:^|\b)TH?R(?:OUGH)?WA?Y\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "TRCE" : .Pattern = "(?:^|\b)TRA?CES?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "TRAK" : .Pattern = "(?:^|\b)TRA?C?KS?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "TRFY" : .Pattern = "(?:^|\b)TRA?F(?:FICWA)?Y\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "TRL" : .Pattern = "(?:^|\b)TR(?:(?:AI)?LS?\.?|\.?)(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "TUNL" : .Pattern = "(?:^|\b)TUNN?E?LS?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "TPKE" : .Pattern = "(?:^|\b)T(?:U?RN?)?PI?KE?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "UPAS" : .Pattern = "(?:^|\b)U(?:NDER)?PA?SS?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "UNS" : .Pattern = "(?:^|\b)U(?:NIO)?NS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "VIA" : .Pattern = "(?:^|\b)V(?:IA)?(?:DU?CT)\.?(?:\b|$)" : .NoAbrv = True : End With  ' was: .Pattern = "(?:^|\b)V(?:IA)?(?:DU?CT)?\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "VLG" : .Pattern = "(?:^|\b)V(?:LG|ILL(?:I?AGE?)?)[^AES]?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        ' does not end in "AS" like Villas"
        f = New FindPatttern
        With f : .Abrv = "VLGS" : .Pattern = "^(?!.*(AS)$)(?:^|\b)V(?:LG|ILL(?:I?AGE?)?)[^E]?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "VL" : .Pattern = "(?:^|\b)V(?:L[^CGLY]*|I?LL?E)\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "VIS" : .Pattern = "(?:^|\b)VI?S(?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)VI?S(?:TA?)?\.?(?:\b|$)" leave "Vista" out of it ...
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "WALK" : .Pattern = "(?:^|\b)WALKS?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "WALL" : .Pattern = "(?:^|\b)WALL\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "WAY" : .Pattern = "(?:^|\s)WA?Y[^S]?(?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)WA?Y[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "WAYS" : .Pattern = "(?:^|\b)WA?YS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "WL" : .Pattern = "(?:^|\b)WE?LL?[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)WE?LL?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "WLS" : .Pattern = "(?:^|\b)WE?LL?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)

        f = New FindPatttern
        With f : .Abrv = "CMNS" : .Pattern = "(?:^|\b)CO?M+O?NS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)

        ' added for misspellings of Avenue

        f = New FindPatttern
        With f
            .Abrv = "AVE"
            .Pattern = "(?:^|\b)(AV(\b|$)|AV[EN](\b|\w+))"
            ' was:      .Pattern = "(?:^|\b)A(?:[^L])?V(?:[^IA])(?:\w+)?(?:\b|$)"
        End With
        _lstStSuf.Add(f)

        ' "SVE"
        f = New FindPatttern
        With f : .Abrv = "AVE" : .Pattern = "(?:^|\b)(?:SVE|AEV)(?:\w+)?(?:\b|$)" : End With
        _lstStSuf.Add(f)

        ' added for misspellings of Blvd

        f = New FindPatttern
        With f : .Abrv = "BLVD" : .Pattern = "(?:^|\b)BOULE(?:\w+)?(?:\b|$)" : End With
        _lstStSuf.Add(f)

        f = New FindPatttern
        ' do not match if it includes "M" (like "STADIUM DR") or "O"
        With f : .Abrv = "ST" : .Pattern = "^(?!.*[MN\s])(?:^|\b)S(?:T|R).*?(?:R|T)(?:E)?(?:\b|$)" : End With
        _lstStSuf.Add(f)

        f = New FindPatttern
        With f : .Abrv = "ST" : .Pattern = "(?:^|\b)SST|\bTREET(?:\b|$)" : End With
        _lstStSuf.Add(f)

        f = New FindPatttern
        With f : .Abrv = "VW" : .Pattern = "(?:^|\b)VW(?:\b|$)" : End With
        _lstStSuf.Add(f)

        f = New FindPatttern
        With f : .Abrv = "VLY" : .Pattern = "(?:^|\b)VLY(?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)VA?LL?E?Y[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)

        Exit Sub

        ' don't use these:

        f = New FindPatttern
        With f : .Abrv = "UN" : .Pattern = "(?:^|\b)U(?:NIO)?N[^IS](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)U(?:NIO)?N[^IS]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)

        f = New FindPatttern
        With f : .Abrv = "CLB" : .Pattern = "(?:^|\b)CLU?B\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)

        f = New FindPatttern
        With f : .Abrv = "MSN" : .Pattern = "(?:^|\b)MI?S+(?:IO)?N\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)

        f = New FindPatttern
        With f : .Abrv = "MT" : .Pattern = "(?:^|\b)M(?:OU)?N?T[^A-Z](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)M(?:OU)?N?T[^A-Z]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)

        f = New FindPatttern
        With f : .Abrv = "LK" : .Pattern = "(?:^|\b)LA?KE?[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)LA?KE?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "LKS" : .Pattern = "(?:^|\b)LA?KE?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)

        f = New FindPatttern
        With f : .Abrv = "VW" : .Pattern = "(?:^|\b)V(?:IE)?W[^S]?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "VWS" : .Pattern = "(?:^|\b)V(?:IE)?WS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "VLY" : .Pattern = "(?:^|\b)VA?LL?E?Y[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)VA?LL?E?Y[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "VLYS" : .Pattern = "(?:^|\b)VA?LL?E?YS\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FT" : .Pattern = "(?:^|\b)FO?R?T[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)FO?R?T[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "SHR" : .Pattern = "(?:^|\b)SH(?:OA?)?RE?[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)SH(?:OA?)?RE?[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "IS" : .Pattern = "(?:^|\b)IS(?:LA?ND)?[^A-Z](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)IS(?:LA?ND)?[^A-Z]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "ISS" : .Pattern = "(?:^|\b)IS(?:LA?ND)?S\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FLT" : .Pattern = "(?:^|\b)FLA?T[^S](?:\b|$)" : End With  ' was: .Pattern = "(?:^|\b)FLA?T[^S]*\.?(?:\b|$)" 
        _lstStSuf.Add(f)
        f = New FindPatttern
        With f : .Abrv = "FRST" : .Pattern = "(?:^|\b)FO?RE?STS?\.?(?:\b|$)" : End With
        _lstStSuf.Add(f)

    End Sub

    ''' <summary>
    ''' Erase values.
    ''' </summary>
    ''' <remarks>Run at beginning each time Parse is run.</remarks>
    Private Sub Reset()
        _HouseNum = ""
        _HouseSuf = ""
        _PreDir = ""
        _Street = ""
        _StSuf = ""
        _PostDir = ""
        _SUD = New List(Of SUDtype)
        _recur = 0
    End Sub

    ''' <summary>
    ''' Maintain dictionary of RegEx objects by their patterns.
    ''' </summary>
    ''' <param name="pattern"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Allow us to great each RegEx once and re-use same pattern.
    ''' </remarks>
    Private Function GetReg(ByVal pattern As String) As Regex
        Dim r As Regex
        If _dicReg.ContainsKey(pattern) Then
            r = _dicReg.Item(pattern)
        Else
            r = New Regex(pattern, _regOpt)
            _dicReg.Add(pattern, r)
        End If
        Return r
    End Function

    ''' <summary>
    ''' Save SUD and range, remove from address.
    ''' </summary>
    ''' <param name="prs">String being parsed.</param>
    ''' <param name="rawSUD">SUD to find.</param>
    Private Sub GetSUD(ByRef prs As String, ByRef rawSUD As String)

        ' for parsing with regular expressons
        Dim m As Match
        Dim mc As MatchCollection
        Dim rg As Regex

        Dim fnd As Integer                  ' match found to use
        Dim thisSUD As SUDtype              ' SUD to add
        Dim rng As String                   ' range to add

        Dim foPat As FoundPatttern                  ' for saving patterns found
        Dim fiPat As FindPatttern                   ' for looping thorugh patterns to find
        Dim SUDmatch As New List(Of FoundPatttern)
        Dim i As Integer

        SUDmatch = New List(Of FoundPatttern)

        prs = Regex.Replace(prs, "#\s", "#")

        For Each fiPat In _lstSUD

            If fiPat.Flag = 1 Then
                rg = GetReg(fiPat.Pattern & _patsfx_rng)
            Else
                rg = GetReg(fiPat.Pattern & _patsfx_norng)
            End If

            mc = rg.Matches(prs)
            For Each m In mc
                If m.Success Then
                    foPat = New FoundPatttern With {.Abrv = fiPat.Abrv, .Match = m, .Flag = fiPat.Flag}
                    If m.Groups.Count = 2 AndAlso rg.GetGroupNames(1) = "RNG" Then foPat.HasRng = True
                    SUDmatch.Add(foPat)
                End If
            Next

            If fiPat.Flag = 1 And mc.Count = 0 Then

                rg = GetReg(fiPat.Pattern & _patsfx_norng)

                mc = rg.Matches(prs)
                For Each m In mc
                    If m.Success Then
                        foPat = New FoundPatttern With {.Abrv = fiPat.Abrv, .Match = m, .Flag = fiPat.Flag}
                        SUDmatch.Add(foPat)
                    End If
                Next

            End If

        Next

        ' sort by order in string
        If SUDmatch.Count > 1 Then SUDmatch = SUDmatch.OrderBy(Function(x) x.Match.Index).ToList

        ' ignore "Lower" if part of street name
        If SUDmatch.Count > 0 Then
            If (SUDmatch(0).Match.Value = "LOWER " Or SUDmatch(0).Match.Value = "UN ") And SUDmatch(0).Match.Index = 0 Then SUDmatch.RemoveAt(0)
        End If

        If SUDmatch.Count > 0 Then

            For fnd = 1 To SUDmatch.Count - 1

                ' find range for each one from text before the next, e.g. "floor 2 room 5"

                If SUDmatch(fnd - 1).Match.Groups("RNG").Value = " " Then
                    rng = Mid(prs, SUDmatch(fnd - 1).Match.Index + 1, SUDmatch(fnd).Match.Index - SUDmatch(fnd - 1).Match.Index)
                    rng = Trim(Mid(rng, Len(SUDmatch(fnd - 1).Match.Value) + 1))
                    If InStr(rng, " ") > 0 Then rng = Left(rng, InStr(rng, " ") - 1)
                Else
                    ' remove any ordinal suffix before saving range, e.g., "2ND"
                    rng = Replace(SUDmatch(fnd - 1).Match.Groups("RNG").Value, SUDmatch(fnd - 1).Match.Groups("SFX").Value, "")
                End If

                ' remove anything after range if present
                If Len(rng) > 0 Then
                    If Mid(prs, SUDmatch(fnd - 1).Match.Groups("RNG").Index + Len(SUDmatch(fnd - 1).Match.Groups("RNG").Value) + 1, 1) <> " " Then
                        i = InStr(Mid(prs, SUDmatch(fnd - 1).Match.Groups("RNG").Index + Len(SUDmatch(fnd - 1).Match.Groups("RNG").Value) + 1), " ")
                        If i = 0 Then
                            prs = Left(prs, SUDmatch(fnd - 1).Match.Groups("RNG").Index + Len(SUDmatch(fnd - 1).Match.Groups("RNG").Value))
                        Else
                            prs = Left(prs, SUDmatch(fnd - 1).Match.Groups("RNG").Index + Len(SUDmatch(fnd - 1).Match.Groups("RNG").Value)) & Mid(prs, SUDmatch(fnd - 1).Match.Groups("RNG").Index + Len(SUDmatch(fnd - 1).Match.Groups("RNG").Value) + i)
                        End If
                    End If
                End If

                ' remove SUD and range
                prs = Left(prs, SUDmatch(0).Match.Index) + LTrim(Mid(prs, SUDmatch(0).Match.Index + Len(SUDmatch(0).Match.Value) + Len(rng) + 1))

                If Not (SUDmatch(fnd - 1).Flag = 1 And Len(rng) = 0) Then

                    ' if no matches, look for SUD missing range so it can be removed

                    thisSUD = New SUDtype With {.SUD = SUDmatch(fnd - 1).Abrv, .Range = Replace(rng, "#", "")}
                    _SUD.Add(thisSUD)

                End If

            Next

            ' one at the end (or only one)

            If SUDmatch(fnd - 1).Match.Groups("RNG").Value = " " Then
                rng = Mid(prs, SUDmatch(fnd - 1).Match.Index + 1)
                rng = Trim(Mid(rng, Len(SUDmatch(fnd - 1).Match.Value) + 1))
                If InStr(rng, " ") > 0 Then rng = Left(rng, InStr(rng, " ") - 1)
            Else
                rng = SUDmatch(fnd - 1).Match.Groups("RNG").Value
            End If

            ' remove anything after range if present, including ordinal suffix 
            If Len(rng) > 0 Then
                If Mid(prs, SUDmatch(fnd - 1).Match.Groups("RNG").Index + Len(SUDmatch(fnd - 1).Match.Groups("RNG").Value) + 1, 1) <> " " Then
                    i = InStr(Mid(prs, SUDmatch(fnd - 1).Match.Groups("RNG").Index + Len(SUDmatch(fnd - 1).Match.Groups("RNG").Value) + 1), " ")
                    If i = 0 Then
                        prs = Left(prs, SUDmatch(fnd - 1).Match.Groups("RNG").Index + Len(SUDmatch(fnd - 1).Match.Groups("RNG").Value))
                    Else
                        prs = Left(prs, SUDmatch(fnd - 1).Match.Groups("RNG").Index + Len(SUDmatch(fnd - 1).Match.Groups("RNG").Value)) & Mid(prs, SUDmatch(fnd - 1).Match.Groups("RNG").Index + Len(SUDmatch(fnd - 1).Match.Groups("RNG").Value) + i)
                    End If
                End If
            End If

            ' remove SUD and range
            prs = Left(prs, SUDmatch(0).Match.Index) + LTrim(Mid(prs, SUDmatch(0).Match.Index + Len(SUDmatch(0).Match.Value) + 1))

            If Not (SUDmatch(fnd - 1).Flag = 1 And Len(rng) = 0) Then
                thisSUD = New SUDtype With {.SUD = SUDmatch(fnd - 1).Abrv, .Range = Replace(rng, "#", "")}
                _SUD.Add(thisSUD)
                rawSUD = SUDmatch(fnd - 1).Match.Value
            End If

        End If  ' any SUD found

    End Sub

    ''' <summary>
    ''' Name for direction abbreviation.
    ''' </summary>
    ''' <param name="thisdir"></param>
    ''' <returns></returns>
    Private Function DirName(ByVal thisdir As String) As String

        Dim d As String = ""

        If Len(thisdir) > 0 Then
            Select Case Left(thisdir, 1)
                Case "S"
                    d = "SOUTH"
                Case "N"
                    d = "NORTH"
                Case "E"
                    d = "EAST"
                Case "W"
                    d = "WEST"
            End Select
        End If

        If Len(thisdir) = 2 Then
            Select Case Right(thisdir, 1)
                Case "E"
                    d &= "EAST"
                Case "W"
                    d &= "WEST"
            End Select
        End If

        Return d

    End Function

#End Region

#Region "-- constructors --"

    ''' <summary>
    ''' Constructor for single address, or for not compiling RegEx's.
    ''' </summary>
    ''' <param name="address"></param>
    ''' <remarks></remarks>
    Public Sub New(Optional ByVal address As String = "")
        Call Setup_lstSUD()
        Call Setup_lstStSuf()
        Call Reset()
        If Len(address) > 0 Then Call Parse(address)
    End Sub

    ''' <summary>
    ''' Constructor for setting compilation for RegEx's.
    ''' </summary>
    ''' <param name="CompileRegEx"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal CompileRegEx As Boolean)
        If CompileRegEx Then _regOpt = (_regOpt Or RegexOptions.Compiled)
        Call Setup_lstSUD()
        Call Setup_lstStSuf()
        Call Reset()
    End Sub

#End Region

End Class
