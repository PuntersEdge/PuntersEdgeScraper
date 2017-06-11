Imports HtmlAgilityPack
Imports System
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Data.OleDb
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Web.Script.Serialization
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports PuntersEdgeScraperApp.DatabseActions
Imports System.Data.SqlClient

Public Module GlobalVariables
    Public SessionToken As String
    Public Appkey As String = "ZfI8hcEMs3uAzPmD"
    Public Threadcount As Int16 = 0
End Module '

Module Module1

    Sub Main()
        '-------------------------------------------------------------------------------------------------------------------
        'Pull out MarketId's to update
        '-------------------------------------------------------------------------------------------------------------------


        Dim database As New DatabseActions
        Dim Time As String = Format(Date.Now, "HH:mm")
        Dim marketIds As DataTable = database.SELECTSTATEMENT("MarketID", "BetfairMarketIds", "WHERE RaceTime > '" & Time & "'")
        Dim API As New BetfairAPI


        SessionToken = API.GetSessionKey("ZfI8hcEMs3uAzPmD", "username=00alawre&password=portsmouth1")

        database.SQL("TRUNCATE TABLE ProcessLog")
        '-------------------------------------------------------------------------------------------------------------------
        '
        'Iterate through row and update odds
        '-------------------------------------------------------------------------------------------------------------------

        For Each marketID_row As DataRow In marketIds.Rows()

            'Call BetfairOddsUpdater(marketID_row)

            Dim Thread As New Threading.Thread(AddressOf BetfairOddsUpdater)
            Thread.IsBackground = False

            Thread.Start(marketID_row)
        Next

        '-----------------------------------------------------------------------------------------------------------------------------

        Dim con As New SqlClient.SqlConnection(Configuration.ConfigurationManager.ConnectionStrings("PuntersEdgeDB").ConnectionString) 'connection string

        Dim command As SqlCommand = New SqlCommand
        Dim dt As New DataSet
        Dim cmd As String = "SELECT DISTINCT Meeting FROM TodaysRaces"

        Dim adapter As New SqlDataAdapter(cmd, con)
        Dim row As DataRow


        'get todays races to build URL for odds
        con.Open()
        adapter.Fill(dt, "races")
        con.Close()


        For Each row In dt.Tables("races").Rows()

            'Grab race parameters from today's races dataset

            Dim Meet As String = row.Item("Meeting").ToString

            'Multi Thread
            Dim t1 As New Threading.Thread(AddressOf Updater)
            t1.IsBackground = False

            t1.Start(Meet)

        Next
    End Sub

    Private Sub BetfairOddsUpdater(ByVal row As DataRow)

        Dim database As New DatabseActions
        Dim API As New BetfairAPI


        'Declare marketIDs and get Json for horse names and prices
        Dim marketID As String = row.Item(0).ToString
        Dim priceJson As String = API.GetPrices(Appkey, SessionToken, marketID)

        Dim o As JObject = JObject.Parse(priceJson)
        Dim results As List(Of JToken) = o.Children().ToList

        For Each item As JProperty In results
            item.CreateReader()

            If item.Value.Type = JTokenType.Array Then
                For Each selectionID As JObject In item.Values



                    Dim Matched As String = selectionID("totalMatched")


                    Dim runners As List(Of JToken) = selectionID.Children().ToList

                    For Each runner As JProperty In runners
                        runner.CreateReader()

                        If runner.Value.Type = JTokenType.Array Then

                            For Each horse As JObject In runner.Values

                                Dim values As String = ""

                                Dim Status As String = horse("status")

                                If Not Status = "REMOVED" Then

                                    Dim lastprice As Decimal = 0
                                    Dim selection As String = horse("selectionId")

                                    If Not IsNothing(horse("lastPriceTraded")) Then
                                        lastprice = horse("lastPriceTraded")
                                    End If


                                    Dim selection_TotalMatched As String = horse("totalMatched")


                                    Using con As New SqlConnection(Configuration.ConfigurationManager.ConnectionStrings("PuntersEdgeDB").ConnectionString)

                                        Dim comm As New SqlCommand
                                        comm.CommandText = "UPDATE BetFairData SET LastTradedprice=" & lastprice & ",
                                                            Market_TotalMatched = " & Matched & ",
                                                            selection_TotalMatched = " & selection_TotalMatched & "
                                                            WHERE SelectionID = " & selection

                                        comm.Connection = con

                                        con.Open()
                                        comm.ExecuteNonQuery()
                                        con.Close()

                                    End Using


                                Else


                                End If


                            Next

                        End If

                    Next

                Next

            End If

        Next

        database.EXECSPROC("RunLiveSelections", "")



    End Sub 'Updates the betfairData table with odds and money data (called by first scrape and update)                                       ***** COMPLETE *****


    Private Sub Updater(ByVal Meeting As String)


        'Pull out the races for the meeting passed in -----------------------------------------------------------------------------------------------------
        Dim con As New SqlClient.SqlConnection(Configuration.ConfigurationManager.ConnectionStrings("PuntersEdgeDB").ConnectionString) 'connection string
        Dim command As SqlCommand = New SqlCommand
        Dim dt As New DataSet
        Dim cmd As String = "SELECT Meeting, CONVERT(varchar(5), RaceTime, 108) As RaceTime FROM TodaysRaces WHERE Meeting='" & Meeting & "'"
        Dim adapter As New SqlDataAdapter(cmd, con)
        'Dim row As DataRow

        con.Open()
        adapter.Fill(dt, "races")
        con.Close()

        '--------------------------------------------------------------------------------------------------------------------------------------------------
        Dim TotalCount As Int16 = dt.Tables("races").Rows.Count
        Dim Count1 As Int16 = TotalCount / 2
        Dim Count2 As Int16 = TotalCount - Count1

        Dim Half1 As String = "TOP " & Count1.ToString & ",ASC"
        Dim Half2 As String = "TOP " & Count2.ToString & ",DESC"

        'First Thread
        Dim t1 As New Threading.Thread(AddressOf MT_Update)
        t1.IsBackground = False

        t1.Start(Meeting & "," & Half1)

        Thread.Sleep(1000)

        'Second Thread
        Dim t2 As New Threading.Thread(AddressOf MT_Update)
        t2.IsBackground = False

        t2.Start(Meeting & "," & Half2)



    End Sub ' Multithread scraping to update odds/results throughout the day. Calls MT_Update                                                         ***** COMPLETE *****
    Private Sub MT_Update(ByVal Parameters As String)



        'Create a variable for start time:
        Dim TimerStart As DateTime
        TimerStart = Now
        Dim TimeFormatted As String = TimerStart.ToString("HH:mm:ss")



        Dim parts As String() = Parameters.Split(New Char() {","c})
        Dim Meet As String = parts(0).ToString
        Dim Count As String = parts(1).ToString
        Dim Sort As String = parts(2).ToString

        'Pull out the races for the meeting passed in -----------------------------------------------------------------------------------------------------
        Dim con As New SqlClient.SqlConnection(Configuration.ConfigurationManager.ConnectionStrings("PuntersEdgeDB").ConnectionString) 'connection string
        Dim command As SqlCommand = New SqlCommand
        Dim dt As New DataSet
        Dim RowsAffected As Int16 = 0



        command.CommandText = "INSERT INTO ProcessLog(Meeting, ProcessTime, Log, Processed) VALUES ('" & Meet & "','" & TimeFormatted & "', 'Processing of " & Meet & " started at " & TimeFormatted & "', 0)"
        command.Connection = con
        con.Open()
        command.ExecuteNonQuery()
        con.Close()

        Dim cmd As String = "SELECT " & Count & " Meeting, CONVERT(varchar(5), RaceTime, 108) As RaceTime FROM TodaysRaces WHERE Meeting='" & Meet & "' ORDER BY RaceTime " & Sort
        Dim adapter As New SqlDataAdapter(cmd, con)
        Dim row As DataRow

        con.Open()
        adapter.Fill(dt, "races")
        con.Close()

RetryThread:

        Try


            ' --------------------------------------------------------------------------------------------------------------------------------------------------
            'Go through each race in results
            For Each row In dt.Tables("races").Rows()

                'Grab race parameters from today's races dataset
                Dim Time As String = row.Item("RaceTime").ToString
                'Dim Meet As String = row.Item("Meeting").ToString

                Dim db As New DatabseActions


                'Create URL for race
                Dim url As String = "http://www.oddschecker.com/horse-racing" & "/" & Meet & "/" & Time & "/winner"



                'Load entire HTML page into variable
                Dim htmlDoc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlWeb().Load(url)
                Dim html As String = htmlDoc.DocumentNode.OuterHtml

                'Pull out neccesary tags from HTML variable
                Dim tabletag = htmlDoc.DocumentNode.SelectNodes("//tr[@class='diff-row eventTableRow bc']")
                Dim nonrunners = htmlDoc.DocumentNode.SelectNodes("//tr[@class='diff-row eventTableRowNonRunner']")
                Dim bookie = htmlDoc.DocumentNode.SelectNodes("//td[@data-bk]")

                'Declare arrays for indexing and building SQL commands
                Dim bookiearray As ArrayList = New ArrayList
                Dim Details As ArrayList = New ArrayList



                'Load available bookies into array for indexing
                For Each bookienode As HtmlNode In bookie

                    If bookienode.Attributes.Contains("data-bk") Then

                        If Not bookiearray.Contains(bookienode.Attributes("data-bk").Value.ToString) Then

                            bookiearray.Add(bookienode.Attributes("data-bk").Value.ToString)

                        End If


                    End If

                Next

                If Not IsNothing(nonrunners) Then

                    For Each nonrunner As HtmlNode In nonrunners




                        Dim cunt As String = nonrunner.Attributes("data-bname").Value.ToString


                        Dim table As String = Meet.Replace("-", " ")
                        command.CommandType = CommandType.Text
                        command.Connection = con
                        command.CommandText = "UPDATE [" & table & "] SET Odds= 0 WHERE Horse ='" & cunt & "'"

                        con.Open()
                        command.ExecuteNonQuery()
                        con.Close()

                        command.CommandText = "UPDATE [Results] SET Result = 'NR' WHERE Horse='" & cunt & "' AND Time = '" & Time & "'"

                        con.Open()
                        command.ExecuteNonQuery()
                        con.Close()



                    Next
                End If


                If Convert.ToDateTime(Time) < DateAdd(DateInterval.Minute, -3, DateTime.Now) Then

                    'db.UPDATE("Results", "Result", "-1", "WHERE Time='" & Time & "'")
                    command.CommandText = "UPDATE Results SET Result ='-1' WHERE Time = '" & Time & "' AND Result <> 'NR'"
                    command.Connection = con

                    con.Open()
                    command.ExecuteNonQuery()
                    con.Close()



                End If



                'recurse through each horse and pull out horse name and odds
                For Each rowtag As HtmlNode In tabletag

                    'Dim update command to update all odds for horse
                    Dim Update As String = "" 'UPDATE instead of truncate

                    Dim horseindex As Integer = 0
                    'Check to see if the race has results ----------------------------------------------------------------------------------------------------
                    If rowtag.ChildNodes(0).ChildNodes(0).HasAttributes = True Then
                        Dim result As String = rowtag.ChildNodes(0).ChildNodes(0).InnerText.ToString
                        Dim horse As String = rowtag.Attributes("data-bname").Value.ToString

                        If result Is Nothing Then result = "Unplaced"

                        command.CommandType = CommandType.Text
                        command.Connection = con
                        command.CommandText = "UPDATE [Results] SET Result ='" & result & "' WHERE Horse='" & horse & "' AND Time = '" & Time & "'"

                        con.Open()
                        command.ExecuteNonQuery()
                        con.Close()


                        '------------------------------------------------------------------------------------------------------------------------------------

                    Else

                        'Check the time, if the race is less than 20 minutes away, bail out
                        If Convert.ToDateTime(DateTime.Now.ToShortTimeString) > DateAdd(DateInterval.Minute, -20, Convert.ToDateTime(Time)) Then

                            GoTo Escape

                        Else
                            '---------------------------------------------------------------------------------------------------------------------------------
                            'If there are no results and the time is not within 20 minutes, update the horse odds
                            If rowtag.HasChildNodes = True Then

                                Dim node As HtmlAgilityPack.HtmlNode
                                Dim RaceHorse As String = rowtag.ChildNodes(1).ChildNodes("a").Attributes("data-name").Value.ToString



                                'Pull out all odds for that horse by recursing through tags within the parent horse tag
                                For Each node In rowtag.ChildNodes
                                    Dim table As String = Meet.Replace("-", " ")
                                    If IsNothing(node.Attributes("Class")) Then

                                        GoTo Escape

                                    Else

                                        If node.Attributes("class").Value.Contains("bc bs") Then

                                            Dim odd As String = node.Attributes("data-odig").Value.ToString


                                            'Build the SQL command for updating odds

                                            Update = Update + "UPDATE [" & table & "] SET Odds = " & odd & " WHERE Horse='" & RaceHorse & "' AND Meeting='" & Meet & "' AND BookMaker='" & bookiearray.Item(horseindex).ToString & "'"

                                            RowsAffected = RowsAffected + 1
                                            horseindex = horseindex + 1


                                        ElseIf node.Attributes("class").Value = "np o" Then



                                            Update = Update + "UPDATE [" & table & "] SET Odds = -1 WHERE Horse='" & RaceHorse & "' AND Meeting='" & Meet & "' AND BookMaker='" & bookiearray.Item(horseindex).ToString & "'"


                                            horseindex = horseindex + 1




                                        End If
                                    End If
                                Next

                            End If

                            'once the IF STATEMENT finished, execute the update command

                            command.CommandType = CommandType.Text
                            command.CommandText = Update.TrimEnd(",")
                            command.Connection = con


                            con.Open()
                            command.ExecuteNonQuery()
                            con.Close()


                        End If

                    End If
Escape:
                Next

            Next 'Move to the next horse


            Dim TimeSpent As System.TimeSpan
            TimeSpent = Now.Subtract(TimerStart)
            'MsgBox(TimeSpent.TotalSeconds & " seconds spent on " & Meet)

            command.CommandText = "UPDATE ProcessLog SET Log = 'Updated " & RowsAffected.ToString & " horses at " & Meet.Replace("-", " ") & " in " & TimeSpent.TotalSeconds.ToString & " seconds.', Processed=1 WHERE ProcessTime = '" & TimeFormatted & "' AND Meeting='" & Meet & "'"
            command.Connection = con

            con.Open()
            command.ExecuteNonQuery()
            con.Close()

        Catch ex As Exception

            con.Close()

            command.CommandText = "UPDATE ProcessLog SET Log = '" & ex.ToString & "' WHERE ProcessTime = '" & TimeFormatted & "' AND Meeting='" & Meet & "'"
            command.Connection = con

            con.Open()
            command.ExecuteNonQuery()
            con.Close()

            GoTo RetryThread

        End Try


    End Sub 'Thread object for scraper                                                                                                           ***** COMPLETE *****

End Module
