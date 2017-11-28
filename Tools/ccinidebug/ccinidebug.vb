Imports System.IO
Module ccinidebug
    Sub Main()

        Dim lines As New ArrayList()
        Dim trimmedCCIniEntryList As New List(Of String)
        Dim trimmedWA2IniEntryList As New List(Of String)
        Dim entriesToPrune As New List(Of String)
        If File.Exists(Environment.CurrentDirectory & "\ccleaner.ini") = False Then
            Console.WriteLine("ccleaner.ini file could not be located in the current working directory (" & Environment.CurrentDirectory & ")")
            Console.ReadKey()
            End
        End If

        Dim r As IO.StreamReader
        Try

            r = New IO.StreamReader(Environment.CurrentDirectory & "\ccleaner.ini")
            Do While (r.Peek() > -1)
                Dim currentLine As String = r.ReadLine.ToString
                lines.Add(currentLine)
                If currentLine.StartsWith("(App)") Then
                    Dim tmp1 As String() = Split(currentLine, "(App)")
                    Dim tmp2 As String() = Split(tmp1(1), "=")
                    If tmp2(0).Contains("*") Then
                        trimmedCCIniEntryList.Add(tmp2(0))
                    End If
                End If

            Loop
            r.Close()
            Console.WriteLine("Done adding the lines to a listy thingy")
            lines.Sort()
            lines.Remove("[Options]")
            lines.Insert(0, "[Options]")
            Console.WriteLine("Press Y to prune stale winapp2.ini entries from ccleaner.ini")
            Try
                If Console.ReadLine().ToLower = "y" Then
                    Dim w As IO.StreamReader
                    w = New IO.StreamReader(Environment.CurrentDirectory & "\winapp2.ini")
                    Do While (w.Peek > -1)
                        Dim curWA2Line As String = w.ReadLine.ToString()

                        If curWA2Line.StartsWith("[") Then

                            curWA2Line = curWA2Line.Remove(0, 1)
                            curWA2Line = curWA2Line.Remove(curWA2Line.Length - 1)
                            trimmedWA2IniEntryList.Add(curWA2Line)

                        End If

                    Loop
                    w.Close
                    For Each entry As String In trimmedCCIniEntryList

                        If Not trimmedWA2IniEntryList.Contains(entry) Then

                            entriesToPrune.Add(entry)

                        End If

                    Next

                    Dim linesCopy As New ArrayList()
                    linesCopy.AddRange(lines)
                    Console.WriteLine("The following stale entries will be pruned: ")
                    For Each item As String In entriesToPrune
                        Console.WriteLine(item)
                        For Each entry As String In lines

                            If entry.Contains(item) Then

                                linesCopy.Remove(entry)

                            End If

                        Next

                    Next
                    Console.WriteLine("****************************************************************************************************")
                    lines = linesCopy

                End If
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                Console.ReadKey()
            End Try
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.ReadKey()
        End Try
        Console.WriteLine("Press Y to replace the alphabetical ordering in ccleaner.ini")

        If Console.ReadLine.ToString.ToLower = "y" Then
            Console.WriteLine("Modifying ccleaner.ini....")
            Try
                Dim file As New System.IO.StreamWriter(Environment.CurrentDirectory & "\ccleaner.ini", False)

                For Each line As String In lines
                    file.WriteLine(line.ToString)
                Next
                file.WriteLine(Environment.NewLine)
                file.Close()
                Console.WriteLine("****************************************************************************************************")
                Console.WriteLine("Finished modifying ccleaner.ini, press any key to exit.")
                Console.ReadKey
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                Console.ReadKey()
            End Try
        End If
    End Sub

End Module