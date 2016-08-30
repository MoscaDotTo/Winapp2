Module Module1
    Dim apath As String = Environment.CurrentDirectory
    Sub Main()
        Dim winapp2 As String = apath & "\winapp2.ini"
        Dim winapp As String = apath & "\winapp.ini"
        Dim winreg As String = apath & "\winreg.ini"
        Dim winsys As String = apath & "\winsys.ini"

        'Export the winapp files.
        If IO.File.Exists(Environment.CurrentDirectory & "\ccleaner.exe") Then
            Process.Start(Environment.CurrentDirectory & "\ccleaner.exe", "/export")
        Else
            Console.WriteLine("Could not find CCleaner.exe in current directory.")
            Console.ReadKey()
        End If



        'Declare the textreader
        Dim r As IO.StreamReader
        Dim rule As String

        'Declare the textwriter
        Dim SW As IO.TextWriter
        Dim out_text As String = apath & "\regkeys.output.txt"
        SW = IO.File.AppendText(out_text)

        'Read winapp2.ini   
        If My.Computer.FileSystem.FileExists(winapp2) = True Then
            Try

                r = New IO.StreamReader(winapp2)
                Do While (r.Peek() > -1)
                    rule = (r.ReadLine.ToString)

                    'Perform some analysis
                    If rule.StartsWith("RegKey") = True Then
                        rule = rule.Remove(0, 8)
                        rule = rule.Replace("|", "\")
                        rule = rule.Replace("=", "")
                        rule = rule.Replace("HKU", "HKUS")
                        SW.WriteLine(rule)
                    End If
                Loop
            Catch ex As Exception
            End Try
        End If


        'Read winapp.ini
        Try

            r = New IO.StreamReader(winapp)
            Do While (r.Peek() > -1)
                rule = (r.ReadLine.ToString)

                'Perform some analysis
                If rule.StartsWith("RegKey") = True Then
                    rule = rule.Remove(0, 8)
                    rule = rule.Replace("|", "\")
                    rule = rule.Replace("=", "")
                    rule = rule.Replace("HKU", "HKUS")
                    SW.WriteLine(rule)
                End If
            Loop
        Catch ex As Exception
        End Try

        'Read winreg.ini
        Try

            r = New IO.StreamReader(winreg)
            Do While (r.Peek() > -1)
                rule = (r.ReadLine.ToString)

                'Perform some analysis
                If rule.StartsWith("RegKey") = True Then
                    rule = rule.Remove(0, 8)
                    rule = rule.Replace("|", "\")
                    rule = rule.Replace("=", "")
                    rule = rule.Replace("HKU", "HKUS")
                    SW.WriteLine(rule)
                End If
            Loop
        Catch ex As Exception
        End Try

        'Read winsys.ini
        Try

            r = New IO.StreamReader(winsys)
            Do While (r.Peek() > -1)
                rule = (r.ReadLine.ToString)

                'Perform some analysis
                If rule.StartsWith("RegKey") = True Then
                    rule = rule.Remove(0, 8)
                    rule = rule.Replace("|", "\")
                    rule = rule.Replace("=", "")
                    rule = rule.Replace("HKU", "HKUS")
                    SW.WriteLine(rule)
                End If
            Loop
        Catch ex As Exception
        End Try


    End Sub

End Module

