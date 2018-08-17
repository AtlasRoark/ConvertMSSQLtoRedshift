Imports System.IO
Imports System.Text

Module Convert

    Private rxBracket = New RegularExpressions.Regex("\[([^\]]*)\]")
    Private newTable As String = String.Empty

    Sub Main()

        Dim args() As String = System.Environment.GetCommandLineArgs()

        If args.Count <> 3 Then
            Console.WriteLine()
            Console.WriteLine("Usage:")
            Console.WriteLine()
            Console.WriteLine("ConvertSQL mssql_filepath local_repo_path_of_customdatasets")
            Console.WriteLine()
            Console.WriteLine("File Name in mssql_filepath is the Folder Name to write to in customdatasets")
            Console.WriteLine()
            Console.WriteLine("Example:")
            Console.WriteLine()
            Console.WriteLine("""C:\Users\Peter Grillo\Documents\SQL Server Management Studio\PrecisionFranchiseAdministration.sql"" ""C: \Users\Peter Grillo\source\repos\customdatasets""")
            Exit Sub
        End If

        Dim inPath As String = args(1)
        If Not File.Exists(inPath) Then
            Console.WriteLine($"File not found: {inPath}")
            Exit Sub
        End If

        If Not Directory.Exists(args(2)) Then
            Console.WriteLine($"Directory not found: {args(2)}")
            Exit Sub
        End If

        Dim sql As String = My.Computer.FileSystem.ReadAllText(inPath)
        Dim outPath = args(2) + inPath.Substring(InStrRev(inPath, "\") - 1)
        outPath = outPath.Replace(".sql", "\query.sql")


        Dim sqlLine As String = String.Empty
        Dim sqlPos As Integer = 0
        Dim sqlEnd As Integer = 0
        Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(inPath)
        Dim writer As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(outPath, False)
        Dim rxQuote = New RegularExpressions.Regex("\'([^\]]*)\'")
        Dim rxFrom As New RegularExpressions.Regex("FROM \[([^\]]*)\].")
        Dim rxJoin As New RegularExpressions.Regex("JOIN \[([^\]]*)\].")
        Dim schemaName As String = String.Empty
        Dim tmpTable As String = String.Empty

        Do
            sqlLine = reader.ReadLine

            If sqlLine Is Nothing Then Exit Do
            If sqlLine.Contains("End Query") Then Exit Do
            If sqlLine.Trim.StartsWith("--") Then Continue Do

            If sqlLine.Contains("External Parameters") Then
                writer.WriteLine(sqlLine)
                sqlLine = reader.ReadLine()
                'writer.WriteLine(sqlLine.Insert(0, "--"))
                sqlLine = reader.ReadLine()
                'writer.WriteLine(sqlLine.Insert(0, "--"))

                writer.WriteLine("    DROP TABLE IF EXISTS #tmp_variables;")
                writer.WriteLine("    CREATE TEMP TABLE #tmp_variables AS SELECT")
                writer.WriteLine("        (SELECT @From_Utc)::DATETIME AS utc_start,")
                writer.WriteLine("        (SELECT @To_Utc)::DATETIME AS utc_end;")
                writer.WriteLine()

                sqlLine = reader.ReadLine
            End If

            sqlPos = InStr(sqlLine, "'tempdb..")
            If sqlPos > 0 Then
                If rxQuote.IsMatch(sqlLine) Then
                    tmpTable = rxQuote.Match(sqlLine).Value.Replace("tempdb..", "")
                    tmpTable = tmpTable.Replace("'", "")
                    'sqlEnd = InStr(sqlPos + 1, sqlLine, "'")
                    'Dim tmpTable As String = sqlLine.Substring(sqlPos + 8, sqlEnd - sqlPos - 10)
                    sqlLine = $"    DROP TABLE IF EXISTS {tmpTable};"
                End If
            End If

            sqlPos = InStr(sqlLine, "FROM")
            If sqlPos > 0 Then
                If rxFrom.IsMatch(sqlLine) Then
                    schemaName = rxBracket.Match(sqlLine).Value
                    sqlLine = ReplaceView(sqlLine, schemaName)
                End If
            End If

            sqlPos = InStr(sqlLine, "JOIN")
            If sqlPos > 0 Then
                If rxJoin.IsMatch(sqlLine) Then
                    schemaName = rxBracket.Match(sqlLine).Value
                    sqlLine = ReplaceView(sqlLine, schemaName)
                    sqlLine = sqlLine.Replace($"#v_.{tmpTable}", $"#v_{newTable.ToLower}")
                End If
            End If

            sqlLine = sqlLine.Replace("[", """")
            sqlLine = sqlLine.Replace("]", """")

            writer.WriteLine(sqlLine)
        Loop

        reader.Close()
        writer.Close()

    End Sub

    Private Function ReplaceView(sqlLine As String, schema As String) As String
        Dim rxTable = New RegularExpressions.Regex("\.(\S*)\ ")

        sqlLine = sqlLine.Replace(schema, "#v_")
        If rxTable.IsMatch(sqlLine) Then
            Dim tableName As String = rxTable.Match(sqlLine).Value.Trim
            newTable = tableName.Replace(".", "")
            If rxBracket.IsMatch(tableName) Then
                newTable = newTable.Replace("[", "")
                newTable = newTable.Replace("]", "")
            End If
            sqlLine = sqlLine.Replace(tableName, newTable.ToLower)
        End If
        Return sqlLine
    End Function
End Module
