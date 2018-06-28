Imports System
Imports MailKit.Net
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Public Class mrepl
    Private Shared ImapServer As String
    Private Shared ImapPort As Integer = 993 ' 993 by default
    Private Shared ImapLogin As String
    Private Shared ImapPassword As String
    Private Shared ImapSSL As Boolean = True ' true by default

    Private Shared SmtpServer As String
    Private Shared SmtpPort As Integer = 25  ' 25 by default
    Private Shared SmtpLogin As String
    Private Shared SmtpPassword As String
    Private Shared SmtpSSL As Boolean

    Private Shared SmtpFromName As String
    Private Shared SmtpFromAddress As String

    Private Shared SourceFilter As String
    Private Shared EnableLogging As Boolean

    Public Shared Sub Main()
        ' first initialize necessary settings
        ' from probable config file
        Dim fi() As System.Reflection.FieldInfo = GetType(mrepl).GetFields(System.Reflection.BindingFlags.Static Or System.Reflection.BindingFlags.NonPublic)
        If System.IO.File.Exists("mrepl.conf") Then
            ParseSettingsFile("mrepl.conf", fi)
        ElseIf System.IO.File.Exists("mrepl.ini") Then
            ParseSettingsFile("mrepl.ini", fi)
        End If
        ' or from command line args (have priority over config file)
        ParseCommandLine(System.Environment.GetCommandLineArgs, fi)
        Erase fi : fi = Nothing



        ' next initialize logging and log the time of run
        If EnableLogging Then
            Dim logPath As String = IO.Path.Combine(IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "log")
            If IO.Directory.Exists(logPath) = False Then
                Try
                    IO.Directory.CreateDirectory(logPath)
                Catch
                    logPath = Nothing
                End Try
            End If
            If String.IsNullOrEmpty(logPath) = False Then
                Try
                    logStream = New System.IO.FileStream(IO.Path.Combine(logPath, "mrepl.log"), _
                                                         System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write, System.IO.FileShare.Read)
                    logWriter = New System.IO.StreamWriter(logStream)
                Catch : End Try
            End If
        End If
        Console.ForegroundColor = ConsoleColor.Gray
        LogLine("-------- " & Date.Now.ToString("dd.MM.yyyy HH:mm") & " --------")

        ' next check necessary settings are initialized (log errors as needed)
        If String.IsNullOrEmpty(ImapServer) Then
            LogLine("Error: IMAP server not specified, exiting")
            Exit Sub
        End If
        If String.IsNullOrEmpty(ImapLogin) Or String.IsNullOrEmpty(ImapPassword) Then
            LogLine("Error: IMAP server credentials not specified, exiting")
            Exit Sub
        End If
        If String.IsNullOrEmpty(ImapServer) Then
            LogLine("Error: SMTP server not specified, exiting")
            Exit Sub
        End If
        If String.IsNullOrEmpty(SmtpFromName) Or String.IsNullOrEmpty(SmtpFromAddress) Then
            LogLine("Error: 'From' user name or address not specified, exiting")
            Exit Sub
        End If

        ' next initialize IMAP client
        Dim imapClient As New Imap.ImapClient()
        Try
            ' a dirty hack to pass certificate issue
            System.Net.ServicePointManager.ServerCertificateValidationCallback = Function(s As Object, _
                                                                                          cert As System.Security.Cryptography.X509Certificates.X509Certificate, _
                                                                                          chain As System.Security.Cryptography.X509Certificates.X509Chain, _
                                                                                          sslPolicyErrors As System.Net.Security.SslPolicyErrors)
                                                                                     Return True
                                                                                 End Function

            ' Connect IMAP client to server
            Log(String.Format("Connecting to {0}:{1}... ", ImapServer, ImapPort))
            imapClient.Connect(ImapServer, ImapPort, ImapSSL)
            LogLine("done.")

            ' Authenticate IMAP client
            Log(String.Format("Authenticating as [{0}]... ", ImapLogin))
            imapClient.Authenticate(ImapLogin, ImapPassword)
            LogLine("done.")

            ' Open inbox and get list of message ids filtered by "NotSeen" criterium (to speed up reading long list of messages and give the possibility to fine-tune
            ' what messages to replicate, event for a second time
            Log("Getting new messages... ")
            imapClient.Inbox.Open(MailKit.FolderAccess.ReadWrite)
            Dim uids As IList(Of MailKit.UniqueId) = imapClient.Inbox.Search(MailKit.Search.HeaderSearchQuery.NotSeen)

            ' If there are some unseen messages...
            If uids IsNot Nothing AndAlso uids.Count > 0 Then
                LogLine("total " & uids.Count.ToString & ".")

                ' Now we know that SMTP client is probably also needed, so initialize it and authenticate if there's SMTP authentication
                Dim smtpClient As New Smtp.SmtpClient()
                smtpClient.Connect(SmtpServer, SmtpPort, SmtpSSL)
                If String.IsNullOrEmpty(SmtpLogin) = False And String.IsNullOrEmpty(SmtpPassword) = False Then
                    smtpClient.Authenticate(SmtpLogin, SmtpPassword)
                End If

                ' For each unseen message... 
                For Each uid As MailKit.UniqueId In uids
                    ' Download it from server and log on screen
                    Dim imapMessage As MimeKit.MimeMessage = imapClient.Inbox.GetMessage(uid)
                    Console.ForegroundColor = ConsoleColor.White
                    LogLine(String.Format("{0} [{1}] {2}", imapMessage.Date.ToString("dd.MM.yyyy HH:mm"), _
                                                                        TrimString(imapMessage.From(0).Name, 20), _
                                                                        TrimString(imapMessage.Subject, 39)))
                    Console.ForegroundColor = ConsoleColor.Gray

                    ' Check sender filter
                    If String.IsNullOrEmpty(SourceFilter) OrElse imapMessage.From.Any(Function(f) Regex.IsMatch(f.ToString, SourceFilter)) Then

                        ' If OK, parse subject. If there's no [some_prefix:] in subject, save for "re:" and "fwd:", leave it alone
                        Dim imapSubject As String = System.Text.RegularExpressions.Regex.Replace(imapMessage.Subject, "^(?:re|fwd?)\:\s*", String.Empty, Text.RegularExpressions.RegexOptions.IgnoreCase)
                        If imapSubject.Contains(":") Then
                            ' Otherwise look if there's a mailing list named the same as [some_prefix:] in message subject
                            Dim addresses As List(Of String) = GetAddressesByPrefix(imapSubject.Substring(0, imapSubject.IndexOf(":"c)).Trim)
                            ' If so, we prepare message for forwarding: alter the date, "from:" and "to:" fields
                            If addresses.Count > 0 Then
                                imapMessage.Date = Date.Now.ToUniversalTime
                                imapMessage.Subject = imapSubject
                                imapMessage.From.Clear() : imapMessage.From.Add(New MimeKit.MailboxAddress(SmtpFromName, SmtpFromAddress))
                                ' Then forward message to each of the addresses from the mailing list
                                addresses.ForEach(Sub(addr)
                                                      imapMessage.To.Clear() : imapMessage.To.Add(New MimeKit.MailboxAddress(String.Empty, addr))
                                                      Log(String.Format("    Forwarding to {0}... ", addr))
                                                      Try
                                                          smtpClient.Send(imapMessage, _
                                                                           imapMessage.From.Mailboxes.Last, _
                                                                           imapMessage.To.Mailboxes
                                                                           )
                                                          LogLine("success.")
                                                      Catch ex As Exception
                                                          LogLine("Error: " & ex.Message)
                                                      End Try
                                                  End Sub)
                            End If
                            addresses.Clear() : addresses = Nothing
                        End If
                        imapSubject = Nothing
                    End If

                    ' Mark message as seen
                    imapClient.Inbox.AddFlags(uid, MailKit.MessageFlags.Seen, True)
                    imapMessage = Nothing
                Next

                ' Dispose SMTP client
                smtpClient.Disconnect(True)
                smtpClient.Dispose()
                uids.Clear()
            Else
                LogLine("total 0.")
            End If

            ' Close mailbox and dispose IMAP client
            imapClient.Inbox.Close()
            imapClient.Disconnect(True)
            LogLine("All tasks completed." & Environment.NewLine)

        Catch ex As Exception
            LogLine("Error: " & ex.Message)
        End Try

        imapClient.Dispose()

        ' dispose logging
        If logWriter IsNot Nothing Then
            logWriter.Close()
            logWriter.Dispose()
        End If
        If logStream IsNot Nothing Then
            logStream.Close()
            logStream.Dispose()
        End If
    End Sub

    Private Shared Function GetAddressesByPrefix(Text As String) As List(Of String)
        Dim retval As New List(Of String)
        If String.IsNullOrEmpty(Text) OrElse Text.Equals("sample", StringComparison.InvariantCultureIgnoreCase) Then Return retval
        Dim listDir As String = IO.Path.Combine(IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "lists")
        If IO.Directory.Exists(listDir) = False Then Return retval

        Try
            Dim strFilename As String = System.IO.Directory _
                .GetFiles(listDir, "*.lst") _
                .FirstOrDefault(Function(filename)
                                    Return System.IO.Path.GetFileNameWithoutExtension(filename).Equals(Text, StringComparison.InvariantCultureIgnoreCase)
                                End Function)
            If strFilename IsNot Nothing Then
                Using sr As New System.IO.StreamReader(strFilename)
                    Dim ln As String
                    While Not sr.EndOfStream
                        ln = sr.ReadLine
                        If String.IsNullOrEmpty(ln) = False AndAlso Regex.IsMatch(ln, "^\s*[\w\-\.]+\@[\w\-\.]+") Then
                            retval.Add(Regex.Match(ln, "[\w\-\.]+\@[\w\-\.]+").Value)
                        End If
                    End While
                End Using
            End If
        Catch
        End Try
        Return retval
    End Function

#Region "Settings"
    Private Shared Sub ParseSettingsFile(File As String, Fields As System.Reflection.FieldInfo())
        If System.IO.File.Exists(File) = False Or Fields Is Nothing Then Exit Sub
        Dim l As String, poe As Integer
        Try
            Using fs As New System.IO.FileStream(File, System.IO.FileMode.Open, System.IO.FileAccess.Read)
                Using sr As New System.IO.StreamReader(fs)
                    While Not sr.EndOfStream
                        l = sr.ReadLine.Trim
                        If String.IsNullOrEmpty(l) = False AndAlso (l.StartsWith("#") = False And l.StartsWith(";") = False) Then
                            poe = l.IndexOf("="c)
                            If poe <= 0 Or poe >= l.Length - 1 Then Throw New System.ArgumentException("Wrong configuration file format")
                            InitSetting(l.Substring(0, poe).Trim, l.Substring(poe + 1).Trim, Fields)
                        End If
                    End While
                End Using
            End Using
        Catch
        End Try
        l = Nothing
    End Sub
    Private Shared Sub ParseCommandLine(Args() As String, Fields As System.Reflection.FieldInfo())
        If Args Is Nothing OrElse Args.Length <= 1 OrElse Fields Is Nothing Then Exit Sub
        For it As Integer = 1 To Args.GetUpperBound(0)
            Dim poe As Integer = Args(it).IndexOf("="c)
            If Args(it).StartsWith("--") And poe > 2 And poe < Args(it).Length - 1 Then
                InitSetting(Args(it).Substring(2, poe - 2).Trim, Args(it).Substring(poe + 1), Fields)
            End If
        Next
    End Sub

    Private Shared Sub InitSetting(Key As String, Value As String, Fields As System.Reflection.FieldInfo())
        If String.IsNullOrEmpty(Key) Or String.IsNullOrEmpty(Value) Then Exit Sub
        For it As Integer = 0 To Fields.GetUpperBound(0)
            If Fields(it).Name.Equals(Key) Then
                Try
                    If Fields(it).FieldType Is GetType(Integer) Then
                        Dim i As Integer
                        Integer.TryParse(Value, i)
                        Fields(it).SetValue(Nothing, i)
                    ElseIf Fields(it).FieldType Is GetType(Boolean) Then
                        Dim b As Boolean
                        Boolean.TryParse(Value, b)
                        Fields(it).SetValue(Nothing, b)
                    Else
                        Fields(it).SetValue(Nothing, Value)
                    End If
                Catch : End Try
                Exit Sub
            End If
        Next
    End Sub
#End Region

#Region "Logging"
    Private Shared logStream As System.IO.FileStream
    Private Shared logWriter As System.IO.StreamWriter
    Private Shared Sub Log(Text As String)
        Console.Write(Text)
        If logWriter IsNot Nothing Then logWriter.Write(Text)
    End Sub
    Private Shared Sub LogLine(Text As String)
        Console.WriteLine(Text)
        If logWriter IsNot Nothing Then logWriter.WriteLine(Text)
    End Sub

    Private Shared Function TrimString(Input As String, Length As Integer) As String
        If Input Is Nothing Then Return String.Empty
        If Length < 3 Then Length = 3
        If Input.Length > Length Then
            Return Input.Substring(0, Length - 3) & "..."
        Else
            Return Input.PadRight(Length)
        End If
    End Function
#End Region
End Class

