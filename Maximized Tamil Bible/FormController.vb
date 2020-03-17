Imports System.Text.RegularExpressions

Public Class FormController
    Dim currentBook As Integer
    Dim currentChapter As Integer
    Dim currentVerse As Integer
    Dim CHAPTER_COUNT_IN_BOOK() As Integer = {-100, 50, 40, 27, 36, 34, 24, 21, 4, 31, 24, 22, 25, 29, 36, 10, 13, 10, 42, 150, 31, 12, 8, 66, 52, 5, 48, 12, 14, 3, 9, 1, 4, 7, 3, 3, 3, 2, 14, 4, 28, 16, 24, 21, 28, 16, 16, 13, 6, 6, 4, 4, 5, 3, 6, 4, 3, 1, 13, 5, 5, 3, 5, 1, 1, 1, 22}
    Dim WORKSPACE_DIRECTORY_LOCATION As String = My.Settings.filedir_config
    Dim CONFIG_KEY As String = "config_default"

    Private Sub ListBoxOt1_Click(sender As Object, e As EventArgs) Handles ListBoxOt1.Click
        Try
            'Validation - Needed because clicking in blank spaces (usually at the bottom) in the listbox results in random data
            If (ListBoxOt1.SelectedIndex < 0) Then
                ListBoxOt1.SelectedIndex = ListBoxOt1.Items.Count - 1
            End If

            Dim defaultIndex As String = 1 ' Book Number for Genesis = 1
            Dim selectedBook As Integer = ListBoxOt1.SelectedIndex + defaultIndex

            NumericUpDownBook.Value = selectedBook

            ' Removing selections in other list boxes - to let user clearly know which was last selected
            ListBoxOt2.SetSelected(0, False)
            ListBoxNt1.SetSelected(0, False)
            ListBoxNt2.SetSelected(0, False)

            LabelChapters.Text = CHAPTER_COUNT_IN_BOOK(selectedBook)
            NumericUpDownChapter.Maximum = CHAPTER_COUNT_IN_BOOK(selectedBook)

            NumericUpDownChapter.Focus()

        Catch ex As Exception
            LabelStatus.Text = "Error in selecting book from Old Testament - First Half"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub ListBoxOt2_Click(sender As Object, e As EventArgs) Handles ListBoxOt2.Click
        Try
            'Validation - Needed because clicking in blank spaces (usually at the bottom) in the listbox results in random data
            If (ListBoxOt2.SelectedIndex < 0) Then
                ListBoxOt2.SelectedIndex = ListBoxOt2.Items.Count - 1
            End If

            Dim defaultIndex As String = 20 ' Book Number for Proverbs = 20
            Dim selectedBook As Integer = ListBoxOt2.SelectedIndex + defaultIndex

            NumericUpDownBook.Value = selectedBook

            ' Removing selections in other list boxes - to let user clearly know which was last selected
            ListBoxOt1.SetSelected(0, False)
            ListBoxNt1.SetSelected(0, False)
            ListBoxNt2.SetSelected(0, False)

            LabelChapters.Text = CHAPTER_COUNT_IN_BOOK(selectedBook)
            NumericUpDownChapter.Maximum = CHAPTER_COUNT_IN_BOOK(ListBoxOt2.SelectedIndex + defaultIndex)

            NumericUpDownChapter.Focus()

        Catch ex As Exception
            LabelStatus.Text = "Error in selecting book from Old Testament - Second Half"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub ListBoxNt1_Click(sender As Object, e As EventArgs) Handles ListBoxNt1.Click
        Try
            'Validation - Needed because clicking in blank spaces (usually at the bottom) in the listbox results in random data
            If (ListBoxNt1.SelectedIndex < 0) Then
                ListBoxNt1.SelectedIndex = ListBoxNt1.Items.Count - 1
            End If

            Dim defaultIndex As String = 40 ' Book Number for Matthew = 40
            Dim selectedBook As Integer = ListBoxNt1.SelectedIndex + defaultIndex

            NumericUpDownBook.Value = selectedBook

            ' Removing selections in other list boxes - to let user clearly know which was last selected
            ListBoxOt1.SetSelected(0, False)
            ListBoxOt2.SetSelected(0, False)
            ListBoxNt2.SetSelected(0, False)

            LabelChapters.Text = CHAPTER_COUNT_IN_BOOK(selectedBook)
            NumericUpDownChapter.Maximum = CHAPTER_COUNT_IN_BOOK(selectedBook)

            NumericUpDownChapter.Focus()

        Catch ex As Exception
            LabelStatus.Text = "Error in selecting book from New Testament - First Half"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub ListBoxNt2_Click(sender As Object, e As EventArgs) Handles ListBoxNt2.Click
        Try
            'Validation - Needed because clicking in blank spaces (usually at the bottom) in the listbox results in random data
            If (ListBoxNt2.SelectedIndex < 0) Then
                ListBoxNt2.SelectedIndex = ListBoxNt2.Items.Count - 1
            End If

            Dim defaultIndex As String = 54 ' Book Number for 1 Timothy = 54
            Dim selectedBook As Integer = ListBoxNt2.SelectedIndex + defaultIndex

            NumericUpDownBook.Value = selectedBook

            ' Removing selections in other list boxes - to let user clearly know which was last selected
            ListBoxOt1.SetSelected(0, False)
            ListBoxOt2.SetSelected(0, False)
            ListBoxNt1.SetSelected(0, False)

            LabelChapters.Text = CHAPTER_COUNT_IN_BOOK(selectedBook)
            NumericUpDownChapter.Maximum = CHAPTER_COUNT_IN_BOOK(selectedBook)

            NumericUpDownChapter.Focus()

        Catch ex As Exception
            LabelStatus.Text = "Error in selecting book from New Testament - Second Half"
            errorLogger(ex)
        End Try
    End Sub


    Private Sub ButtonGo_Click(sender As Object, e As EventArgs) Handles ButtonGo.Click
        Try
            ' To show the displaying form even if the user has closed it by mistake
            FormDisplay.Show()

            go(NumericUpDownBook.Text, NumericUpDownChapter.Text, NumericUpDownVerse.Text)
        Catch ex As Exception
            LabelStatus.Text = "Error in Navigating to the selected verse"
            errorLogger(ex)
        End Try
    End Sub

    Function GetRtfFile(Optional Book = 1, Optional Chapter = 1)
        Return My.Resources.ResourceManager.GetObject("_" + Book + "_" + Chapter)
    End Function

    Function ValidateAndGetWorkspaceLocation()
        Dim workspaceLocationFromConfig As String = WORKSPACE_DIRECTORY_LOCATION
        Dim isDirectoryValid As Boolean = My.Computer.FileSystem.DirectoryExists(workspaceLocationFromConfig)

        If isDirectoryValid Then
            Return workspaceLocationFromConfig
        Else
            'TODO Need to check in exe mode
            Return System.AppDomain.CurrentDomain.BaseDirectory
        End If
    End Function

    Sub go(Book, Optional Chapter = 1, Optional Verse = 1)
        Try
            Dim selectionStart As Integer

            currentBook = NumericUpDownBook.Value
            currentChapter = NumericUpDownChapter.Value
            currentVerse = NumericUpDownVerse.Value

            ' Display the location of the verse we are currently in 
            Label22.Text = "Currently In : "
            LabelFullName.Text = getBookNameFromBookNumber() + " " + NumericUpDownChapter.Value.ToString + ":" + NumericUpDownVerse.Value.ToString

            'Use LoadFile() to display verses from external location. We display only from resources now
            'FormDisplay.RichTextBox1.LoadFile("C:/Program Files/Maximized Tamil Bible/Bible/" + Book + "/" + Chapter + ".rtf")
            FormDisplay.RichTextBox1.Rtf = GetRtfFile(Book, Chapter)
            RichTextBoxMiniDisplay.Rtf = FormDisplay.RichTextBox1.Rtf

            ' Adds displayed verse to the top of the 'Recents' listbox
            If (ListBoxRecent.Items.Count > 14) Then
                ListBoxRecent.Items.RemoveAt(14)
            End If
            ListBoxRecent.Items.Insert(0, LabelFullName.Text)

            ' Adds displayed verse to session if freeze is disabled
            If (CheckBoxFreezeSession.Checked = False) Then
                ListBoxSession.Items.Insert(0, LabelFullName.Text)
            End If

            ' Set selected font
            FormDisplay.RichTextBox1.SelectAll()
            FormDisplay.RichTextBox1.SelectionFont = New Font(FormDisplay.RichTextBox1.SelectionFont.FontFamily, Int(NumericUpDownZoom.Value))
            RichTextBoxMiniDisplay.SelectAll()
            RichTextBoxMiniDisplay.SelectionFont = New Font(FormDisplay.RichTextBox1.SelectionFont.FontFamily, Int(12))

            ' Highlight selected verse in display
            Dim searchString As String
            searchString = Chapter + ":" + Verse + " "
            selectionStart = FormDisplay.RichTextBox1.Find(searchString)
            FormDisplay.RichTextBox1.Select(selectionStart, searchString.Length)
            selectionStart = RichTextBoxMiniDisplay.Find(searchString)
            RichTextBoxMiniDisplay.Select(selectionStart, searchString.Length)

            RichTextBoxMiniDisplay.Focus()
            FormDisplay.RichTextBox1.Focus()

        Catch ex As Exception
            LabelStatus.Text = "Error in go method while navigating to the selected verse"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub Controller_Click(sender As Object, e As EventArgs) Handles Me.Click
        Try
            ' Clicking the form focuses to the NumericUpDown for book
            NumericUpDownBook.Select(0, NumericUpDownBook.Value.ToString.Length)
            NumericUpDownBook.Focus()
        Catch ex As Exception
            LabelStatus.Text = "Error in Focusing on Book Input Field"
            errorLogger(ex)
        End Try

    End Sub

    Private Sub Controller_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' Show form that displays verses
            FormDisplay.Show()

            ' Load available sessions 
            loadListOfSessions()
            loadLayoutConfig()


            ' A jiggle to force 'Currently In' label to show value. Unless this is done, the label won't work if the user opens Genesis (Because the NumericUpDownBook's value never changed from its default value of 1)
            NumericUpDownBook.Value = 2
            NumericUpDownBook.Value = 1

            NumericUpDownBook.Select(0, NumericUpDownBook.Value.ToString.Length)
            NumericUpDownBook.Focus()
        Catch ex As Exception
            LabelStatus.Text = "Error while loading the form"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub NumericUpDownChapter_GotFocus(sender As Object, e As EventArgs) Handles NumericUpDownChapter.GotFocus
        Try
            ' Select everything on focus
            NumericUpDownChapter.Select(0, NumericUpDownChapter.Value.ToString.Length)
        Catch ex As Exception
            LabelStatus.Text = "Error while focusing on Chapter Input Field"
            errorLogger(ex)
        End Try
    End Sub


    Private Sub NumericUpDownVerse_GotFocus(sender As Object, e As EventArgs) Handles NumericUpDownVerse.GotFocus
        Try
            ' Select everything on focus
            NumericUpDownVerse.Select(0, NumericUpDownVerse.Value.ToString.Length)
        Catch ex As Exception
            LabelStatus.Text = "Error while focusing on Verse Input Field"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub NumericUpDownBook_GotFocus(sender As Object, e As EventArgs) Handles NumericUpDownBook.GotFocus
        Try
            ' Select everything on focus
            NumericUpDownBook.Select(0, NumericUpDownBook.Value.ToString.Length)
        Catch ex As Exception
            LabelStatus.Text = "Error while focusing on Book Input Field"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub NumericUpDownBook_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownBook.ValueChanged
        Try
            ' Setting value for 'Going to/Currently in' label
            Label22.Text = "Going to : "
            LabelFullName.Text = getBookNameFromBookNumber() + " " + NumericUpDownChapter.Value.ToString + ":" + NumericUpDownVerse.Value.ToString

            ' Find chapter count, use it for display, limiting NumericUpDown for chapters
            Dim bookIndex As Integer = NumericUpDownBook.Value
            LabelChapters.Text = CHAPTER_COUNT_IN_BOOK(bookIndex)

            ' Set rest of the NumericUpDowns to 1, set chapter UpDown's max value
            NumericUpDownChapter.Maximum = CHAPTER_COUNT_IN_BOOK(bookIndex)
            NumericUpDownChapter.Value = 1
            NumericUpDownVerse.Value = 1

            renderNumberOfVersesInCurrentChapter()

            ' Set focus on the NumericUpDown for chapters
            NumericUpDownChapter.Focus()
        Catch ex As Exception
            LabelStatus.Text = "Error in using the entered Book value"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub NumericUpDownChapter_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownChapter.ValueChanged

        Try
            ' Setting value for 'Going to/Currently in' label
            Label22.Text = "Going to : "
            LabelFullName.Text = getBookNameFromBookNumber() + " " + NumericUpDownChapter.Value.ToString + ":" + NumericUpDownVerse.Value.ToString

            renderNumberOfVersesInCurrentChapter()

            ' Set focus on the NumericUpDown for verses
            NumericUpDownVerse.Focus()
        Catch ex As Exception
            LabelStatus.Text = "Error in using the entered Chapter value"
            errorLogger(ex)
        End Try
    End Sub

    Function getBookNameFromBookNumber()
        Try
            If (NumericUpDownBook.Value < 20) Then
                Return ListBoxOt1.Items.Item(NumericUpDownBook.Value - 1)
            ElseIf (NumericUpDownBook.Value < 40) Then
                Return ListBoxOt2.Items.Item(NumericUpDownBook.Value - 20)
            ElseIf (NumericUpDownBook.Value < 54) Then
                Return ListBoxNt1.Items.Item(NumericUpDownBook.Value - 40)
            Else
                Return ListBoxNt2.Items.Item(NumericUpDownBook.Value - 54)
            End If
        Catch ex As Exception
            LabelStatus.Text = "Error in finding Book name from Book number"
            errorLogger(ex)
            Return ""
        End Try
    End Function


    Private Sub NumericUpDownVerse_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownVerse.ValueChanged

        Try
            ' Setting value for 'Going to/Currently in' label
            Label22.Text = "Going to : "
            LabelFullName.Text = getBookNameFromBookNumber() + " " + NumericUpDownChapter.Value.ToString + ":" + NumericUpDownVerse.Value.ToString

            ' Set focus on the GO button
            ButtonGo.Focus()
        Catch ex As Exception
            LabelStatus.Text = "Error in using the entered Verse value"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub ListBoxRecent_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxRecent.SelectedIndexChanged
        Try
            renderVerseFromBookmarkEntry(ListBoxRecent.SelectedItem)
        Catch ex As Exception
            LabelStatus.Text = "Error in navigating to a verse from Recently Loaded verses"
            errorLogger(ex)
        End Try

    End Sub

    Private Sub ButtonAddBookmark_Click(sender As Object, e As EventArgs) Handles ButtonAddBookmark.Click
        Try
            ListBoxBookmark.Items.Add(ListBoxRecent.Items.Item(0))
        Catch ex As Exception
            LabelStatus.Text = "Error in adding bookmark"
            errorLogger(ex)
        End Try
    End Sub
    Private Sub errorLogger(ex As System.Exception)
        Try
            Dim currentDate As String
            currentDate = System.DateTime.Now.ToString("dd MMMM yyyy HH:mm:ss")
            Dim file As System.IO.StreamWriter

            file = My.Computer.FileSystem.OpenTextFileWriter(ValidateAndGetWorkspaceLocation() + "\log.txt", True)
            file.Write(currentDate + vbNewLine + ex.ToString + vbNewLine + vbNewLine)
            file.Close()
        Catch ex2 As Exception
            MsgBox("COULD NOT WRITE IN LOG FILE. THE FOLLOWING IS THE ERROR: " + Constants.vbCrLf + Constants.vbCrLf + ex.ToString)
            ' errorLogger(ex)
        End Try
    End Sub

    Private Sub ListBoxBookmark_Click(sender As Object, e As EventArgs) Handles ListBoxBookmark.Click
        Try
            'Validation - Needed because clicking in blank spaces (usually at the bottom) in the listbox results in random data
            If (ListBoxBookmark.SelectedIndex < 0) Then
                Exit Sub
            End If

            renderVerseFromBookmarkEntry(ListBoxBookmark.SelectedItem)
        Catch ex As Exception
            LabelStatus.Text = "Error in navigating to a bookmark"
            errorLogger(ex)
        End Try
    End Sub


    Private Sub renderVerseFromBookmarkEntry(bookmarkEntry As String)
        Try
            ' Retrieving book number from text such as "23 - PSALMS 1:3"
            Dim parsedTextBySpace As Array = bookmarkEntry.Split(" ")
            NumericUpDownBook.Value = parsedTextBySpace(0)

            ' Retrieving chapter and verse number from text such as "23 - PSALMS 1:3" i.e.,  parse the last entry from the array
            Dim parsedTextBySemicolon As Array = parsedTextBySpace(parsedTextBySpace.Length - 1).split(":")
            NumericUpDownChapter.Value = parsedTextBySemicolon(0)
            NumericUpDownVerse.Value = parsedTextBySemicolon(1)

            go(NumericUpDownBook.Text, NumericUpDownChapter.Text, NumericUpDownVerse.Text)
        Catch ex As Exception
            LabelStatus.Text = "Error in navigating to a verse from a list"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub ButtonDeleteBookmark_Click(sender As Object, e As EventArgs) Handles ButtonDeleteBookmark.Click
        Try
            ' Do nothing if no entry is selected
            If (ListBoxBookmark.SelectedIndex >= 0) Then
                ListBoxBookmark.Items.RemoveAt(ListBoxBookmark.SelectedIndex)
            End If
        Catch ex As Exception
            LabelStatus.Text = "Error in Deleting Bookmark"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub ButtonSaveSession_Click(sender As Object, e As EventArgs) Handles ButtonSaveSession.Click

        Try
            Dim file As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(ValidateAndGetWorkspaceLocation() + TextBoxSession.Text + ".mtbsession", False)
            For i = 1 To ListBoxSession.Items.Count
                file.WriteLine(ListBoxSession.Items.Item(i - 1))
            Next
            file.Flush()
            file.Close()
            MsgBox("Wrote session to file " + ValidateAndGetWorkspaceLocation() + TextBoxSession.Text + ".mtbsession")

            loadListOfSessions()
        Catch ex As Exception
            LabelStatus.Text = "Error in saving session"
            errorLogger(ex)
        End Try

    End Sub

    Private Sub TextBoxSession_TextChanged(sender As Object, e As EventArgs) Handles TextBoxSession.TextChanged
        Try
            ' Enable Save button only if a name is provided
            If (Not (Trim(TextBoxSession.Text) = "")) Then
                ButtonSaveSession.Enabled = True
            Else
                ButtonSaveSession.Enabled = False
            End If
        Catch ex As Exception
            LabelStatus.Text = "Error in toggling Save Session button using session name"
            errorLogger(ex)
        End Try

    End Sub

    Sub loadListOfSessions()
        Try
            Dim filename As Array
            Dim filenameTruncated As Array

            ' Lists all .mtbsession files in the workspace folder
            ComboBoxSessions.Items.Clear()
            Dim workspaceLocation = ValidateAndGetWorkspaceLocation()
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(workspaceLocation)

                If foundFile.EndsWith(".mtbsession") Then
                    ' Remove folder names
                    filename = foundFile.Split("\")
                    filenameTruncated = (filename(filename.Length - 1)).split(".mtbsession")
                    ComboBoxSessions.Items.Add(filenameTruncated(0))
                End If
            Next
        Catch ex As Exception
            LabelStatus.Text = "Error loading list of sessions. Config location = " + WORKSPACE_DIRECTORY_LOCATION
            errorLogger(ex)
        End Try

    End Sub

    Private Sub ButtonLoadSession_Click(sender As Object, e As EventArgs) Handles ButtonLoadSession.Click
        Try
            ' Clears the session ListBox for the new session
            ListBoxSession.Items.Clear()

            ' Read the selected .mtbsession file
            Dim filename As String = WORKSPACE_DIRECTORY_LOCATION + ComboBoxSessions.SelectedItem + ".mtbsession"
            Dim objReader As New System.IO.StreamReader(filename)
            Do While objReader.Peek() <> -1
                ListBoxSession.Items.Add(objReader.ReadLine())
            Loop
            objReader.Close()

            ' Disable Save button until user enters a name
            ButtonSaveSession.Enabled = False
            TextBoxSession.Text = ""

            ' FEATURE: Freeze session is enabled by default when you are using a session. You're most probably re-using a session and don't want to edit it.
            CheckBoxFreezeSession.Checked = True
        Catch ex As Exception
            LabelStatus.Text = "Error in loading session"
            errorLogger(ex)
        End Try

    End Sub

    Private Sub ButtonNewSession_Click(sender As Object, e As EventArgs) Handles ButtonNewSession.Click
        Try
            ' FEATURE: Freeze session is disabled when a new session is started. Most probably, you want to record this session.
            CheckBoxFreezeSession.Checked = False

            ' Clears the session ListBox for the new session
            ListBoxSession.Items.Clear()

            ' Disable Save button until user enters a name
            ButtonSaveSession.Enabled = False
            TextBoxSession.Text = ""
        Catch ex As Exception
            LabelStatus.Text = "Error in starting a new session"
            errorLogger(ex)
        End Try

    End Sub

    Private Sub ButtonSaveConfig_Click(sender As Object, e As EventArgs) Handles ButtonSaveConfig.Click
        Try
            Dim configValue As String = FormDisplay.Width.ToString + "|" + FormDisplay.Height.ToString + "|" + FormDisplay.Left.ToString + "|" + FormDisplay.Top.ToString
            My.Settings.Item(CONFIG_KEY) = configValue
            My.Settings.Save()
        Catch ex As Exception
            LabelStatus.Text = "Error in saving config"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub loadLayoutConfig()
        Try
            'Parses config string by splitting with |
            Dim configStringArray As Array = My.Settings.Item(CONFIG_KEY).Split("|")
            FormDisplay.Width = configStringArray(0)
            FormDisplay.Height = configStringArray(1)
            FormDisplay.Left = configStringArray(2)
            FormDisplay.Top = configStringArray(3)
        Catch ex As Exception
            LabelStatus.Text = "Error in applying config settings"
            errorLogger(ex)
        End Try

    End Sub

    Private Sub ButtonLoadConfig_Click(sender As Object, e As EventArgs) Handles ButtonLoadConfig.Click
        Try
            loadLayoutConfig()
        Catch ex As Exception
            LabelStatus.Text = "Error in loading config"
            errorLogger(ex)
        End Try
    End Sub


    Private Sub ButtonDeleteSession_Click(sender As Object, e As EventArgs) Handles ButtonDeleteSession.Click
        Try
            'Validation - Needed because clicking in blank spaces (usually at the bottom) in the listbox results in random data
            If (ComboBoxSessions.SelectedIndex < 0) Then
                Exit Sub
            End If

            My.Computer.FileSystem.DeleteFile(ValidateAndGetWorkspaceLocation() + ComboBoxSessions.Items.Item(ComboBoxSessions.SelectedIndex) + ".mtbsession")
            ComboBoxSessions.Items.RemoveAt(ComboBoxSessions.SelectedIndex)

            ' Clearing session contents
            ListBoxSession.Items.Clear()

            ' Clearing list of available sessions (Setting SelectedIndex=-1 doesn't clear the text in the combo box, so have to use this)
            ComboBoxSessions.Text = ""
            loadListOfSessions()
        Catch ex As Exception
            LabelStatus.Text = "Error in deleting session"
            errorLogger(ex)
        End Try
    End Sub



    Function regexForNumberOfVerses(ByVal text As String, ByVal expr As String)
        Try
            'FUNCTION THAT LOOKS FOR REGEX AND COUNTS OCCURENCES
            Dim count As Integer = 0
            Dim mc As MatchCollection = Regex.Matches(text, expr)
            Dim m As Match
            For Each m In mc
                count = count + 1
            Next m
            Return count
        Catch ex As Exception
            LabelStatus.Text = "Error in locating verse pattern in display"
            errorLogger(ex)
        End Try

        Return vbNull
    End Function

    Private Sub renderNumberOfVersesInCurrentChapter()
        Try
            RichTextBoxDummy.Rtf = GetRtfFile(NumericUpDownBook.Value.ToString, NumericUpDownChapter.Value.ToString)
            Dim count As Integer = regexForNumberOfVerses(RichTextBoxDummy.Text, "\b:\d?\S")
            LabelVerses.Text = count
            NumericUpDownVerse.Maximum = count
        Catch ex As Exception
            LabelStatus.Text = "Error in finding number of verses"
            errorLogger(ex)
        End Try
    End Sub


    Private Sub ButtonAbout_Click(sender As Object, e As EventArgs) Handles ButtonAbout.Click
        Try
            Dim text As String = "Maximized Tamil Bible v4.0"
            text = text + vbNewLine + "Released in March 2020"
            text = text + vbNewLine + "For queries, feedback and suggestions, mail sjoedanny@gmail.com."

            MsgBox(text, MsgBoxStyle.Information, AcceptButton)
        Catch ex As Exception
            LabelStatus.Text = "Error in showing About Info"
            errorLogger(ex)
        End Try

    End Sub

    Private Sub ButtonHelp_Click(sender As Object, e As EventArgs) Handles ButtonHelp.Click
        Try
            Dim text As String = "This software works only on Windows. Latest version of .NET Framework is required."
            text = text + vbNewLine + "If there is a font problem, install Tamil Bible font in your system and try again."
            text = text + vbNewLine + "Tamil Bible font is provided with this package or you can find it online."
            text = text + vbNewLine + vbNewLine + "For queries, feedback and suggestions, mail sjoedanny@gmail.com."
            MsgBox(text, MsgBoxStyle.Information, AcceptButton)
        Catch ex As Exception
            LabelStatus.Text = "Error in showing Help Info"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub ButtonFind_Click(sender As Object, e As EventArgs) Handles ButtonFind.Click
        Try
            Dim chapterText As String = RichTextBoxMiniDisplay.Text
            Dim textToFind As String = TextBoxFind.Text

            ' TODO: This assumes each verse ends with a period. This has to be improved.
            Dim tokenizedVerses As Array = chapterText.Split(".")

            Dim foundVerses As String = ""

            For i = 0 To tokenizedVerses.Length - 1
                ' If a result is found, it is appended to the result
                If (tokenizedVerses(i).ToString.Contains(textToFind)) Then
                    foundVerses = foundVerses + tokenizedVerses(i) + vbNewLine + "------------------"
                End If
            Next

            RichTextBoxFindResults.Text = foundVerses
        Catch ex As Exception
            LabelStatus.Text = "Error in find function"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub NumericUpDownChapter_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NumericUpDownChapter.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            go(NumericUpDownBook.Text, NumericUpDownChapter.Text, NumericUpDownVerse.Text)
        End If
    End Sub

    Private Sub NumericUpDownBook_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NumericUpDownBook.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            go(NumericUpDownBook.Text, NumericUpDownChapter.Text, NumericUpDownVerse.Text)
        End If
    End Sub

    Private Sub NumericUpDownVerse_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NumericUpDownVerse.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            go(NumericUpDownBook.Text, NumericUpDownChapter.Text, NumericUpDownVerse.Text)
        End If
    End Sub

    Private Sub ListBoxSession_Click(sender As Object, e As EventArgs) Handles ListBoxSession.Click
        Try
            'Validation - Needed because clicking in blank spaces (usually at the bottom) in the listbox results in random data
            If (ListBoxSession.SelectedIndex < 0) Then
                Exit Sub
            End If

            renderVerseFromBookmarkEntry(ListBoxSession.SelectedItem)
        Catch ex As Exception
            LabelStatus.Text = "Error in navigating to a bookmark"
            errorLogger(ex)
        End Try

    End Sub

    Private Sub ButtonSetZoom_Click(sender As Object, e As EventArgs)
        Dim zoomFactor As Double = NumericUpDownZoom.Value / 100.0
        FormDisplay.RichTextBox1.ZoomFactor = zoomFactor
    End Sub
End Class
