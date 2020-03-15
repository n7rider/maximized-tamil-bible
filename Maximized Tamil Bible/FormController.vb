Imports System.Text.RegularExpressions

Public Class FormController
    Dim currentBook As Integer
    Dim currentChapter As Integer
    Dim currentVerse As Integer
    Dim chaptersInEachBook() As Integer = {50, 40, 27, 36, 34, 24, 21, 4, 31, 24, 22, 25, 29, 36, 10, 13, 10, 42, 150, 31, 12, 8, 66, 52, 5, 48, 12, 14, 3, 9, 1, 4, 7, 3, 3, 3, 2, 14, 4, 28, 16, 24, 21, 28, 16, 16, 13, 6, 6, 4, 4, 5, 3, 6, 4, 3, 1, 13, 5, 5, 3, 5, 1, 1, 1, 22}
    Private Sub ListBoxOt1_Click(sender As Object, e As EventArgs) Handles ListBoxOt1.Click
        
        Try
            'COPYING TO TEMPVALUE AND USING IT TO AVOID GETTING -1 AS SELECTEDINDEX WHICH HAPPENS WHEN CLICKED IN BLANK AREA OF LIST BOX
            'GIVING SELECTEDINDEX VALUE TO UPDOWNBOOK
            Dim tempValue As Integer
            tempValue = ListBoxOt1.SelectedIndex + 1
            If (tempValue < 1) Then
                NumericUpDownBook.Value = 1
            Else
                NumericUpDownBook.Value = tempValue
            End If

            NumericUpDownChapter.Focus()

            'DESELECTING SELECTIONS IN OTHER LIST BOXES
            ListBoxOt2.SetSelected(0, False)
            ListBoxNt1.SetSelected(0, False)
            ListBoxNt2.SetSelected(0, False)

            'COPYING TO TEMPVALUE AND USING IT TO AVOID GETTING -1 AS SELECTEDINDEX WHICH HAPPENS WHEN CLICKED IN BLANK AREA OF LIST BOX
            'GIVING SELECTEDINDEX VALUE TO FEED NO-OF-CHAPTERS-IN-THIS-BOOK LABEL
            tempValue = ListBoxOt1.SelectedIndex
            If (tempValue < 0) Then
                tempValue = 0
            End If
            LabelChapters.Text = chaptersInEachBook(tempValue)


            'GIVING SELECTEDINDEX VALUE TO LIMIT NUYMERICUPDOWN-CHAPTER'S MAXIMUM
            NumericUpDownChapter.Maximum = chaptersInEachBook(tempValue)

        Catch ex As Exception
            LabelStatus.Text = "Error in selecting book from Old Testament - First Half"
            errorLogger(ex)
        End Try
        

       
    End Sub

    Private Sub ListBoxOt2_Click(sender As Object, e As EventArgs) Handles ListBoxOt2.Click
        Try
            'GIVING SELECTEDINDEX VALUE TO UPDOWNBOOK
            NumericUpDownBook.Value = ListBoxOt2.SelectedIndex + 20
            NumericUpDownChapter.Focus()

            'DESELECTING SELECTIONS IN OTHER LIST BOXES
            ListBoxOt1.SetSelected(0, False)
            ListBoxNt1.SetSelected(0, False)
            ListBoxNt2.SetSelected(0, False)

            'GIVING SELECTEDINDEX VALUE TO FEED NO-OF-CHAPTERS-IN-THIS-BOOK LABEL
            LabelChapters.Text = chaptersInEachBook(ListBoxOt2.SelectedIndex + 19)

            'GIVING SELECTEDINDEX VALUE TO LIMIT NUYMERICUPDOWN-CHAPTER'S MAXIMUM
            NumericUpDownChapter.Maximum = chaptersInEachBook(ListBoxOt2.SelectedIndex + 19)
        Catch ex As Exception
            LabelStatus.Text = "Error in selecting book from Old Testament - Second Half"
            errorLogger(ex)
        End Try
       
    End Sub

    Private Sub ListBoxNt1_Click(sender As Object, e As EventArgs) Handles ListBoxNt1.Click
        Try
            'GIVING SELECTEDINDEX VALUE TO UPDOWNBOOK
            NumericUpDownBook.Value = ListBoxNt1.SelectedIndex + 40
            NumericUpDownChapter.Focus()

            'DESELECTING SELECTIONS IN OTHER LIST BOXES
            ListBoxOt1.SetSelected(0, False)
            ListBoxOt2.SetSelected(0, False)
            ListBoxNt2.SetSelected(0, False)

            'GIVING SELECTEDINDEX VALUE TO FEED NO-OF-CHAPTERS-IN-THIS-BOOK LABEL
            LabelChapters.Text = chaptersInEachBook(ListBoxNt1.SelectedIndex + 39)

            'GIVING SELECTEDINDEX VALUE TO LIMIT NUYMERICUPDOWN-CHAPTER'S MAXIMUM
            NumericUpDownChapter.Maximum = chaptersInEachBook(ListBoxNt1.SelectedIndex + 39)
        Catch ex As Exception
            LabelStatus.Text = "Error in selecting book from New Testament - First Half"
            errorLogger(ex)
        End Try
       
    End Sub

    Private Sub ListBoxNt2_Click(sender As Object, e As EventArgs) Handles ListBoxNt2.Click
        Try
            'GIVING SELECTEDINDEX VALUE TO UPDOWNBOOK
            NumericUpDownBook.Value = ListBoxNt2.SelectedIndex + 54
            NumericUpDownChapter.Focus()

            'DESELECTING SELECTIONS IN OTHER LIST BOXES
            ListBoxOt1.SetSelected(0, False)
            ListBoxOt2.SetSelected(0, False)
            ListBoxNt1.SetSelected(0, False)

            'GIVING SELECTEDINDEX VALUE TO FEED NO-OF-CHAPTERS-IN-THIS-BOOK LABEL
            LabelChapters.Text = chaptersInEachBook(ListBoxNt2.SelectedIndex + 53)

            'GIVING SELECTEDINDEX VALUE TO LIMIT NUYMERICUPDOWN-CHAPTER'S MAXIMUM
            NumericUpDownChapter.Maximum = chaptersInEachBook(ListBoxNt2.SelectedIndex + 53)
        Catch ex As Exception
            LabelStatus.Text = "Error in selecting book from New Testament - Second Half"
            errorLogger(ex)
        End Try
       
    End Sub



    Private Sub ButtonGo_Click(sender As Object, e As EventArgs) Handles ButtonGo.Click
        Try
            FormDisplay.Show()

            'GO FUNCTION FETCHES RTF AND DISPLAYS IT
            go(NumericUpDownBook.Text, NumericUpDownChapter.Text, NumericUpDownVerse.Text)
        Catch ex As Exception
            LabelStatus.Text = "Error in Navigating to the selected verse"
            errorLogger(ex)
        End Try
       



    End Sub


    Sub go(Book, Optional Chapter = 1, Optional Verse = 1)
        Try
            Dim selectionStart As Integer

            'KEEP CURRENT BOOK, CURRENT CHAPTER, CURRENT VERSE IN A VARIABLE
            currentBook = NumericUpDownBook.Value
            currentChapter = NumericUpDownChapter.Value
            currentVerse = NumericUpDownVerse.Value

            'FINDING BOOKNAME FROM BOOK NO. USING LISTBOXES. AVOIDING OVERHEADS AND WORKAROUNDS
            findBookNameFromBookNumber()

            'GENERATE LIVE LABEL TEXT - BOOK, CHAPTER, VERSE
            LabelFullName.Text = LabelFullName.Text + " " + NumericUpDownChapter.Value.ToString + ":" + NumericUpDownVerse.Value.ToString

            'CHANGE STATUS LABEL TO CURRENTLY-IN
            Label22.Text = "Currently In : "

            'LOADS VERSE
            FormDisplay.RichTextBox1.LoadFile("C:/Program Files/Maximized Tamil Bible/Bible/" + Book + "/" + Chapter + ".rtf")

            'ADDS DISPLAYED VERSE TO TOP OF LISTBOX-RECENT
            If (ListBoxRecent.Items.Count > 14) Then
                ListBoxRecent.Items.RemoveAt(14)
            End If
            ListBoxRecent.Items.Insert(0, LabelFullName.Text)

            'ADDS DISPLAYED VERSE TO SESSION IF FREEZE NOT ENABLED
            If (CheckBoxFreezeSession.Checked = False) Then
                ListBoxSession.Items.Insert(0, LabelFullName.Text)
            End If

            'SELECTS ALL TEXT AND SETS SELECTED FONT SIZE
            FormDisplay.RichTextBox1.SelectAll()
            FormDisplay.RichTextBox1.SelectionFont = New Font(FormDisplay.RichTextBox1.SelectionFont.FontFamily, Int(NumericUpDownZoom.Value))

            'HIGHLIGHTS SELECTED VERSE IN DISPLAY
            Dim searchString As String
            searchString = Chapter + ":" + Verse + " "
            selectionStart = FormDisplay.RichTextBox1.Find(searchString)
            FormDisplay.RichTextBox1.Select(selectionStart, searchString.Length)

            'COPYING TEXT FROM DISPLAY TO MINI DISPLAY
            searchString = Chapter + ":" + Verse + " "
            RichTextBoxMiniDisplay.Rtf = FormDisplay.RichTextBox1.Rtf

            'SETTING MINI DISPLAY'S FONT TO 12
            RichTextBoxMiniDisplay.SelectAll()
            RichTextBoxMiniDisplay.SelectionFont = New Font(FormDisplay.RichTextBox1.SelectionFont.FontFamily, Int(12))

            'HIGHLIGHTS SELECTED VERSE IN MINI DISPLAY
            selectionStart = RichTextBoxMiniDisplay.Find(searchString)
            RichTextBoxMiniDisplay.Select(selectionStart, searchString.Length)

           

            RichTextBoxMiniDisplay.Focus()

            'FOCUS ON VERSE
            FormDisplay.RichTextBox1.Focus()

        Catch ex As Exception
            MsgBox(ex.ToString + "hi")
            LabelStatus.Text = "Error in go method while navigating to the selected verse"
            errorLogger(ex)
        End Try
       

    End Sub

    Private Sub FormController_Click(sender As Object, e As EventArgs) Handles Me.Click
        Try
            'CLICKING FORM FOCUSES TO BOOK-UPDOWN
            NumericUpDownBook.Focus()
        Catch ex As Exception
            LabelStatus.Text = "Error in Focusing on Book Input Field"
            errorLogger(ex)
        End Try
       
    End Sub

    Private Sub Controller_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'ON-LOAD OF FORM
            FormDisplay.Show()

            'SESSION INIT
            loadListOfSessions()

            'CONFIG INIT
            loadListOfConfigs()

            'LOAD CONFIG IN USE
            configLoader()

            'CHANGING MINIMUM FROM 0 TO 1. 
            'THIS CHANGE MAKES GENESIS 1:1 GETS ITS VERSE VALUE AS NUMERICUPDOWN_VALUECHANGED IS RUN
            NumericUpDownBook.Minimum = 1
            NumericUpDownChapter.Minimum = 1
            NumericUpDownVerse.Minimum = 1

            'SET FONT SIZE IN FONT-SIZE-SAMPLE-RTB
            setSizeInSample()

            NumericUpDownBook.Focus()
        Catch ex As Exception
            LabelStatus.Text = "Error while loading the form"
            errorLogger(ex)
        End Try
       

    End Sub

    Private Sub NumericUpDownChapter_GotFocus(sender As Object, e As EventArgs) Handles NumericUpDownChapter.GotFocus
        Try
            'SELECT ALL ON FOCUS
            NumericUpDownChapter.Select(0, NumericUpDownChapter.Value.ToString.Length)
        Catch ex As Exception
            LabelStatus.Text = "Error in focusing on Chapter Input Field"
            errorLogger(ex)
        End Try
       
    End Sub


    Private Sub NumericUpDownVerse_GotFocus(sender As Object, e As EventArgs) Handles NumericUpDownVerse.GotFocus
        Try
            'SELECT ALL ON FOCUS
            NumericUpDownVerse.Select(0, NumericUpDownVerse.Value.ToString.Length)
        Catch ex As Exception
            LabelStatus.Text = "Error in focusing on Verse Input Field"
            errorLogger(ex)
        End Try
      


    End Sub

    Private Sub NumericUpDownBook_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownBook.ValueChanged
        Try
            'FINDING BOOKNAME FROM BOOK NO. USING LISTBOXES. AVOIDING OVERHEADS AND WORKAROUNDS
            findBookNameFromBookNumber()

            'CHANGE STATUS LABEL TO CURRENTLY-GOING-TO
            Label22.Text = "Going to : "

            'FOCUS ON CHAPTER AFTER BOOK HAS BEEN CHANGED BY LISTBOXES
            NumericUpDownChapter.Focus()

            Dim tempvalue As Integer
            'GIVING SELECTEDINDEX VALUE TO FEED NO-OF-CHAPTERS-IN-THIS-BOOK LABEL
            tempvalue = NumericUpDownBook.Value
            If (tempvalue < 1) Then
                tempvalue = 1
            End If

            '-1 IS USED SINCE THE ARRAY COUNTS FROM 0
            LabelChapters.Text = chaptersInEachBook(tempvalue - 1)

            'GIVING SELECTEDINDEX VALUE TO LIMIT NUYMERICUPDOWN-CHAPTER'S MAXIMUM
            NumericUpDownChapter.Maximum = chaptersInEachBook(tempvalue)

            'AVOID CHAPTER NO. BECOMING LARGER THAN THE NO. OF CHAPTERS IN THE BOOK (OR) RESETTING CHAPTER NO. TO 1 IF BOOK CHANGES
            If (CheckBoxFirstChapterSetter.Checked = True) Then
                NumericUpDownChapter.Value = 1
            Else
                If (NumericUpDownChapter.Value > Val(LabelChapters.Text)) Then
                    NumericUpDownChapter.Value = Val(LabelChapters.Text)
                End If
            End If

            'LOADS NO OF VERSES IN LABEL-VERSES
            If (NumericUpDownChapter.Value < 1) Then
                NumericUpDownChapter.Value = 1
            End If
            showNoOfVerses()
        Catch ex As Exception
            LabelStatus.Text = "Error in using the entered Book value"
            errorLogger(ex)
        End Try
       

    End Sub

    Private Sub NumericUpDownChapter_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownChapter.ValueChanged

        Try
            'FINDING BOOKNAME FROM BOOK NO. USING LISTBOXES. AVOIDING OVERHEADS AND WORKAROUNDS
            findBookNameFromBookNumber()

            'GENERATE LIVE LABEL TEXT - BOOK, CHAPTER
            LabelFullName.Text = LabelFullName.Text + " " + NumericUpDownChapter.Value.ToString + ":"

            'CHANGE STATUS LABEL TO CURRENTLY-GOING-TO
            Label22.Text = "Going to : "

            'FOCUS ON VERSE AFTER CHAPTER HAS BEEN CHANGED 
            NumericUpDownVerse.Focus()

            'SHOWS NO OF VERSES IN LABEL-VERSES
            showNoOfVerses()
        Catch ex As Exception
            LabelStatus.Text = "Error in using the entered Chapter value"
            errorLogger(ex)
        End Try
        


    End Sub

    Sub findBookNameFromBookNumber()
        Try
            'FINDING BOOKNAME FROM BOOK NO. USING LISTBOXES. AVOIDING OVERHEADS AND WORKAROUNDS
            If (NumericUpDownBook.Value < 20) Then
                LabelFullName.Text = ListBoxOt1.Items.Item(NumericUpDownBook.Value - 1)
            ElseIf (NumericUpDownBook.Value < 40) Then
                LabelFullName.Text = ListBoxOt2.Items.Item(NumericUpDownBook.Value - 20)
            ElseIf (NumericUpDownBook.Value < 54) Then
                LabelFullName.Text = ListBoxNt1.Items.Item(NumericUpDownBook.Value - 40)
            Else
                LabelFullName.Text = ListBoxNt2.Items.Item(NumericUpDownBook.Value - 54)
            End If
        Catch ex As Exception
            LabelStatus.Text = "Error in finding Book name from Book number"
            errorLogger(ex)
        End Try
       
    End Sub

    Private Sub NumericUpDownVerse_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownVerse.ValueChanged

        Try
            'FINDING BOOKNAME FROM BOOK NO. USING LISTBOXES. AVOIDING OVERHEADS AND WORKAROUNDS
            findBookNameFromBookNumber()

            'GENERATE LIVE LABEL TEXT - BOOK, CHAPTER, VERSE
            LabelFullName.Text = LabelFullName.Text + " " + NumericUpDownChapter.Value.ToString + ":" + NumericUpDownVerse.Value.ToString

            'CHANGE STATUS LABEL TO CURRENTLY-GOING-TO
            Label22.Text = "Going to : "

            'FOCUS ON GO AFTER VERSE HAS BEEN CHANGED 
            ButtonGo.Focus()
        Catch ex As Exception
            LabelStatus.Text = "Error in using the entered Verse value"
            errorLogger(ex)
        End Try
        


    End Sub

    Private Sub ListBoxRecent_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxRecent.SelectedIndexChanged
        Try
            'GO TO VERSE FROM RECENT LIST
            goToVerseFromList(ListBoxRecent)
        Catch ex As Exception
            LabelStatus.Text = "Error in navigating to a verse from Recently Loaded verses"
            errorLogger(ex)
        End Try
       
    End Sub

    Private Sub ButtonAddBookmark_Click(sender As Object, e As EventArgs) Handles ButtonAddBookmark.Click
        Try
            'ADDS CURRENT VERSE TO BOOKMARKS
            ListBoxBookmark.Items.Add(ListBoxRecent.Items.Item(0))

            'SAVES BOOKMARK IN AN INTERNAL FILE
            Dim file As System.IO.StreamWriter
            file = My.Computer.FileSystem.OpenTextFileWriter("C:\Program Files\Maximized Tamil Bible\bookmarks.txt", True)
            For i = 1 To ListBoxBookmark.Items.Count
                file.WriteLine(ListBoxBookmark.Items.Item(i - 1))
            Next
            file.Close()
            MsgBox(System.DateTime.Now.ToString("dd MMMM yyyy HH:mm:ss"))
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
            file = My.Computer.FileSystem.OpenTextFileWriter("C:\Program Files\Maximized Tamil Bible\log.txt", True)
            file.Write(currentDate + vbNewLine + ex.ToString + vbNewLine + vbNewLine)
            file.Close()
        Catch ex2 As Exception
            MsgBox("COULD NOT WRITE IN LOG FILE. THE FOLLOWING IS THE ERROR: " + ex.ToString)
            errorLogger(ex)
        End Try
    End Sub

    Private Sub ButtonGoBookmark_Click(sender As Object, e As EventArgs) Handles ButtonGoBookmark.Click
        Try
            'GO TO THE VERSE SELECTED IN THE BOOKMARK
            goToVerseFromList(ListBoxBookmark)
        Catch ex As Exception
            LabelStatus.Text = "Error in navigating to a bookmark"
            errorLogger(ex)
        End Try
       
    End Sub


    Private Sub goToVerseFromList(listBox As ListBox)
        Try
            Dim parsedText As Array
            Dim tempValue As String
            Dim parsedText2 As Array

            'CLIPPING BOOK NUMBER FROM TEXT SUCH AS '23 - PSALMS 1:3'
            tempValue = listBox.SelectedItem
            parsedText = tempValue.Split(" ")
            NumericUpDownBook.Value = parsedText(0)

            'FROM THE FINAL PART OF THE TEXT SUCH AS '23 - PSALMS 1:3' TAKE 1:3 ALONE AND THEN TAKE 1 AND 3 SEPARATELY FOR CHAPTER AND VERSE RESPECTIVELY
            parsedText2 = parsedText(parsedText.Length - 1).split(":")
            NumericUpDownChapter.Value = parsedText2(0)
            NumericUpDownVerse.Value = parsedText2(1)

            go(NumericUpDownBook.Text, NumericUpDownChapter.Text, NumericUpDownVerse.Text)
        Catch ex As Exception
            LabelStatus.Text = "Error in navigating to a verse from a list"
            errorLogger(ex)
        End Try
       
    End Sub

    Private Sub ButtonDeleteBookmark_Click(sender As Object, e As EventArgs) Handles ButtonDeleteBookmark.Click

        Try
            'DELETE A VERSE FROM BOOKMARK; CATCH EXCEPTION IF USER HAS SELECTED NOTHING AND PRESSED DELETE
            ListBoxBookmark.Items.RemoveAt(ListBoxBookmark.SelectedIndex)
        Catch ex As Exception
            LabelStatus.Text = "Error in Deleting Bookmark"
            errorLogger(ex)
        End Try

    End Sub

    Private Sub ListBoxSession_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxSession.SelectedIndexChanged
        Try
            'LOAD THE SELECTED VERSE IN THIS SESSION
            goToVerseFromList(ListBoxSession)
        Catch ex As Exception
            LabelStatus.Text = "Error in navigating to a session"
            errorLogger(ex)
        End Try
       
    End Sub

    Private Sub ButtonSaveSession_Click(sender As Object, e As EventArgs) Handles ButtonSaveSession.Click

        Try
            'SAVES EVERYTHING IN LISTBOXSESSION TO THE FILE WITH NAME AS IN TEXTBOXSESSION
            Dim file As System.IO.StreamWriter
            file = My.Computer.FileSystem.OpenTextFileWriter("C:\Program Files\Maximized Tamil Bible\" + TextBoxSession.Text + ".mtbsession", True)
            For i = 1 To ListBoxSession.Items.Count
                file.WriteLine(ListBoxSession.Items.Item(i - 1))
            Next
            file.Close()

            loadListOfSessions()
        Catch ex As Exception
            LabelStatus.Text = "Error in saving session"
            errorLogger(ex)
        End Try
       
    End Sub

    Private Sub TextBoxSession_TextChanged(sender As Object, e As EventArgs) Handles TextBoxSession.TextChanged
        Try
            'ENABLE SAVE SESSON BUTTON ONLY IF NAME IS ENTERED IN TEXTBOX SESSION
            If (Not (Trim(TextBoxSession.Text) = "")) Then
                ButtonSaveSession.Enabled = True
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

            'LISTS ALL .MTBSESSION FILES IN THE LOCAL FOLDER TO THE COMBOBOX
            ComboBoxSessions.Items.Clear()
            For Each foundFile As String In My.Computer.FileSystem.GetFiles("C:\Program Files\Maximized Tamil Bible")
                If foundFile.EndsWith(".mtbsession") Then
                    'TAKES ONLY THE NAME OF THE .MTBSESSION AND STRIPS OFF THE FOLDER NAMES
                    filename = foundFile.Split("\")
                    filenameTruncated = (filename(filename.Length - 1)).split(".mtbsession")
                    ComboBoxSessions.Items.Add(filenameTruncated(0))
                End If
            Next
        Catch ex As Exception
            LabelStatus.Text = "Error in loading list of sessions"
            errorLogger(ex)
        End Try
       
    End Sub

    Private Sub ButtonLoadSession_Click(sender As Object, e As EventArgs) Handles ButtonLoadSession.Click
        Try
            'CLEARS THE VERSES IN THE CURRENT SESSION
            ListBoxSession.Items.Clear()

            'READ THE SELECTED .MTBSESSION LINE BY LINE AND PUT THE LINES IN LISTBOXSESSIONS
            Dim filename As String = "C:\Program Files\Maximized Tamil Bible\" + ComboBoxSessions.SelectedItem + ".mtbsession"
            Dim objReader As New System.IO.StreamReader(filename)
            Do While objReader.Peek() <> -1
                ListBoxSession.Items.Add(objReader.ReadLine())
            Loop

            objReader.Close()

            'DISABLES SAVE BUTTON FOR NOW; NEED TO TYPE A NAME IN TEXTBOXSESSION TO ENABLE IT
            ButtonSaveSession.Enabled = False
            TextBoxSession.Text = ""

            'FREEZE LISTBOX MODIFICATION IF YOU HAVE LOADED A SESSION
            CheckBoxFreezeSession.Checked = True
        Catch ex As Exception
            LabelStatus.Text = "Error in loading session"
            errorLogger(ex)
        End Try
        
    End Sub

    Private Sub ButtonNewSession_Click(sender As Object, e As EventArgs) Handles ButtonNewSession.Click
        Try
            'STOP FREEZING MODIFICATION
            CheckBoxFreezeSession.Checked = False

            'CLEAR ALL LINES IN LISTBOXSESSION
            ListBoxSession.Items.Clear()
            TextBoxSession.Text = ""

            'DISABLES SAVE BUTTON FOR NOW; NEED TO TYPE A NAME IN TEXTBOXSESSION TO ENABLE IT
            ButtonSaveSession.Enabled = False
        Catch ex As Exception
            LabelStatus.Text = "Error in starting a new session"
            errorLogger(ex)
        End Try
        
    End Sub

    Private Sub configLoader()
        Try
            '.MASTERCONFIG STORES WHICH .CONFIG IS THE DEFAULT CONFIG; RETRIEVE THE .CONFIG NAME
            Dim configFile As String
            Dim filename As String = "C:\Program Files\Maximized Tamil Bible\mtb.masterconfig"
            Dim objReader1 As New System.IO.StreamReader(filename)
            configFile = objReader1.ReadLine()
            objReader1.Close()

            'LOAD SETTINGS IN THAT .CONFIG FILE
            LoadConfigAndApply(configFile)
        Catch ex As Exception
            LabelStatus.Text = "Error in loading default config from master config"
            errorLogger(ex)
        End Try
       

    End Sub

    Private Sub ButtonSaveConfig_Click(sender As Object, e As EventArgs) Handles ButtonSaveConfig.Click
        Try
            Dim file As System.IO.StreamWriter

            'NOT ALLOWING TO SAVE WITH THE NAME DEFAULT
            If (TextBoxConfig.Text.ToUpper = "DEFAULT") Then
                MsgBox("Error. Cannot save with the name default. Specify another name.")
            Else
                file = My.Computer.FileSystem.OpenTextFileWriter("C:\Program Files\Maximized Tamil Bible\" + TextBoxConfig.Text + ".config", False)

                'SAVE FORMDISPLAY'S DIMENSION SETTINGS TO A NEW CONFIG, NAME OF THE CONFIG IS FROM TEXTBOXCONFIG
                file.WriteLine(FormDisplay.Height)
                file.WriteLine(FormDisplay.Width)
                file.WriteLine(FormDisplay.Top)
                file.WriteLine(FormDisplay.Left)

                file.Close()

                loadListOfConfigs()
            End If
        Catch ex As Exception
            LabelStatus.Text = "Error in saving config"
            errorLogger(ex)
        End Try
       

    End Sub

    Private Sub TextBoxConfig_TextChanged(sender As Object, e As EventArgs) Handles TextBoxConfig.TextChanged
        Try
            'ENABLE SAVE CONFIG BUTTON ONLY IF SOME TEXT IS ENTERED IN TEXTBOXCONFIG
            If (Not Trim(TextBoxConfig.Text) = "") Then
                ButtonSaveConfig.Enabled = True
            End If
        Catch ex As Exception
            LabelStatus.Text = "Error in toggling Save COnfig button using Config Name"
            errorLogger(ex)
        End Try
       

    End Sub

    Private Sub loadListOfConfigs()
        Try
            Dim filename As Array
            Dim filenameTruncated As Array

            ComboBoxConfigs.Items.Clear()

            'LOAD LIST OF CONFIGS EXCEPT DEFAULT
            For Each foundFile As String In My.Computer.FileSystem.GetFiles("C:\Program Files\Maximized Tamil Bible")
                If (foundFile.EndsWith(".config") And Not foundFile.EndsWith("default.config")) Then
                    filename = foundFile.Split("\")
                    filenameTruncated = (filename(filename.Length - 1)).split(".config")
                    ComboBoxConfigs.Items.Add(filenameTruncated(0))
                End If
            Next
        Catch ex As Exception
            LabelStatus.Text = "Error in loading list of configs"
            errorLogger(ex)
        End Try
       
    End Sub

    Private Sub ButtonLoadConfig_Click(sender As Object, e As EventArgs) Handles ButtonLoadConfig.Click
        Try
            'LOAD AND APPLY SETTINGS FROM SELECTED ITEM IN COMBOBOX CONFIG
            LoadConfigAndApply(ComboBoxConfigs.SelectedItem)
        Catch ex As Exception
            LabelStatus.Text = "Error in loading config"
            errorLogger(ex)
        End Try
       
    End Sub


    Private Sub ButtonLoadDefaultConfig_Click(sender As Object, e As EventArgs) Handles ButtonLoadDefaultConfig.Click
        Try
            'LOAD AND APPLY SETTINGS FROM DEFAULT.CONFIG
            LoadConfigAndApply("default.config")
        Catch ex As Exception
            LabelStatus.Text = "Error in loading default config"
            errorLogger(ex)
        End Try
       
    End Sub


    Private Sub NumericUpDownZoom_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownZoom.ValueChanged
        Try
            'CALLING FUNCTION TO CHANGE FONT SIZE IN SAMPLE-RTB
            setSizeInSample()
        Catch ex As Exception
            LabelStatus.Text = "Error in adjusting display zoom"
            errorLogger(ex)
        End Try
       
    End Sub

    Private Sub setSizeInSample()
        Try
            'INSTEAD OF USING FONT SIZE, USING ZOOM. ZOOM IS ONLY FROM 1 TO 5 BUT ALLOWS DECIMAL. SO I USE 1 TO 50 AND DIVIDE BY 10.
            RichTextBoxSample.SelectAll()
            RichTextBoxSample.SelectionFont = New Font(RichTextBoxSample.SelectionFont.FontFamily, Int(NumericUpDownZoom.Value))
            RichTextBoxSample.DeselectAll()
        Catch ex As Exception
            LabelStatus.Text = "Error in adjusting display zoom"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub ButtonDeleteConfig_Click(sender As Object, e As EventArgs) Handles ButtonDeleteConfig.Click
        Try
            'DELETE THE SELECTED CONFIG. CATCHING EXCEPTION IF NOTHING WAS SELECTED
            Try
                My.Computer.FileSystem.DeleteFile("C:\Program Files\Maximized Tamil Bible\" + ComboBoxConfigs.Items.Item(ComboBoxConfigs.SelectedIndex) + ".config")
                ComboBoxConfigs.Items.RemoveAt(ComboBoxConfigs.SelectedIndex)
            Catch ex As Exception

            End Try
        Catch ex As Exception
            LabelStatus.Text = "Error in deleting config"
            errorLogger(ex)
        End Try
       
    End Sub


    Private Sub ButtonDeleteSession_Click(sender As Object, e As EventArgs) Handles ButtonDeleteSession.Click
        Try
            'DELETE THE SELECTED SESSION. CATCHING EXCEPTION IF NOTHING WAS SELECTED
            My.Computer.FileSystem.DeleteFile("C:\Program Files\Maximized Tamil Bible\" + ComboBoxSessions.Items.Item(ComboBoxSessions.SelectedIndex) + ".mtbsession")
            ComboBoxSessions.Items.RemoveAt(ComboBoxSessions.SelectedIndex)
            ListBoxSession.Items.Clear()
        Catch ex As Exception
            LabelStatus.Text = "Error in deleting session"
            errorLogger(ex)
        End Try
    End Sub



    Function regexForNumberOfVerses(ByVal text As String, ByVal expr As String)
        Try
            'FUNCTION THAT LOOKS FOR REGEX AND COUNTS OCCURENCES
            Dim count As Integer
            count = 0
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



    Private Sub LoadConfigAndApply(configFile As String)
        Try
            'LOADS THE LINES IN .CONFIG AND APPLIES THE PROPERTIES TO THE APPLICATION

            Dim count As Integer
            Dim filename As String = "C:\Program Files\Maximized Tamil Bible\" + configFile + ".config"
            Dim objReader As New System.IO.StreamReader(filename)
            count = 0
            Do While objReader.Peek() <> -1
                If (count = 0) Then
                    FormDisplay.Height = Val(objReader.ReadLine)
                ElseIf (count = 1) Then
                    FormDisplay.Width = Val(objReader.ReadLine)
                ElseIf (count = 2) Then
                    FormDisplay.Top = Val(objReader.ReadLine)
                ElseIf (count = 3) Then
                    FormDisplay.Left = Val(objReader.ReadLine)
                End If
                count = count + 1
            Loop
            objReader.Close()

            ButtonSaveConfig.Enabled = False
            TextBoxConfig.Text = ""
        Catch ex As Exception
            LabelStatus.Text = "Error in applying selected config settings"
            errorLogger(ex)
        End Try
       
    End Sub

    Private Sub showNoOfVerses()
        Try
            'SHOWS THE NUMBER OF VERSES IN THIS CHAPTER IN LABEL-VERSES
            Dim chapterText As String
            RichTextBoxDummy.LoadFile("Bible/" + NumericUpDownBook.Value.ToString + "/" + NumericUpDownChapter.Value.ToString + ".rtf")
            chapterText = RichTextBoxDummy.Text
            Dim count As Integer
            count = 0
            count = regexForNumberOfVerses(chapterText, "\b:\d?\S")

            LabelVerses.Text = count
            NumericUpDownVerse.Maximum = count
        Catch ex As Exception
            LabelStatus.Text = "Error in finding number of verses"
            errorLogger(ex)
        End Try
       
    End Sub

   
    Private Sub ButtonAbout_Click(sender As Object, e As EventArgs) Handles ButtonAbout.Click
        Try
            'ABOUT THIS APP INFO
            Dim text As String
            text = "Maximized Tamil Bible v3.0"
            text = text + vbNewLine + "Released on 02-October-2013"
            text = text + vbNewLine + "For queries or suggestions, visit http://creations-danny.yolasite.com or mail to sjoedanny@gmx.com."

            MsgBox(text, MsgBoxStyle.Information, AcceptButton)
        Catch ex As Exception
            LabelStatus.Text = "Error in showing About Info"
            errorLogger(ex)
        End Try

    End Sub

    Private Sub ButtonHelp_Click(sender As Object, e As EventArgs) Handles ButtonHelp.Click
        Try
            'HELP INFO
            Dim text As String
            text = "This software works with Windows Only. Latest version of .NET Framework is required."
            text = text + vbNewLine + "If there is a font problem, install Tamil Bible font in your system and try again."
            text = text + vbNewLine + "Tamil Bible font is provided with this package or you can find it online."
            text = text + vbNewLine + vbNewLine + "For queries or suggestions, visit http://creations-danny.yolasite.com or mail to sjoedanny@gmx.com."

            MsgBox(text, MsgBoxStyle.Information, AcceptButton)
        Catch ex As Exception
            LabelStatus.Text = "Error in showing Help Info"
            errorLogger(ex)
        End Try
    End Sub

    Private Sub ButtonFind_Click(sender As Object, e As EventArgs) Handles ButtonFind.Click
        Try
            'FIND FUNCTION
            Dim chapterText As String
            chapterText = RichTextBoxMiniDisplay.Text

            Dim currentVerse As Integer
            currentVerse = 0

            Dim textToFind As String
            textToFind = TextBoxFind.Text

            'PRESUMING EVERY VERSE ENDS WITH DOT AND SPLITTING THE CHAPTER THAT WAY (THE BEST POSSIBLE APPROACH AS OF NOW)
            Dim tokenizedVerses As Array
            tokenizedVerses = chapterText.Split(".")

            Dim foundVerses As String
            foundVerses = ""

            For i = 0 To tokenizedVerses.Length - 1
                'IF THE STRING IS FOUND, WE STORE IT FOR LATER PRINT
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
End Class
