Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Public Class FindText

    Private rt As RichTextBox = Nothing

    Private Sub ButtonFindNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonFindNext.Click
        rt = GetActiveRichTextBox()
        If (rt Is Nothing) Then
            Me.Close()
            Exit Sub
        End If

        If CheckBoxRegEx.Checked Then
            FindRegEx(False)
        Else
            FindString(False)
        End If
    End Sub

    Private Sub ButtonFindPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonFindPrevious.Click
        rt = GetActiveRichTextBox()
        If (rt Is Nothing) Then
            Me.Close()
            Exit Sub
        End If

        If CheckBoxRegEx.Checked Then
            FindRegEx(True)
        Else
            FindString(True)
        End If
    End Sub

    Private Sub FindRegEx(ByVal boolReverse As Boolean)

        ClearSelection()

        Dim boolFound As Boolean = False

        Dim strPattern As String = ComboBoxSearchString.Text

        If strPattern = "" Then Exit Sub

        If Not ComboBoxSearchString.Items.Contains(strPattern) Then
            FormRemotePCAdmin.ListRecentSearches.Insert(0, strPattern)
            ComboBoxSearchString.Items.Insert(0, strPattern)
        End If

        Dim options As RegexOptions = RegexOptions.IgnoreCase
        If CheckBoxMatchCase.Checked Then
            options = RegexOptions.None
        End If

        Dim colMatches As MatchCollection = Regex.Matches(rt.Text, strPattern, options)

        If colMatches.Count = 0 Then
            MsgBox("The specified text was not found.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim intStep As Integer
        Dim intStart As Integer
        Dim intEnd As Integer
        Dim intSelectionStart As Integer

        If boolReverse Then
            intStep = -1
            intStart = colMatches.Count - 1
            intEnd = 0
            intSelectionStart = rt.SelectionStart
        Else
            intStep = 1
            intStart = 0
            intEnd = colMatches.Count - 1
            intSelectionStart = rt.SelectionStart + rt.SelectedText.Length
        End If

        boolFound = SelectMatch(colMatches, intStart, intEnd, intStep, intSelectionStart, boolReverse)

        'If not found and reached beginning or end, loop
        Dim msg As String = ""
        If boolFound = False Then
            If boolReverse Then
                'Were we at the end when we started?
                If intSelectionStart < rt.TextLength Then
                    msg = "Reached the beginning of the text, restarting from the end."
                    intSelectionStart = rt.TextLength
                End If
            Else
                'Were we at the beginning when we started?
                If intSelectionStart > 0 Then
                    msg = "Reached the end of the text, restarting from the beginning."
                    intSelectionStart = 0
                End If
            End If
            If msg <> "" AndAlso MsgBox(msg, MsgBoxStyle.Exclamation + MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                boolFound = SelectMatch(colMatches, intStart, intEnd, intStep, intSelectionStart, boolReverse)
            End If

        End If



        If boolFound Then rt.ScrollToCaret()


    End Sub

    Private Function SelectMatch(ByVal colMatches As MatchCollection, ByVal intStart As Integer, _
                                ByVal intEnd As Integer, ByVal intStep As Integer, ByVal intSelectionStart As Integer, _
                                ByVal boolReverse As Boolean) As Boolean
        For n As Integer = intStart To intEnd Step intStep
            If boolReverse Then
                If colMatches(n).Index < intSelectionStart Then
                    rt.Select(colMatches(n).Index, colMatches(n).Length)
                    rt.SelectionBackColor = Color.LightGray
                    Return True
                End If
            Else
                If colMatches(n).Index > intSelectionStart Then
                    rt.Select(colMatches(n).Index, colMatches(n).Length)
                    rt.SelectionBackColor = Color.LightGray
                    Return True
                End If
            End If
        Next
        Return False
    End Function

    Private Sub FindString(ByVal boolReverse As Boolean)

        ClearSelection()

        Dim boolFound As Boolean = False
        Dim intStart As Integer

        Dim strFindString As String = ComboBoxSearchString.Text

        If strFindString = "" Then Exit Sub

        If Not ComboBoxSearchString.Items.Contains(strFindString) Then
            FormRemotePCAdmin.ListRecentSearches.Insert(0, strFindString)
            ComboBoxSearchString.Items.Insert(0, strFindString)
        End If

        Dim options As RichTextBoxFinds = RichTextBoxFinds.None
        If CheckBoxMatchCase.Checked Then
            options = RichTextBoxFinds.MatchCase
        End If

        Dim x As Integer
        If boolReverse Then
            options += RichTextBoxFinds.Reverse
            intStart = rt.SelectionStart
            x = rt.Find(strFindString, 0, intStart, options)
        Else
            intStart = rt.SelectionStart + rt.SelectedText.Length
            x = rt.Find(strFindString, intStart, options)
        End If

        If x > 0 Then boolFound = True

        If boolFound = False Then
            Dim msg As String
            If boolReverse Then
                msg = "Reached the beginning of the text, restarting from the end."
                intStart = rt.TextLength
            Else
                msg = "Reached the end of the text, restarting from the beginning."
                intStart = 0
            End If
            If MsgBox(msg, MsgBoxStyle.Exclamation + MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                x = rt.Find(strFindString, intStart, options)
                If x > 0 Then boolFound = True
            End If
        End If

        If boolFound = True Then
            rt.Select(x, strFindString.Length)
            rt.SelectionBackColor = Color.LightGray
            rt.ScrollToCaret()
        Else
            MsgBox("The specified text was not found.", MsgBoxStyle.Exclamation)
        End If
    End Sub


    Private Sub ClearSelection()
        rt.SelectionBackColor = rt.BackColor
        rt.SelectedText = ""
    End Sub

    Private Sub ButtonClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClose.Click
        Me.Close()
    End Sub

    Private Sub FindText_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not rt Is Nothing Then
            rt.SelectionBackColor = rt.BackColor
            rt.Focus()
        End If
    End Sub

    Private Sub FindText_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        rt = GetActiveRichTextBox
        If rt Is Nothing Then Me.Close()


        For Each strItem As String In FormRemotePCAdmin.ListRecentSearches
            ComboBoxSearchString.Items.Add(strItem)
        Next

        If ComboBoxSearchString.Items.Count > 0 Then ComboBoxSearchString.SelectedIndex = 0

        rt.Select(0, 0)

    End Sub

    Private Function GetActiveRichTextBox() As RichTextBox

        Dim cc As TabPage.TabPageControlCollection = FormRemotePCAdmin.TabControl1.SelectedTab.Controls

        Dim ListControls As New List(Of Control)

        'Find all controls on active tab
        For Each cont As Control In cc
            ListControls.Add(cont)
            If cont.Controls.Count Then
                For Each subcont As Control In cont.Controls
                    ListControls.Add(subcont)
                Next
            End If
        Next

        'Find the richtextbox
        For Each cont As Control In ListControls
            If TypeOf cont Is RichTextBox Then
                Return CType(cont, RichTextBox)
                Exit For
            End If
        Next

        Return Nothing

    End Function

End Class
