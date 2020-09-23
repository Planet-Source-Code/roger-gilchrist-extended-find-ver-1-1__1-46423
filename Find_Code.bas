Attribute VB_Name = "Find_Code"
Option Explicit
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
'This is a very slightly modified version of routines I have used in earlier versions
'Slight changes to workaround UserDocument limits
Public Enum EnumReportAction
  Search
  Complete
  inComplete
  Missing
  Found
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private Search, Complete, inComplete, Missing, Found
#End If
'Search Settings
Public Const LongLimit               As Single = 2147483647
'Prevent FlexGrid from over-flowing;
'very unlikely that anything could be found that often
'just a safety valve inherited from the old listbox/IntegerLimit
Public bWholeWordonly                As Boolean
Public bCaseSensitive                As Boolean
Public bPunctuationAware             As Boolean
Public bNoComments                   As Boolean
Public bNoStrings                    As Boolean
Public bStringsOnly                  As Boolean
Public bCurCodePane                  As Boolean
Public BPatternSearch                As Boolean
Public bPTmpWholeWordonly            As Boolean
Public bPTmpCaseSensitive            As Boolean
Public bPTmpPunctuationAware         As Boolean
Public bPTmpNoComments               As Boolean
Public bPTmpNoStrings                As Boolean
Public bPTmpStringsOnly              As Boolean
Public bPTmpFindSelectWholeLine      As Boolean
Public bGridlines                    As Boolean
Private bComplete                    As Boolean
'halt search before completion
Public bCancel                       As Boolean
Public Const Apostrophe              As String = "'"
Public bShowProject                  As Boolean
Public bShowComponent                As Boolean
Public bShowRoutine                  As Boolean
Public bFindSelectWholeLine          As Boolean

Private Sub ApplyStringCommentFilters(cde As String, _
                                      ByVal StrTarget As String)

  Dim Codepos As Long

  Codepos = InStr(cde, StrTarget)
  If bNoComments Then
    If InComment(cde, Codepos) Then
      cde = vbNullString
    End If
  End If
  If bStringsOnly Then
    If InQuotes(cde, Codepos) = False Then
      cde = vbNullString
    End If
  End If
  If bNoStrings Then
    If InQuotes(cde, Codepos) Then
      cde = vbNullString
    End If
  End If

End Sub

Private Function CancelSearch(G_Rows As Long) As Boolean

  CancelSearch = bCancel Or EscPressed Or G_Rows = LongLimit

End Function

Public Sub ClearFGrid(Fgrd As MSFlexGrid)

  Fgrd.Rows = 2

End Sub

Public Sub ComboBoxSavePreviousSearch(combo As ComboBox, _
                                      Optional AddWord As String = vbNullString, _
                                      Optional ListLimit As Long = 20)

  'add a new word to history
  'if already in history move to top of history

  On Error Resume Next
  If LenB(AddWord) Then
    combo.Text = AddWord
   Else
    AddWord = combo.Text
  End If
  With combo
    If LenB(.Text) Then
      If PosInCombo(.Text, combo) = -1 Then
        .AddItem .Text, 0
        If .ListCount > ListLimit Then
          .RemoveItem ListLimit
        End If
       Else
        .RemoveItem PosInCombo(.Text, combo)
        .AddItem AddWord, 0
        .Text = AddWord
      End If
    End If
  End With
  On Error GoTo 0

End Sub

Public Sub DOFind(Fgrd As MSFlexGrid, _
                  cmb As ComboBox, _
                  cmdF1 As CommandButton)

  Dim Proj           As VBProject
  Dim Comp           As VBComponent
  Dim Pane           As CodePane
  Dim tmpLindex      As Long
  Dim StartText      As Long
  Dim EndText        As Long
  Dim StrProjName    As String
  Dim StrCompName    As String
  Dim CompLineNo     As Long
  Dim StrTarget      As String
  Dim strRoutine     As String
  Dim strFound       As String
  Dim StartCol       As Long
  Dim EndLine        As Long
  Dim EndCol         As Long

  On Error Resume Next
  With Fgrd
    StrProjName = .TextMatrix(.Row, 0)
    StrCompName = .TextMatrix(.Row, 1)
    strRoutine = .TextMatrix(.Row, 2)
    strFound = .TextMatrix(.Row, 3)
    CompLineNo = .RowData(.Row)
  End With 'Fgrd
  StrTarget = cmb
  For Each Proj In VBInstance.VBProjects
    If StrProjName = Proj.Name Then
      For Each Comp In Proj.VBComponents
        If Comp.Name = StrCompName Then
          Set Pane = Comp.CodeModule.CodePane
          With Pane
            'when docked only the first instance selected in GrdFound got highlighted
            'until I added next line, no idea why it works.
            .Window.Visible = False
            .Show
            .Window.SetFocus
            .TopLine = Abs(Int(.CountOfVisibleLines / 2) - CompLineNo) + 1
          End With
          If bFindSelectWholeLine Then
            'select the whole line
            StartText = InStr(1, Comp.CodeModule.Lines(CompLineNo, 1), strFound, vbTextCompare)
            EndText = StartText + Len(strFound)
           Else
            'select the search word
            If Not BPatternSearch Then
              StartText = FilteredInStr(Comp.CodeModule.Lines(CompLineNo, 1), StrTarget)
              EndText = StartText + Len(StrTarget)
             Else
              ' the string/comment filters cannot work on a PatternSearch StrFind
              'but you can get the actual string found and test that
              StartCol = 1
              EndLine = -1
              EndCol = -1
              With Comp
                If .CodeModule.Find(StrTarget, CompLineNo, StartCol, EndLine, EndCol, bWholeWordonly, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch) Then
                  .CodePane.SetSelection CompLineNo, StartCol, EndLine, EndCol
                  'StrHardPattern = Mid$(Comp.CodeModule.Lines(CompLineNo, 1), StartCol, EndCol - StartCol)
                  StartText = StartCol 'FilteredInStr(Comp.CodeModule.Lines(CompLineNo, 1), StrHardPattern)
                  EndText = EndCol - StartCol 'StartText + Len(StrHardPattern)
                End If
              End With 'Comp
            End If
          End If
          If StartText > 0 Then
            Fgrd.TextMatrix(0, 3) = ReportAction(Fgrd, Found)
            With Pane
              .SetSelection CompLineNo, StartText, CompLineNo, EndText
            End With
            Set Pane = Nothing
            Exit Sub
           Else
            tmpLindex = Fgrd.Row
            DoSearch Fgrd, cmb, cmdF1
            If Fgrd.Rows > 1 Then
              Fgrd.Row = tmpLindex
              Pane.Window.SetFocus
              Pane.SetSelection CompLineNo, StartText, CompLineNo, EndText
              Set Pane = Nothing
            End If
          End If
        End If
      Next Comp
    End If
  Next Proj
  On Error GoTo 0

End Sub

Public Sub DoSearch(Fgrd As MSFlexGrid, _
                    cmb As ComboBox, _
                    cmdF1 As CommandButton)

  'ver1.1.02 major rewrite to allow simple pattern searching
  
  Dim code                     As String
  Dim ProcName                 As String
  Dim CodeLineNo               As Long
  Dim CompMod                  As CodeModule
  Dim Comp                     As VBComponent
  Dim Proj                     As VBProject
  Dim strFind                  As String
  Dim LongestPrj               As String
  Dim longestPrc               As String
  Dim LongestCmp               As String
  Dim ResizeNeeded             As Boolean
  Dim StrHardPattern           As String
  Dim StartCol                 As Long
  Dim EndLine                  As Long
  Dim EndCol                   As Long
  Dim SecondRun                As Boolean

  On Error Resume Next
  strFind = cmb.Text
  If LenB(strFind) = 0 Then
    Exit Sub
  End If
  If strFind = " " Then
    MsgBox "Search for single spaces is cancelled, it overloads the system", vbInformation
    SetFocus_Safe cmb
    Exit Sub
  End If
retry:
  If LenB(strFind) > 0 Then
    LongestPrj$ = vbNullString
    bComplete = False
    ClearFGrid Fgrd
    ComboBoxSavePreviousSearch cmb, , HistDeep
    cmdF1.Default = True
    cmdF1.Visible = True
    DoEvents
    bCancel = False
    GetCounts
    For Each Proj In VBInstance.VBProjects
      If Len(Proj.Name) > Len(LongestPrj$) Then
        LongestPrj$ = Proj.Name
        ResizeNeeded = True
      End If
      For Each Comp In Proj.VBComponents
        If Fgrd.Rows < LongLimit Then
          If SafeCompToProcess(Comp) Then
            If bCurCodePane Then
              If Comp.CodeModule.CodePane.Window.WindowState = vbext_ws_Normal Then
                GoTo SkipComp
              End If
            End If
            If Len(Comp.Name) > Len(LongestCmp$) Then
              LongestCmp$ = Comp.Name
              ResizeNeeded = True
            End If
            With Comp
              Set CompMod = .CodeModule
              '5hould I quit?
              Fgrd.TextMatrix(0, 3) = ReportAction(Fgrd, Search)
              If LenB(.Name) = 0 Then
                bCancel = True
                bComplete = True
              End If
            End With
            If CancelSearch(Fgrd.Rows) Then
              Exit For
            End If
            'Safety turns off filters if comment/double quote is actually in the search phrase
            If bNoComments Then
              If InStr(strFind, Apostrophe) > 0 Then
                bNoComments = True
                mobjDoc.SetFilterButtons
              End If
            End If
            If bNoStrings Then
              If InStr(strFind, Chr$(34)) > 0 Then
                bNoStrings = False
                mobjDoc.SetFilterButtons
              End If
            End If
            If CompMod.Find(strFind, 1, 1, CompMod.CountOfLines, 1, bWholeWordonly, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch) Then
              'if exits at all, then look for the line(s)
              CodeLineNo = 1
              Do
                DoEvents
                If CancelSearch(Fgrd.Rows) Then
                  Exit For
                End If
                StartCol = 1
                EndLine = -1
                EndCol = -1
                If CompMod.Find(strFind, CodeLineNo, StartCol, EndLine, EndCol, bWholeWordonly, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch) Then
                  ProcName = CompMod.ProcOfLine(CodeLineNo, vbext_pk_Proc)
                  If LenB(ProcName) = 0 Then
                    ProcName = "(Declarations)"
                  End If
                  If Len(ProcName) > Len(longestPrc$) Then
                    longestPrc$ = ProcName
                    ResizeNeeded = True
                  End If
                  code$ = CompMod.Lines(CodeLineNo, 1)
                  If Not BPatternSearch Then
                    ApplyStringCommentFilters code$, strFind
                   Else
                    ' the string/comment filters cannot work on a PatternSearch StrFind
                    'but you can get the actual string found and test that
                    CompMod.CodePane.SetSelection CodeLineNo, StartCol, EndLine, EndCol
                    StrHardPattern = Mid$(code$, StartCol, EndCol - StartCol)
                    ApplyStringCommentFilters code$, StrHardPattern
                  End If
                  If LenB(code$) Then
                    If ResizeNeeded Then
                      'slight speed advantage of not doing this unless called for
                      mobjDoc.GridReSize LongestPrj$, LongestCmp$, longestPrc$
                      ResizeNeeded = False
                    End If
                    With Fgrd
                      'ver 1.1.02
                      'this is it the bug that stuffed ver1.1.01
                      'Moved this line to before filling row
                      .Rows = .Rows + 1
                      .TextMatrix(.Row, 0) = Proj.Name
                      .TextMatrix(.Row, 1) = Comp.Name
                      .TextMatrix(.Row, 2) = ProcName
                      .TextMatrix(.Row, 3) = code$
                      'stores the line number for retrieval in DoFind
                      .RowData(.Row) = CodeLineNo
                      .Row = .Row + 1
                      code$ = vbNullString
                    End With 'Fgrd
                  End If
                End If
                CodeLineNo = CodeLineNo + 1
                If CancelSearch(Fgrd.Rows) Then
                  Exit Do
                End If
              Loop While CodeLineNo > 0 And CodeLineNo <= CompMod.CountOfLines
            End If
          End If
        End If
SkipComp:
        Set Comp = Nothing
        If CancelSearch(Fgrd.Rows) Then
          Exit For
        End If
      Next Comp
      If CancelSearch(Fgrd.Rows) Then
        Exit For
      End If
    Next Proj
    Set Proj = Nothing
    Set CompMod = Nothing
    With cmdF1
      .Visible = False
      .Default = False
    End With
  End If
  With Fgrd
    .Rows = .Rows - 1
    If .Row = 0 Then
      'nothing found
      .BackColorSel = .BackColorFixed
     Else
      .BackColorSel = &H8000000D
    End If
    'automatically switch to pattern search if ordinary fails
  End With 'Fgrd
  If Fgrd.Row = 0 Then
    If Not BPatternSearch Then
      If instrAny(strFind, "*", "!", "[", "]", "\") Then
        BPatternSearch = Not BPatternSearch
        mobjDoc.ClearForPattern
        SecondRun = True
        GoTo retry
      End If
    End If
    If bComplete = False Then
      Fgrd.TextMatrix(0, 3) = ReportAction(Fgrd, Found)
      SetFocus_Safe cmb
     Else
      Fgrd.TextMatrix(0, 3) = ReportAction(Fgrd, inComplete)
    End If
    'this turns auto switch to pattern search off if it was used
    If SecondRun Then
      BPatternSearch = Not BPatternSearch
      mobjDoc.ClearForPattern
    End If
  End If
  If Fgrd.Rows = LongLimit Then
    ' as this is 2147483647 rows it is unlikely that this will ever hit but just in case :)
    MsgBox "Search halted because number of finds reached limit of Find ComboBox", vbCritical
  End If
  On Error GoTo 0

End Sub

Private Function FilteredInStr(strSearch As String, _
                               strFind As String) As Long

  'ver 1.1.02
  'improved word detection makes sure that the DoFind routine
  'highlights the correct (or at least first instance)
  'of string that matches all filters
  
  Dim Lbit As String
  Dim Rbit As String

  Do
    FilteredInStr = InStr(FilteredInStr + 1, strSearch, strFind, IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare))
    If strSearch = strFind Then
      Exit Do
    End If
    If bWholeWordonly Then
      Select Case FilteredInStr
       Case 1
        Rbit$ = Mid$(strSearch, Len(strFind) + 1, 1)
        If Not bPunctuationAware Then
          If Rbit$ = " " Then
            Exit Do
          End If
         Else
          If IsPunct(Rbit) Then
            Exit Do
          End If
        End If
       Case Len(strSearch) - Len(strFind) + 1
        Lbit$ = Mid$(strSearch, Len(strSearch) - Len(strFind), 1)
        If Not bPunctuationAware Then
          If Lbit$ = " " Then
            Exit Do
          End If
         Else
          If IsPunct(Lbit) Then
            Exit Function
          End If
        End If
       Case 0
        Exit Do 'not found do nothing
       Case Else
        Lbit$ = Mid$(strSearch, FilteredInStr - 1, 1)
        Rbit$ = Mid$(strSearch, FilteredInStr + Len(strFind), 1)
        If Not bPunctuationAware Then
          If Lbit$ = " " Then
            If Rbit$ = " " Then
              Exit Do
            End If
          End If
         Else
          If IsPunct(Lbit) And IsPunct(Rbit) Then
            Exit Do
          End If
        End If
      End Select
     Else
      Exit Do
    End If
  Loop

End Function

Public Function InComment(ByVal code As String, _
                          ByVal Codepos As Long) As Boolean

  Dim SQuotePos As Long

  On Error Resume Next
  SQuotePos = InStr(code$, Apostrophe)
  If SQuotePos = 1 Then
    InComment = True
  End If
  If SQuotePos > 1 Then
    If Codepos > SQuotePos Then
      InComment = True
    End If
  End If
  On Error GoTo 0

End Function

Public Function InQuotes(ByVal code As String, _
                         ByVal Codepos As Long) As Boolean

  Dim LQ As Long
  Dim FQ As Long

  On Error Resume Next
  LQ = InStr(StrReverse(code$), Chr$(34))
  If LQ > 0 Then
    LQ = Len(code$) - LQ + 1
  End If
  FQ = InStr(code$, Chr$(34))
  If LQ = 0 Then
    If FQ = 0 Then
      Exit Function
    End If
  End If
  If LQ = FQ Then
    Exit Function
  End If
  If FQ < Codepos Then
    If Codepos < LQ Then
      InQuotes = True
    End If
  End If
  On Error GoTo 0

End Function

Public Function IsAlphaIntl(ByVal sChar As String) As Boolean

  IsAlphaIntl = Not (UCase$(sChar) = LCase$(sChar))

End Function

Public Function IsNumeral(ByVal strTest As String) As Boolean

  IsNumeral = InStr("1234567890", strTest) > 0

End Function

Public Function IsPunct(ByVal strTest As String) As Boolean

  'Detect punctuation

  If IsNumeral(strTest) Then
    IsPunct = False
   Else
    IsPunct = Not IsAlphaIntl(strTest)
  End If

End Function

Public Function ReportAction(Fgrd As MSFlexGrid, _
                             ByVal Act As EnumReportAction, _
                             Optional ByVal AppendStr As String) As String

  Dim StrItems           As String
  Dim StrFilterWarning   As String
  Dim strSearchEndStatus As String
  Dim StrPatternWarning  As String

  StrItems = "(" & Fgrd.Rows - 1 & ") Item" & IIf(Fgrd.Rows - 1 <> 1, "s", vbNullString)
  StrFilterWarning = IIf(mobjDoc.AnyFilterOn, " <Filter>", vbNullString)
  StrPatternWarning = IIf(BPatternSearch, " <Pattern>", vbNullString)
  Select Case Act
   Case Search
    strSearchEndStatus = " Searching " & IIf(Len(AppendStr), " in " & AppendStr, "...")
   Case Complete
    strSearchEndStatus = " Search Complete."
   Case inComplete
    strSearchEndStatus = " Search Cancelled."
  End Select
  Select Case Act
   Case Missing
    ReportAction = "Code: " & StrItems & "Item not found" & StrFilterWarning & strSearchEndStatus & StrPatternWarning
   Case Else
    ReportAction = "Code: " & StrItems & " found." & StrFilterWarning & strSearchEndStatus & StrPatternWarning
  End Select
  DoEvents

End Function

Public Sub SelectedText(cmb As ComboBox, _
                        Cmd As CommandButton)

  Dim HiLitSelection As String

  HiLitSelection$ = GetSelectedText(VBInstance)
  If LenB(HiLitSelection$) Then
    If InStr(HiLitSelection$, vbNewLine) Then
      If (HiLitSelection$ <> vbNewLine) Then
        HiLitSelection$ = Left$(HiLitSelection$, InStr(HiLitSelection$, vbNewLine))
      End If
    End If
    If LenB(HiLitSelection$) Then
      cmb.SetFocus
      cmb.Text = HiLitSelection$
      Cmd = True
    End If
  End If

End Sub

Private Function WholeWordTest(ByVal S As String, _
                               ByVal t As String) As Boolean

  Dim Lpos1 As Long
  Dim Rpos1 As Long

  On Error Resume Next
  If LenB(S) > 0 Then
    If LenB(t) > 0 Then
      If Not bCaseSensitive Then
        S = LCase$(S)
        t = LCase$(t)
      End If
      If S = t Then
        WholeWordTest = True
        Exit Function
      End If
      If InStr(S, t) = 0 Then
        Exit Function
      End If
      If InStr(S, " " & t & " ") Then
        WholeWordTest = True
        Exit Function
      End If
      If MultiLeft(S, True, t & " ") Then
        WholeWordTest = True
        Exit Function
      End If
      If MultiRight(S, True, " " & t) Then
        WholeWordTest = True
        Exit Function
      End If
      If bPunctuationAware Then
        Lpos1 = InStr(S, t) - 1
        Rpos1 = InStr(S, t) + Len(t)
        If Lpos1 = 0 Then
          If IsPunct(Mid$(S, Rpos1, 1)) Then
            WholeWordTest = True
            Exit Function
          End If
        End If
        If Rpos1 = Len(S) Then
          If Lpos1 > 0 Then
            If IsPunct(Mid$(S, Lpos1, 1)) Then
              If IsPunct(Mid$(S, Rpos1, 1)) Then
                WholeWordTest = True
                Exit Function
              End If
            End If
          End If
        End If
        If Lpos1 > 0 Then
          If Rpos1 <= Len(S) + 1 Then
            If IsPunct(Mid$(S, Rpos1, 1)) Then
              If IsPunct(Mid$(S, Lpos1, 1)) Then
                WholeWordTest = True
                Exit Function
              End If
            End If
          End If
        End If
      End If
    End If
  End If
  On Error GoTo 0

End Function

':) Roja's VB Code Fixer V1.1.2 (30/06/2003 3:53:46 PM) 43 + 605 = 648 Lines Thanks Ulli for inspiration and lots of code.

