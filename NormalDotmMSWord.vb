' The contents of my "normal.dotm" (Word macros)
' Mostly useful for editing, writing, or translating
' Many originally written by Paul Beverly and accessed from his site:
' http://www.archivepub.co.uk/macros.html
' an unparalleled source for quality editing macros. 


Sub ChangeDoubleStraightQuotes()
'Update 20131107
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = """"
.Replacement.Text = """"
.Forward = True
.Wrap = wdFindContinue
.Format = False
.MatchCase = False
.MatchWholeWord = False
.MatchWildcards = False
.MatchSoundsLike = False
.MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
End Sub



Sub AutoExec()
'
' AutoExec.AutoExec Macro
'
'

End Sub



Sub MakeSmallCaps()
     If Selection.Type = wdSelectionIP Then
          Selection.MoveLeft Unit:=wdWord, Count:=1
          Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
     End If
     If Selection.Type = wdSelectionNormal Then
 Selection.Font.Spacing = 2
Else
 MsgBox "You need to select some text."
End If
     Selection.Range.Case = wdLowerCase
     Selection.Font.SmallCaps = True
End Sub



Sub ShowOldSpellcheckDialog()

If Dialogs(wdDialogToolsSpellingAndGrammar).Show = -1 Then

MsgBox "Spelling and grammar check complete!"

End If

End Sub


Sub AAnAlyse()
' By Paul Beverly
' Source: http://www.archivepub.co.uk/macros.html

' Version 30.12.15
' Check a's and an's for agreement with following word

OKwithA = ",europe,european,once,one,uniform,uniformly,unified"
OKwithA = OKwithA & ",unique,uniquely,unit,unitarian,united,"
OKwithA = OKwithA & ",university,union,united,universe,"
OKwithA = OKwithA & ",universal,universally,unilateral,unilaterally,"
OKwithA = OKwithA & ",useful,usefully,useless,uselessly,user,"
OKwithA = OKwithA & ",usual,usually,,utility,utilities,utilitarian,"
OKwithA = OKwithA & ",utilization,utilisation,"

OKwithAn = ",hour,hourly,honest,honestly,honor,honour,honorary,"
OKwithAn = OKwithAn & ",honorarium,honorific,"

strongColour = wdBrightGreen
mutedColour = wdGray25

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "<[anA]{1,2}>"
  .Wrap = False
  .Replacement.Text = ""
  .Forward = True
  .MatchWildcards = True
  .MatchWholeWord = False
  .MatchSoundsLike = False
  .Execute
End With

OKwithA = "," & OKwithA & ","
OKwithAn = "," & OKwithAn & ","
qts = "'""" & ChrW(8216) & ChrW(8220)

Do While rng.Find.Found = True
  endNow = rng.End
  startArticle = rng.Start
  article = LCase(rng)
  rng.Start = endNow + 1
  rng.End = endNow + 2
  aOK = True
  If Len(rng) > 0 Then
    nextCharacter = Chr(Asc(rng))
  Else
    nextCharacter = ""
  End If
  rng.Expand wdWord
  nextWord = rng

  ' Check for quotes before and after
  rng.MoveEndWhile cset:=ChrW(8217) & "' ", Count:=wdBackward
  If nextCharacter = "'" Then
    rng.Start = rng.Start + 1
    nextWord = rng
    If Len(rng) > 0 Then
      nextCharacter = Chr(Asc(rng))
    Else
      nextCharacter = ""
    End If
  End If
  
  ' Check for apostrophe-s
  aposPosn = InStr(nextWord, ChrW(8217)) + InStr(nextWord, "'")
  If aposPosn > 0 Then rng.End = rng.Start + aposPosn - 1
  nextWord = rng
  
  ' Check for close quotes
  If InStr(qts, nextWord) > 0 Then
    rng.Collapse wdCollapseEnd
    rng.Expand wdWord
    rng.MoveEndWhile cset:=ChrW(8217) & "' ", Count:=wdBackward
    nextWord = rng
    If Len(rng) > 0 Then
      nextCharacter = Chr(Asc(rng))
    Else
      nextCharacter = ""
    End If
  End If
  endWord = rng.End
  
  ' Main check for agreement
  If LCase(nextWord) <> UCase(nextWord) Then
    If article = "a" Then
      If InStr("aAeEiIoOuU", nextCharacter) > 0 Then aOK = False
    End If
    
    If article = "an" Then
      If InStr("aAeEiIoOuU", nextCharacter) = 0 Then aOK = False
    End If
    
    If Len(nextWord) = 1 And article = "a" Then
      If InStr("FfHhLlMmNnRrSsXx", nextCharacter) > 0 Then
        aOK = False
      Else
        aOK = True
      End If
    End If
    
    ' Check single-letter words
    If Len(nextWord) = 1 And article = "an" Then
      If InStr("uU", nextCharacter) > 0 Then
        aOK = False
      Else
        aOK = True
      End If
    End If
    
    rng.Start = startArticle
    rng.End = endWord

    ' Check words from lists above that are exceptions
    testWord = "," & LCase(nextWord) & ","
    If InStr(OKwithA, testWord) > 0 Then
      If article = "a" Then aOK = True Else aOK = False
    End If
  
    If InStr(OKwithAn, testWord) > 0 Then
      If article = "an" Then aOK = True Else aOK = False
    End If

    ' Ignore people with the initial 'A.'
    If InStr(rng, ".") > 0 Then
      aOK = True
      nextWord = "xxx"
    End If

    ' Now highlight definite error
    If aOK = False Then
      rng.HighlightColorIndex = strongColour
      rng.Select
    End If
  
    ' Reduce highlight strength for acronyms
    ' that might not be wrong
    If UCase(nextWord) = nextWord Then
      If InStr("FHLMRSX", nextCharacter) > 0 Then
        If LCase(article) = "an" Then
          rng.HighlightColorIndex = mutedColour
        Else
          rng.HighlightColorIndex = strongColour
        End If
      End If
      If InStr("U", nextCharacter) > 0 Then
        If LCase(article) = "an" Then
          rng.HighlightColorIndex = strongColour
        Else
          rng.HighlightColorIndex = mutedColour
        End If
      End If
      Selection.Collapse wdCollapseEnd
    End If
  End If
    rng.Collapse wdCollapseEnd
  rng.Find.Execute
Loop
Beep
Selection.HomeKey Unit:=wdStory
End Sub


Sub AcronymToSmallCaps()

' By Paul Beverly
' Source: http://www.archivepub.co.uk/macros.html

' Version 31.12.13
' Convert all acronyms to small caps

myColour = False
' or if you want them highlighted:
' myColour = wdTurquoise

Selection.HomeKey Unit:=wdStory
With Selection.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "[A-Z]{3,}"
  .Wrap = False
  .Replacement.Text = ""
  .Forward = True
  .MatchWildcards = True
  .MatchWholeWord = False
  .MatchSoundsLike = False
  .Execute
End With

myCount = 0
Do While Selection.Find.Found = True
  Selection.Text = LCase(Selection.Text)
  If myColour <> False Then Selection.Range.HighlightColorIndex = myColour
  Selection.Font.SmallCaps = True
  Selection.Collapse wdCollapseEnd
  Selection.Find.Execute
Loop
End Sub


Sub ListAllLinks()

' By Paul Beverly
' Source: http://www.archivepub.co.uk/macros.html

' Version 01.09.16
' Creates a table of all URLs in the file

Set thisDoc = ActiveDocument
Documents.Add
Set listDoc = ActiveDocument
Set rng = ActiveDocument.Content
thisDoc.Activate

For Each lnk In ActiveDocument.Fields
  includeLink = True
  If lnk.Kind <> 2 Then
    lnk.Select
    includeLink = False
    ' MsgBox "Different link kind"
  End If
  
  If lnk.Type <> 88 Then
    lnk.Select
    includeLink = False
    ' MsgBox "Different link type"
  End If
  
  If includeLink = True Then
    linkCode = lnk.Code
    myURL = Replace(linkCode, "HYPERLINK", "")
    myURL = Trim(Replace(myURL, """", ""))
    myVisibleText = lnk.Result
  '  lnk.ShowCodes = True
    rng.InsertAfter Text:=myVisibleText & vbTab
    startLink = rng.End
    rng.InsertAfter Text:=myURL & vbCr
    rng.End = rng.End - 1
    rng.Start = startLink - 1
    rng.Font.Color = wdColorBlue
    rng.Font.Italic = True
    rng.Start = rng.End + 1
  End If
Next lnk

listDoc.Activate
Selection.WholeStory
Selection.ConvertToTable Separator:=wdSeparateByTabs
Selection.Tables(1).Style = "Table Grid"
Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
If deleteTableBorders = True Then
  Selection.Tables(1).Borders(wdBorderTop).LineStyle = wdLineStyleNone
  Selection.Tables(1).Borders(wdBorderLeft).LineStyle = wdLineStyleNone
  Selection.Tables(1).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
  Selection.Tables(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
  Selection.Tables(1).Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
  Selection.Tables(1).Borders(wdBorderVertical).LineStyle = wdLineStyleNone
End If
Selection.HomeKey Unit:=wdStory
End Sub


Sub PrintComments()

ActiveDocument.PrintOut Item:=wdPrintComments
End Sub


Sub SerialCommaHighlight()

' By Paul Beverly
' Source: http://www.archivepub.co.uk/macros.html

' Version 26.02.13
' Highlight or underline text that appears to have a serial comma

maxWords = 10
doUnderline = True
doHighlight = False
myColour = wdYellow

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "[a-zA-Z\-]@, [a-zA-Z\- ]@, and "
  .Replacement.Text = ""
  .MatchWildcards = True
  .MatchWholeWord = False
  .MatchSoundsLike = False
  .Execute
End With

While rng.Find.Found
  If rng.Words.Count < maxWords Then
    If doUnderline = True Then rng.Font.Underline = True
    If doHighlight = True Then rng.HighlightColorIndex = myColour
  End If
  rng.Start = rng.End
  rng.Find.Execute
Wend

Set rng = ActiveDocument.Content
With rng.Find
  .Text = "[a-zA-Z\-]@, [a-zA-Z\- ]@, or "
  .Replacement.Text = ""
  .Execute
End With

While rng.Find.Found
  If rng.Words.Count < maxWords Then
    If doUnderline = True Then rng.Font.Underline = True
    If doHighlight = True Then rng.HighlightColorIndex = myColour
  End If
  rng.Start = rng.End
  rng.Find.Execute
Wend
End Sub

Sub SerialNotCommaHighlight()

' By Paul Beverly
' Source: http://www.archivepub.co.uk/macros.html

' Version 26.02.13
' Highlight or underline text that appears not to have a serial comma

maxWords = 10
doUnderline = True
doHighlight = False
myColour = wdYellow

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "[a-zA-Z\-]@, [a-zA-Z\- ]@ and "
  .Replacement.Text = ""
  .MatchWildcards = True
  .MatchWholeWord = False
  .MatchSoundsLike = False
  .Execute
End With

While rng.Find.Found
  If rng.Words.Count < maxWords Then
    If doUnderline = True Then rng.Font.Underline = True
    If doHighlight = True Then rng.HighlightColorIndex = myColour
  End If
  rng.Start = rng.End
  rng.Find.Execute
Wend

Set rng = ActiveDocument.Content
With rng.Find
  .Text = "[a-zA-Z\-]@, [a-zA-Z\- ]@ or "
  .Replacement.Text = ""
  .Execute
End With

While rng.Find.Found
  If rng.Words.Count < maxWords Then
    If doUnderline = True Then rng.Font.Underline = True
    If doHighlight = True Then rng.HighlightColorIndex = myColour
  End If
  rng.Start = rng.End
  rng.Find.Execute
Wend
End Sub


Sub SpellingErrorLister()

' By Paul Beverly
' Source: http://www.archivepub.co.uk/macros.html


' Version 05.09.16
' List all spelling errors

myFind = "´a,´e,¨a,¨e,¨o,¨u,ˆo,"
myReplace = "á,é,ä,ë,ö,ü,ô,"

spellingListName = "Spelling Errors"

CR = vbCr
CR2 = CR & CR

' List possible spelling errors
errorLister:
thisLanguage = Selection.LanguageID
langName = Languages(thisLanguage).NameLocal
myLang = "Unknown language. OK?"
If thisLanguage = wdEnglishUK Then myLang = "UK spelling. OK?"
If thisLanguage = wdEnglishUS Then myLang = "US spelling. OK?"
myResponse = MsgBox(myLang, vbQuestion + vbYesNoCancel, "Spelling Error Highlighter")
If myResponse <> vbYes Then Exit Sub

timeStart = Timer

' Change ligature characters into character pairs
myFind = myFind & "," & ChrW(-1280) & "," & ChrW(-1279) & _
     "," & ChrW(-1278) & "," & ChrW(-1277) & "," _
     & ChrW(-1276)
myReplace = myReplace & ",ff,fi,fl,ffi,ffl"
fnd = Split(myFind, ",")
rpl = Split(myReplace, ",")
For i = 0 To UBound(fnd)
  Set rng = ActiveDocument.Content
  With rng.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = fnd(i)
    .Wrap = wdFindContinue
    .Replacement.Text = rpl(i)
    .MatchCase = False
    .MatchWildcards = False
    .Execute Replace:=wdReplaceAll
  End With
  If ActiveDocument.Footnotes.Count > 0 Then
    Set rng = ActiveDocument.StoryRanges(wdFootnotesStory)
    With rng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = fnd(i)
      .Wrap = wdFindContinue
      .Replacement.Text = rpl(i)
      .MatchCase = False
      .MatchWildcards = False
      .Execute Replace:=wdReplaceAll
    End With
  End If
  If ActiveDocument.Endnotes.Count > 0 Then
    Set rng = ActiveDocument.StoryRanges(wdEndnotesStory)
    With rng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = fnd(i)
      .Wrap = wdFindContinue
      .Replacement.Text = rpl(i)
      .MatchCase = False
      .MatchWildcards = False
      .Execute Replace:=wdReplaceAll
    End With
  End If
Next i

' Create spelling error list
erList1 = CR
erList2 = CR
numFootnotes = ActiveDocument.Footnotes.Count
numEndnotes = ActiveDocument.Endnotes.Count

myEnd = ActiveDocument.Content.End
For i = 1 To 3
  If myResponse = vbNo Then i = 3
  If i = 1 And numFootnotes = 0 Then i = 2
  If i = 1 Then Set rng = ActiveDocument.StoryRanges(wdFootnotesStory)
  If i = 2 And numEndnotes = 0 Then i = 3
  If i = 2 Then Set rng = ActiveDocument.StoryRanges(wdEndnotesStory)
  If i = 3 Then Set rng = ActiveDocument.Content
  k = 0
  For Each wd In rng.Words
  k = k + 1
    If Len(wd) > 2 And LCase(wd) <> UCase(wd) And _
         wd.Font.StrikeThrough = False Then
         padWd = " " & Trim(wd) & " "
      If Application.CheckSpelling(wd, MainDictionary:=langName) = False Then
        pCent = Int((myEnd - wd.End) / myEnd * 100)

        ' Report progress
        If i = 1 Then myPrompt = "Checking footnote text."
        If i = 2 Then myPrompt = "Checking endnote text."
        If i = 3 Then myPrompt = "Checking main text."
        StatusBar = "Generating errors list. " & myPrompt & _
             " To go:  " & Trim(Str(pCent)) & "%"
        DoEvents
        erWord = Trim(wd)
        lastChar = Right(erWord, 1)
        If lastChar = "'" Or lastChar = ChrW(8217) Then
          erWord = Left(erWord, Len(erWord) - 1)
        End If
        cap = Left(erWord, 1)
        If UCase(cap) = cap Then
          If InStr(erList1, CR & erWord & CR) = 0 Then erList1 = erList1 _
               & erWord & CR
        Else
          If InStr(erList2, CR & erWord & CR) = 0 Then erList2 = erList2 _
               & erWord & CR
        End If
      End If
    End If
  Next wd
Next i
Beep
mainFileName = ActiveDocument.Name
Documents.Add
Selection.TypeText Text:=erList2
Selection.WholeStory
Selection.Sort SortOrder:=wdSortOrderAscending


If erList1 <> CR Then
  Selection.EndKey Unit:=wdStory
  Selection.TypeText Text:=CR2
  listStart = Selection.Start
  Selection.TypeText Text:=erList1
  Selection.Start = listStart
  Selection.Sort SortOrder:=wdSortOrderAscending
End If

Selection.WholeStory
Selection.LanguageID = thisLanguage
Selection.Style = wdStyleNormal

Selection.Collapse wdCollapseStart
Selection.TypeText Text:=ChrW(124) & " " & spellingListName & CR

If numFootnotes > 0 Then
  Selection.TypeText Text:=CR & "| footnotes = yes" & CR
End If
If numEndnotes > 0 Then
  Selection.TypeText Text:=CR & "| endnotes = yes" & CR
End If

StatusBar = ""
Selection.WholeStory
Selection.Copy
ActiveDocument.Close SaveChanges:=False
Selection.EndKey Unit:=wdStory
myStart = Selection.Start
Selection.Paste
Selection.End = myStart

totTime = Int(10 * (Timer - timeStart) / 60) / 10
If totTime > 2 Then myResponse = MsgBox((totTime & "  minutes"), _
vbOKOnly, "Spelling Error Lister")
End Sub


Sub SpellingErrorHighlighter()

' By Paul Beverly
' Source: http://www.archivepub.co.uk/macros.html

' Version 27.01.17
' Highlight all spelling errors

spellingListName = ChrW(124) & " Spelling Errors"

CR = vbCr
CR2 = CR & CR

' Find errors list
Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = spellingListName
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .Forward = True
  .MatchCase = False
  .MatchWildcards = False
  .Execute
End With
If rng.Find.Found = False Then
  Beep
  MsgBox ("Can't find the list!")
  Exit Sub
End If
rng.End = ActiveDocument.Content.End

noList = (rng.HighlightColorIndex = 0)
sp = "                                      "
rng.Collapse wdCollapseStart
rng.Expand wdParagraph
rng.Select
If noList Then
  Beep
  myResponse = MsgBox("Please highlight at least one word in the list!", vbQuestion _
          + vbOK, "SpellingErrorHighlighter")
  Exit Sub
End If

' Create list of words needing highlighting in each colour
Dim myWordLists(16) As String
theEnd = ActiveDocument.Content.End
Set rng = Selection.Range.Duplicate

myCount = 0
Do
  rng.Collapse wdCollapseEnd
  rng.Expand wdParagraph
  If Len(rng) > 2 Then
    thisWord = Replace(rng.Text, CR, "")
    myCol = rng.HighlightColorIndex
    If myCol > 0 And myCol < 17 Then
      myWordLists(myCol) = myWordLists(myCol) & thisWord & ","
      myCount = myCount + 1
      StatusBar = sp & sp & sp & myCount
    End If
  End If
Loop Until rng.End = theEnd

totCount = myCount

fnNum = ActiveDocument.Footnotes.Count
enNum = ActiveDocument.Endnotes.Count
ActiveDocument.TrackRevisions = False

' To speed up search
Selection.HomeKey Unit:=wdStory

' For each highlight colour
oldColour = Options.DefaultHighlightColorIndex
For myCol = 1 To 16
  If Len(myWordLists(myCol)) > 0 Then
    Options.DefaultHighlightColorIndex = myCol
    myWds = Split(myWordLists(myCol), ",")
    For wd = 0 To UBound(myWds) - 1
      fWord = myWds(wd)
      For j = 1 To 3
        If j = 1 And fnNum = 0 Then j = 2
        If j = 2 And enNum = 0 Then j = 3
        Select Case j
          Case 1: Set rng = ActiveDocument.StoryRanges(wdFootnotesStory)
          Case 2: Set rng = ActiveDocument.StoryRanges(wdEndnotesStory)
          Case 3: Set rng = ActiveDocument.Content
        End Select
        DoEvents
        With rng.Find
          .ClearFormatting
          .Replacement.ClearFormatting
          .Text = "<" & fWord & ">"
          .Replacement.Text = "^&"
          .Font.StrikeThrough = False
          .Forward = True
          .Replacement.Highlight = True
          .MatchWildcards = True
          .Execute Replace:=wdReplaceAll
        End With
        If j = 3 Then
          myCount = myCount - 1
          StatusBar = sp & sp & sp & "To go: " & myCount
        End If
      Next j
    Next wd
  End If
Next myCol

Options.DefaultHighlightColorIndex = oldColour
Beep
StatusBar = " "
myResponse = MsgBox("All those errors have been highlighted", _
     vbOKOnly, "Spelling Error Highlighter")
End Sub


Sub DuplicatedWordsHighlight()

' By Paul Beverly
' Source: http://www.archivepub.co.uk/macros.html


' Version 03.02.16
' Add a highlight to any duplicate words in a text, e.g. "the the"

myColour1 = wdGray25
myColour2 = wdBrightGreen
myColour3 = wdYellow

doThreeWords = True

find1 = "(<[a-zA-Z]{1,})[ .,\!\?:;]{1,}\1[ .,\!\?:;]"
find2 = "(<[a-zA-Z]{1,}^32[a-zA-Z]{1,})[ .,\!\?:;]{1,}\1[ .,\!\?:;]"
find3 = "(<[a-zA-Z]{1,}^32[a-zA-Z]{1,}^32[a-zA-Z]{1,})" _
      & "[ .,\!\?:;]{1,}\1[ .,\!\?:;]"

oldColour = Options.DefaultHighlightColorIndex
Options.DefaultHighlightColorIndex = myColour1

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = find1
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .Replacement.Highlight = True
  .Forward = True
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

Options.DefaultHighlightColorIndex = myColour2
With rng.Find
  .Text = find2
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

If doThreeWords = True Then
  Options.DefaultHighlightColorIndex = myColour3
  With rng.Find
    .Text = find3
    .Wrap = wdFindContinue
    .Replacement.Text = ""
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
  End With
End If
Options.DefaultHighlightColorIndex = oldColour
End Sub



Sub LanguageSetCanadian()

' By Paul Beverly
' Source: http://www.archivepub.co.uk/macros.html

' Version 20.01.17
' Set language as Canadian English (mod. May 2017)

myLanguage = wdEnglishCanadian

myTrack = ActiveDocument.TrackRevisions
ActiveDocument.TrackRevisions = False
ActiveDocument.Content.LanguageID = myLanguage
If ActiveDocument.Footnotes.Count > 0 Then
  ActiveDocument.StoryRanges(wdFootnotesStory).LanguageID = myLanguage
End If
If ActiveDocument.Endnotes.Count > 0 Then
  ActiveDocument.StoryRanges(wdEndnotesStory).LanguageID = myLanguage
End If
If ActiveDocument.Shapes.Count > 0 Then
  For Each shp In ActiveDocument.Shapes
    If shp.Type <> 24 Then
      If shp.TextFrame.HasText Then
        shp.TextFrame.TextRange.LanguageID = myLanguage
      End If
    End If
  Next
End If
Application.CheckLanguage = True
ActiveDocument.Styles(wdStyleNormal).LanguageID = myLanguage
ActiveDocument.Styles(wdStyleCommentText).LanguageID = myLanguage
ActiveDocument.TrackRevisions = myTrack
End Sub
Sub LanguageSetUS()

' By Paul Beverly
' Source: http://www.archivepub.co.uk/macros.html

' Version 20.01.17
' Set language as US English

myLanguage = wdEnglishUS

myTrack = ActiveDocument.TrackRevisions
ActiveDocument.TrackRevisions = False
ActiveDocument.Content.LanguageID = myLanguage
If ActiveDocument.Footnotes.Count > 0 Then
  ActiveDocument.StoryRanges(wdFootnotesStory).LanguageID = myLanguage
End If
If ActiveDocument.Endnotes.Count > 0 Then
  ActiveDocument.StoryRanges(wdEndnotesStory).LanguageID = myLanguage
End If
If ActiveDocument.Shapes.Count > 0 Then
  For Each shp In ActiveDocument.Shapes
    If shp.Type <> 24 Then
      If shp.TextFrame.HasText Then
        shp.TextFrame.TextRange.LanguageID = myLanguage
      End If
    End If
  Next
End If
Application.CheckLanguage = True
ActiveDocument.Styles(wdStyleNormal).LanguageID = myLanguage
ActiveDocument.Styles(wdStyleCommentText).LanguageID = myLanguage
ActiveDocument.TrackRevisions = myTrack
End Sub

Sub LanguageSetFrenchCanadian()

' By Paul Beverly
' Source: http://www.archivepub.co.uk/macros.html

' Version 20.01.17
' Set language as Canadian French (mod. May 2017)

myLanguage = wdFrenchCanadian

myTrack = ActiveDocument.TrackRevisions
ActiveDocument.TrackRevisions = False
ActiveDocument.Content.LanguageID = myLanguage
If ActiveDocument.Footnotes.Count > 0 Then
  ActiveDocument.StoryRanges(wdFootnotesStory).LanguageID = myLanguage
End If
If ActiveDocument.Endnotes.Count > 0 Then
  ActiveDocument.StoryRanges(wdEndnotesStory).LanguageID = myLanguage
End If
If ActiveDocument.Shapes.Count > 0 Then
  For Each shp In ActiveDocument.Shapes
    If shp.Type <> 24 Then
      If shp.TextFrame.HasText Then
        shp.TextFrame.TextRange.LanguageID = myLanguage
      End If
    End If
  Next
End If
Application.CheckLanguage = True
ActiveDocument.Styles(wdStyleNormal).LanguageID = myLanguage
ActiveDocument.Styles(wdStyleCommentText).LanguageID = myLanguage
ActiveDocument.TrackRevisions = myTrack
End Sub



Sub exportcomments()
Dim s As String
Dim cmt As Word.Comment
Dim doc As Word.Document
For Each cmt In ActiveDocument.Comments
s = s & cmt.Initial & cmt.Index & "," & cmt.Range.Text & vbCr
Next
Set doc = Documents.Add
doc.Range.Text = s
End Sub

Sub CustomKeys()

' By Paul Beverly
' Source: http://www.archivepub.co.uk/macros.html

' Version 03.02.10

' Open the Customize Keyboard dialogue box

   With Dialogs(wdDialogToolsCustomizeKeyboard)

     .Category = 2

     .Show

   End With

End Sub

Sub countWordsHighlightedYellow()

Dim highlightCount
highlightCount = 0

For Each w In ActiveDocument.Words
    If w.HighlightColorIndex = wdYellow Then
        'w.Delete
        highlightCount = highlightCount + 1
    End If
Next

MsgBox ("There are " & highlightCount & " words highlighted yellow.")

End Sub

Sub DictionaryFetch()
' Version 11.11.13
' Launch selected text to dictionary.com

useExplorer = False

runBrowser = "C:\Program Files (x86)\Mozilla Firefox\Firefox"

mySite = "http://dictionary.com"

If Len(Selection) < 3 Then Selection.Words(1).Select
Selection.MoveEndWhile cset:=" " & ChrW(8217), Count:=wdBackward
Selection.MoveStartWhile cset:=" ", Count:=wdForward

mySubject = Replace(Selection, " ", "+")

getThis = mySite & "/browse/" & mySubject

If useExplorer = True Then
  Set objIE = CreateObject("InternetExplorer.application")
  objIE.Visible = True
  objIE.navigate mySite & getThis
  Set objIE = Nothing
Else
  Shell (runBrowser & " " & getThis)
End If
End Sub

Sub GoogleFetchQuotes()

' By Paul Beverly (modified)
' Source: http://www.archivepub.co.uk/macros.html


'Sub GoogleFetchQuotes()
' Version 15.03.17
' Launches selected text - with quotes - on Google

runBrowser = "C:\Program Files (x86)\Mozilla Firefox\Firefox"

mySite = "http://www.google.ca/search?q="

If Len(Selection) = 1 Then Selection.Expand wdWord
mySubject = Trim(Selection)
mySubject = Replace(mySubject, vbCr, "")
mySubject = Replace(mySubject, " ", "+")
mySubject = Replace(mySubject, "&", "%26")
mySubject = "%22" & mySubject & "%22"
Debug.Print mySubject
If useExplorer = True Then
  Set objIE = CreateObject("InternetExplorer.application")
  objIE.Visible = True
  objIE.navigate mySite & mySubject
  Set objIE = Nothing
Else
  Shell (runBrowser & " " & mySite & mySubject)
End If
End Sub

Sub ThesaurusFetch()
' Version 11.11.13
' Launch selected text to thesaurus.com

useExplorer = False

runBrowser = "C:\Program Files (x86)\Mozilla Firefox\Firefox"

mySite = "http://thesaurus.com"

If Len(Selection) < 3 Then Selection.Words(1).Select
Selection.MoveEndWhile cset:=" " & ChrW(8217), Count:=wdBackward
Selection.MoveStartWhile cset:=" ", Count:=wdForward

mySubject = Replace(Selection, " ", "+")

getThis = mySite & "/browse/" & mySubject

If useExplorer = True Then
  Set objIE = CreateObject("InternetExplorer.application")
  objIE.Visible = True
  objIE.navigate mySite & getThis
  Set objIE = Nothing
Else
  Shell (runBrowser & " " & getThis)
End If
End Sub

Sub JumpScroll()

' Source: http://www.archivepub.co.uk/macros.html

' Version 20.02.14
' Scroll this line to top of page
' Provided by Howard Silcock of New Zealand

Set wasSelected = Selection.Range
Application.ScreenUpdating = False

Selection.EndKey Unit:=wdStory
wasSelected.Select
Application.ScreenUpdating = True

ActiveDocument.ActiveWindow.SmallScroll down:=1
End Sub




Sub ScrollDown()
ActiveDocument.ActiveWindow.SmallScroll down:=10
End Sub


Sub LingueeFetch()

' By Paul Beverly (modified)
' Source: http://www.archivepub.co.uk/macros.html

' Version 11.11.13 adapted PS
' Launch selected text to linguee.fr

useExplorer = False

runBrowser = "C:\Program Files (x86)\Mozilla Firefox\Firefox"

mySite = "http://www.linguee.fr"

If Len(Selection) < 3 Then Selection.Words(1).Select
Selection.MoveEndWhile cset:=" " & ChrW(8217), Count:=wdBackward
Selection.MoveStartWhile cset:=" ", Count:=wdForward

mySubject = Replace(Selection, " ", "+")

getThis = mySite & "/francais-anglais/" & "search?source=auto&query=" & mySubject

If useExplorer = True Then
  Set objIE = CreateObject("InternetExplorer.application")
  objIE.Visible = True
  objIE.navigate mySite & getThis
  Set objIE = Nothing
Else
  Shell (runBrowser & " " & getThis)
End If
End Sub

