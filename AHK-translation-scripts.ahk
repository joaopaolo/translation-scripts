

; -------------------------------------------------------------------
; SCRIPTS FOR TRANSLATION, EDITING, AND AUTOMATION
; Imperfect but working for me. Many borrowed here and there; thanks!
; -------------------------------------------------------------------

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance force


; -----------------------
; OPEN SPECIFIC DIRECTORY
; -----------------------

!+d::
Run, "C:\Users\"YOUR USERNAME"\Downloads"
return

!+m::
Run, "C:\Users\"YOUR USERNAME"\Documents"
return

; etc.

; -----------------------------------------------------------
; COPY TO TEXT
; Use hotkey to send selected text to .txt file for later use
; ------------------------------------------------------------

^#c::
clipboard := "" 
Send ^c
Clipwait
ToolTip % clipboard
sleep 3000
ToolTip
FileAppend %clipboard%'n, "C:\Users\"YOUR USERNAME"\Documents\RunningCopy.txt"
Return


; --------------------------------------------------
; SEND TEXT TO SPECIFIC WEBSITES
; Reminder: # = WinKey; ^ = Ctrl; + = Shift; ! = Alt
; -----------------------------------------------------

; ------------------------------------------------------
; ALL HOTKEYS
; #+l = linguee
; ^!c = open all links on clipboard in tabs
; #+t = termium
; #+w = wordreference (FR>EN)
; #+g = google
; #+; = thesaurus.com
;-------------------------------------------------------


; --------------------------------------------
; Search Google for selection, or open URL ###
; --------------------------------------------

; IMPORTANT: Set the correct path to your default broswer
browser="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

#+g::
;Copy Clipboard to prevClipboard variable, clear Clipboard.
prevClipboard := ClipboardAll
Clipboard =
;Copy current selection, continue if no errors.
SendInput, ^c
ClipWait, 2
if !(ErrorLevel) {
;Convert Clipboard to text, auto-trim leading and trailing spaces and tabs.
Clipboard = %Clipboard%
;Clean Clipboard: change carriage returns to spaces, change >=1 consecutive spaces to +
Clipboard := RegExReplace(RegExReplace(Clipboard, "\r?\n"," "), "\s+","+")
;Open URLs, Google non-URLs. URLs contain . but do not contain + or .. or @
if Clipboard contains +,..,@
Run, %browser% http://www.google.ca/search?q=%Clipboard%
else if Clipboard not contains .
Run, %browser% http://www.google.ca/search?q=%Clipboard%
else
Run, %browser% %Clipboard%
}
;Restore Clipboard, clear prevClipboard variable.
Clipboard := prevClipboard
prevClipboard =
return

; ------------------------------------------
; Selected text to wordreference.com (FR>EN)
; ------------------------------------------

; +l### Search Wordreference for selection, or open URL ###

; IMPORTANT: Set the correct path to your default broswer
browser="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

#+w::
;Copy Clipboard to prevClipboard variable, clear Clipboard.
prevClipboard := ClipboardAll
Clipboard =
;Copy current selection, continue if no errors.
SendInput, ^c
ClipWait, 2
if !(ErrorLevel) {
;Convert Clipboard to text, auto-trim leading and trailing spaces and tabs.
Clipboard = %Clipboard%
;Clean Clipboard: change carriage returns to spaces, change >=1 consecutive spaces to +
Clipboard := RegExReplace(RegExReplace(Clipboard, "\r?\n"," "), "\s+","+")
;Open URLs, Google non-URLs. URLs contain . but do not contain + or .. or @
if Clipboard contains +,..,@
; NOTE: FOR A DIFFERENT LANGUAGE, CHANGE "FREN" in url below
Run, %browser% http://www.wordreference.com/fren/%Clipboard%
else if Clipboard not contains .
Run, %browser% http://www.wordreference.com/fren/%Clipboard%
else
Run, %browser% %Clipboard%
}
;Restore Clipboard, clear prevClipboard variable.
Clipboard := prevClipboard
prevClipboard =
return

; ----------------------------
; Search Termium for selection
; ----------------------------

; IMPORTANT: Set the correct path to your default broswer
browser="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

#+t::
;Copy Clipboard to prevClipboard variable, clear Clipboard.
prevClipboard := ClipboardAll
Clipboard =
;Copy current selection, continue if no errors.
SendInput, ^c
ClipWait, 2
if !(ErrorLevel) {
;Convert Clipboard to text, auto-trim leading and trailing spaces and tabs.
Clipboard = %Clipboard%
;Clean Clipboard: change carriage returns to spaces, change >=1 consecutive spaces to +
Clipboard := RegExReplace(RegExReplace(Clipboard, "\r?\n"," "), "\s+","+")
;Open URLs, Google non-URLs. URLs contain . but do not contain + or .. or @
if Clipboard contains +,..,@
Run, %browser% http://www.btb.termiumplus.gc.ca/tpv2alpha/alpha-eng.html?lang=eng&i=1&srchtxt=%Clipboard%
else if Clipboard not contains .
Run, %browser% http://www.btb.termiumplus.gc.ca/tpv2alpha/alpha-eng.html?lang=eng&i=1&srchtxt=%Clipboard%
else
Run, %browser% %Clipboard%
}
;Restore Clipboard, clear prevClipboard variable.
Clipboard := prevClipboard
prevClipboard =
return

; ----------------------------------
; +l### Search Linguee for selection 
; ----------------------------------

; IMPORTANT: Set the correct path to your default broswer
browser="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

#+l::
;Copy Clipboard to prevClipboard variable, clear Clipboard.
prevClipboard := ClipboardAll
Clipboard =
;Copy current selection, continue if no errors.
SendInput, ^c
ClipWait, 2
if !(ErrorLevel) {
;Convert Clipboard to text, auto-trim leading and trailing spaces and tabs.
Clipboard = %Clipboard%
;Clean Clipboard: change carriage returns to spaces, change >=1 consecutive spaces to +
Clipboard := RegExReplace(RegExReplace(Clipboard, "\r?\n"," "), "\s+","+")
;Open URLs, Google non-URLs. URLs contain . but do not contain + or .. or @
if Clipboard contains +,..,@
; NOTE: TO CHANGE LANGUAGES, MODIFY "/english-french/" below 
Run, %browser% http://www.linguee.com/english-french/search?source=auto&query=%Clipboard%
else if Clipboard not contains .
Run, %browser% http://www.linguee.com/english-french/search?source=auto&query=%Clipboard%
;Run, %browser% %Clipboard% PROBLEMATIC; COMMENTED OUT;
}
;Restore Clipboard, clear prevClipboard variable.
Clipboard := prevClipboard
prevClipboard =
return

; --------------------------
; Selection to thesaurus.com
; --------------------------

; IMPORTANT: Set the correct path to your default broswer
browser="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

#+;::
;Copy Clipboard to prevClipboard variable, clear Clipboard.
prevClipboard := ClipboardAll
Clipboard =
;Copy current selection, continue if no errors.
SendInput, ^c
ClipWait, 2
if !(ErrorLevel) {
;Convert Clipboard to text, auto-trim leading and trailing spaces and tabs.
Clipboard = %Clipboard%
;Clean Clipboard: change carriage returns to spaces, change >=1 consecutive spaces to +
Clipboard := RegExReplace(RegExReplace(Clipboard, "\r?\n"," "), "\s+","+")
;Open URLs, Google non-URLs. URLs contain . but do not contain + or .. or @
if Clipboard contains +,..,@
Run, %browser% http://www.thesaurus.com/browse/%Clipboard%
else if Clipboard not contains .
Run, %browser% http://www.thesaurus.com/browse/%Clipboard%
else
Run, %browser% %Clipboard%
}
;Restore Clipboard, clear prevClipboard variable.
Clipboard := prevClipboard
prevClipboard =
return

; ------------------------------
; RANDOM STUFF (NOT TRANSLATION)
; ------------------------------

; --------------------------------
; Spotify (next, prev, play/pause)
; --------------------------------

; "CTRL + LEFT"  for previous 
^+-::Media_Prev

;"CTRL + RIGHT"  for next 
^+=::Media_Next

;^+p::"CTRL + P"  for pause
^+p::Media_Play_Pause

; ----------------------------
; HOME ROW COMPUTING REMAPPING
; ----------------------------

; Created by Gustavo Duarte; modified to my own preferences for writing/editing prose
; See http://duartes.org/gustavo/blog/home-row-computing for more information on this script
; NOTE: WILL NOT WORK until you complete OS-level remapping as described at link above


; ------------------------------------------
; Appskey = CAPS Lock (mapped with Sharpkeys)
; NOTE: AppsKey on its own often performs 
; "right click" operations,
; -------------------------------------------



; --------------------------------------------
; AppsKey + Right hand, home row (j, k, l, ;)
; --------------------------------------------
; j = down; k = up; l = left; semicolon = right 
; NOTE switched j and k
; --------------------------------------------

Appskey & j::Send {Blind}{Down DownTemp}
AppsKey & j up::Send {Blind}{Down Up}

AppsKey & k::Send {Blind}{Up DownTemp}
AppsKey & k up::Send {Blind}{Up Up}

AppsKey & l::Send {Blind}{Left DownTemp}
AppsKey & l up::Send {Blind}{Left Up}

AppsKey & `;::Send {Blind}{Right DownTemp}
AppsKey & `; up::Send {Blind}{Right Up}


; AppsKey + Right hand, upper row (u,i,o,p)
; ----------------------------------------------------------------------------------------
; u = erase word back; i = erase word forward; o = skip back 1 word; p = skip ahead 1 word
; NOTE: "u" and "i" mappings (full word delete) especially beneficial for writing
; NOTE: Alternate versions provided for use in gvim; imperfect
; ----------------------------------------------------------------------------------------


AppsKey & u::SendInput {Ctrl down}{BackSpace}{Ctrl up}
AppsKey & u up::SendInput {Blind}{PgUp Up}
; FOR VIM
; AppsKey & u::SendInput {Ctrl down}{w}{Ctrl up}

AppsKey & i::SendInput {Ctrl down}{Delete}{Ctrl up}
AppsKey & i up::SendInput {Blind}{PgDn Up}
; FOR VIM
; AppsKey & i::SendInput {Esc}{Space}{d}{a}{w}{space}{i}

AppsKey & o::SendInput {Ctrl down}{Left}{Ctrl up}
AppsKey & o up::SendInput {Blind}{Ctrl + Left Up}

AppsKey & p::SendInput {Ctrl down}{Right}{Ctrl up}
AppsKey & p up::SendInput {Blind}{Ctrl Right Down}

----------------------------------------------
;;ADDED UNDO
AppsKey & b::SendInput {Ctrl down}{z}{Ctrl up}
----------------------------------------------


--------------------------------------------
; AppsKey + left hand, home row (a, s, d, f)
; a = select all; s = cut; d = copy; f = paste
;---------------------------------------------
; NOTE: More ergnomic cutting and pasting
; --------------------------------------------

AppsKey & a::SendInput {Ctrl Down}{a Down}
AppsKey & a up::SendInput {Ctrl Up}{a Up}

AppsKey & s::SendInput {Ctrl Down}{x Down}
AppsKey & s up::SendInput {Ctrl Up}{x Up}

AppsKey & d::SendInput {Ctrl Down}{c Down}
AppsKey & d up::SendInput {Ctrl Up}{c Up}

AppsKey & f::SendInput {Ctrl Down}{v Down}
AppsKey & f up::SendInput {Ctrl Up}{v Up}
			
AppsKey::SendInput {AppsKey Down}
AppsKey up::SendInput {AppsKey Up}


; ADDED KEYS FOR FREQUENT ACTIONS
; --------------------------------
; AppsKey + h = delete; n = backspace; r = Escape;
; NOTE: Escape = important to have when Winkey sticks!
; ----------------------------------------------------

AppsKey & h::SendInput {Blind}{Del Down}
AppsKey & w::SendInput {Ctrl down}{F4}{Ctrl up}
AppsKey & e::SendInput {Alt down}{F4}{Alt up}
AppsKey & n::SendInput {Blind}{BS Down}
AppsKey & BS::SendInput {Blind}{BS Down}
AppsKey & r::SendInput {Blind}{Esc Down}

; Make AppsKey & Enter equivalent to Control+Enter
AppsKey & Enter::SendInput {Ctrl down}{Enter}{Ctrl up}

; Make AppsKey & Alt Equivalent to Control+Alt
!AppsKey::SendInput {Ctrl down}{Alt Down}
!AppsKey up::SendInput {Ctrl up}{Alt up}

; Make Windows Key + Apps Key work like Caps Lock
#AppsKey::Capslock



; ------------------------------------------
; LAUNCH APPS (WinKey 1-9) FROM HOMEROW
; NOTE: Remaps right Winky to right Alt; advisability depends on keyboard
; -----------------------------------------------------------------------


;Remap right alt to Windows key 
RAlt::Rwin

;Remap easy Alt+F4 to close windows
RAlt & v::Send !{F4}

;Left hand=launch apps 1-5 in dock
RAlt & a::
SendInput {RWin Down}{1}
SendInput {Rwin Up}
return

RAlt & s::
SendInput {RWin Down}{2}
SendInput {Rwin Up}
return


RAlt & d::
SendInput {RWin Down}{3}
SendInput {Rwin Up}
SendInput {Enter}
return 

RAlt & f::
SendInput {RWin Down}{4}
SendInput {Rwin Up}
return 

RAlt & g::
SendInput {RWin Down}{5}
SendInput {Rwin Up}
return 

;Right hand=launch apps 6-10 in dock; 
;NO SEMICOLON= "i" to launch app #9


RAlt & h::
SendInput {RWin Down}{6}
SendInput {Rwin Up}
return

RAlt & j::
SendInput {RWin Down}{7}
SendInput {Rwin Up}
return

RAlt & k::
SendInput {RWin Down}{8}
SendInput {Rwin Up}
return

RAlt & i::
SendInput {RWin Down}{9}
SendInput {Rwin Up}
return


















; -----------
; IN WORDFAST
; -----------

;----------------------------------
;copy source and go to next segment 
;----------------------------------

!W::

Send !{c}
Sleep, 200
Send !{f}
Sleep, 100
return


;---------------------------------------------
; erase one word forward but not the tag after
; NOTE: NEEDS WORK
; --------------------------------------------

!y::
SendInput {Ctrl down}{Right}
SendInput {Ctrl up}
;Sleep, 50
SendInput {Left}
SendInput {Ctrl down}{Backspace}
SendInput {Ctrl up}
Sleep, 50
return

;----------------------------------------------------
; erase to end of segment, commit, go to next segment
; ---------------------------------------------------
!F8::
Send {Shift Down}{Down 4}
Send {Shift Up}
Sleep, 100
Send {Delete}
Sleep, 200
Send !{f}
Sleep, 400
Send ^+{v}
Sleep, 400
Send ^{Home}
return

; -----------------------
; erase to end of segment
; -----------------------

!F9::
Send {Shift Down}{Down 4}
Send {Shift Up}
Sleep, 100
Send {Delete}
Sleep, 200




; -----------------------------
; OPEN ALL COPIED LINKS IN TABS
; -----------------------------


;^!c::
  oCB := ClipboardAll
  Send ^c
  Loop,parse,clipboard,`n,`r 
  {
    Run %A_LoopField%
  }
  ClipBoard := %oCB%
;-----------------------------------------------------------*
