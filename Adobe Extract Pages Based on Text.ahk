#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force
SetTitleMatchMode, 2


;GUI Interace

Gui, Add, GroupBox, x12 y69 w320 h50 , Search Term
Gui, Add, Edit, x22 y89 w300 h20 vSearchterm,
Gui, Add, GroupBox, x12 y139 w320 h290 , Search List
Gui, Add, Edit, x22 y159 w300 h260 vSearchlist, 
Gui, Add, Button, x123 y449 w100 h30 gExtract Default, Extract Pages
Gui, Add, Text, x22 y19 w310 h40 , Extract pages from active PDF based on the search terms OR search list below.
Gui, Show, x262 y148 h495 w350, PDF Page Extractor 
Return

GuiClose:
ExitApp

;Button Route
Extract:
	msgbox Please activate the PDF window that you want to extract pages from - and then press OK.  Program reference the active PDF file.
	Gui, Submit, Nohide
	;
	;
	;Check if we are using the single term method or using a list
	;
	;
	
	if (Searchterm != "")
	{
		if (Searchlist != "")
		{
			Msgbox You cannot have a search term and a search list.
			return
		}
	}
	
	;
	;
	;if single searchterm method, then search each page within PDF for the search term.  If it exists, export the page to a new PDF document
	;
	;
	if Searchterm != 
	{
			
		app := ComObjCreate("AcroExch.App")

		avDoc := app.GetActiveDoc
		avPageView := avDoc.GetAVPageView
		avPageView.goto(0)

		pdDoc := avDoc.GetPDDoc



		numPages := pdDoc.GetNumPages
		msgbox % "Number of pages: " numPages


		
		new_pdDoc := ComObjCreate("AcroExch.PDDoc") 
		r := new_pdDoc.Create
		if r = 0
		{
			msgbox there was an issue creating new pddoc
			exitapp
		}
		loop
		{
				
			r := avDoc.FindText(Searchterm, 0, 0, 1)
			if r = 0 
			{
				msgbox cannot find text
				return
			}
			pageFound := avPageView.GetPageNum
			sleep 400

			if A_Index = 1
				first_find := pageFound
			
			if (pageFound = first_find)
			{
				if A_Index != 1
				{
					break
				}
			}
			
			
			newDocPages := new_pdDoc.GetNumPages
			r := new_pdDoc.InsertPages(newDocPages-1, pdDoc, pageFound, 1, 0)
			if r = 0
			{
				msgbox could not insert page %pageFound%
				continue
			}


		}
		
		r := new_pdDoc.Save(1, A_WorkingDir "\Extracted_Pages.pdf")
	
	}
	else if  Searchlist !=
	{
		Searchlist := StrReplace(Searchlist, "`r", "")		
		app := ComObjCreate("AcroExch.App")

		avDoc := app.GetActiveDoc
		avPageView := avDoc.GetAVPageView
		avPageView.goto(1)

		pdDoc := avDoc.GetPDDoc



		numPages := pdDoc.GetNumPages
		msgbox % "Number of pages: " numPages


		
		new_pdDoc := ComObjCreate("AcroExch.PDDoc") 
		r := new_pdDoc.Create
		
		loop, parse, Searchlist, `n	
		{
			if A_Loopfield = 
				continue
			Tooltip % "Searching for term: " A_LoopField
			r := avDoc.FindText(A_LoopField, 0, 0, 0)

			if r = 0 
			{
				msgbox Cannot find text: %A_LoopField%
				Continue
			}
			pageFound := avPageView.GetPageNum
			sleep 400

			newDocPages := new_pdDoc.GetNumPages
			r := new_pdDoc.InsertPages(newDocPages-1, pdDoc, pageFound, 1, 0)
			if r = 0
			{
				msgbox Could not add page for search term: %A_LoopField%
				continue
			}


		}
		ToolTip,

		
		r := new_pdDoc.Save(1, A_WorkingDir "\Extracted_Pages.pdf")
		if r = -1
		{
			msgbox Extracted pages exported to: %A_WorkingDir% \Extracted_Pages.pdf
		}
		else
		{
			msgbox Error saving file
		}
	}
Return



esc::exitapp
