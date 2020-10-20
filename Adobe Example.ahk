

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; triggers when Ctrl + Shift + 8 is pressed
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
+^8::
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; create Adobe objects
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
app := ComObjCreate("AcroExch.App")
avDoc := app.GetActiveDoc
pdDoc := avDoc.GetPDDoc
jso := pdDoc.GetJSObject
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; iterate through all fields on active PDF window
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
numFields := jso.numFields
myobj := {}
Qtys := []
mfgs := []
descs := []
pns := []
partlinecounter := 0
Loop, % numFields
{
	ToolTip, % A_Index
	key := jso.getNthFieldName(A_Index)
	value := jso.getField(jso.getNthFieldName(A_Index)).Value
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; collect data from only the fields needed
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		
	if InStr(key, "QtyRow")
	{
		if !InStr(key, "BOM")
		{
			StringRight, num, key, 2
			if num is not integer
				StringRight, num, key, 1
			Qtys.InsertAt(num, value)
		}
	}
	if InStr(key, "Part Row")
	{
		StringRight, num, key, 2
		if num is not integer
			StringRight, num, key, 1
		pns.InsertAt(num, value)
	}
}

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; format data to easily mass-upload to a shopping cart.  
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

clip := ""
Loop, % Qtys.MaxIndex()
{
	if (Qtys[A_Index] = "" and  mfgs[A_Index] = "")
	{
		if (descs[A_Index] = "" and pns[A_Index] = "")
		{
			Continue
		}
	}
	qty := Format("{:i}", Qtys[A_Index])
	pn := Format("{:s}", pns[A_Index])
	clip := clip pn "," qty "`n"
	qty := ""
	pn := ""
	
}
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; Display a user friendly message confirming script is done, data copied to clipboard 
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
ToolTip,
clipboard := clip
Msgbox % clip
Return
