#include <IE.au3>
#include <array.au3>
;
;
;"InnCrowd" is an applet designed to automate scheduling for installations.  With this utility you could potentially schedule twenty or thirty appointments within ten minutes. 
;1.  InnCrowd takes up to five arguments divided by spaces: InnCode Time Date HPRep DHTech
;
;	a. InnCode, Time and Date are always required, if they are not present InnCrowd will crash and exit. 
;	   i.	InnCode, Time and Date must be in that order, HPRep and DHTech can swap at any time.
;	   ii. 	If the InnCode does not exsist, InnCrowd displays a friendly error and exits.
;	   iii.	If time or date are incorrectly formatted, InnCrowd will crash and exit.
;	b. HPRep
;	   i. 	HPRep is required the first time you run but any entry will be accepted.
;	   ii.	Afterwards it will store the given value until the application is terminated. 
;	   iii. HPRep can also be overwritten at anytime by appending it as the 4th or 5th argument.
;	c. Tech
;	   i. 	DHTech is never required, if you don't enter one InnCrowd will select the user that currently covers that shift.
;	   ii.  If a valid tech is not entered InnCrowd will silently select the default tech for that shift.
;
;"InnCrowd-SafeMode.exe" can be run without creating any cases/events.
;
;

;pause or exit using "Pause|Break" or "Esc"
Global $Paused
HotKeySet("{PAUSE}", "TogglePause")
HotKeySet("{ESC}", "Terminate")

Global $oIE = _IEAttach("website_instance")
Global $dhRep
Global $HPRep
Global $user = _IEPropertyGet(_IEGetObjById($oIE, "userNavLabel"), "innertext")
Global $appointments = 2

For $i = 0 To $appointments - 1

	;prompts the user for the name of the HP Rep, inn code, date and time
	Global $request = InputBox("Inn Crowd", "Request: ", "EMPEX 4p 12/11/2013 Linda")
	Global $input = UBound(StringSplit($request, " ")) - 1
	If $input > 5 Then
		MsgBox(0, "Input", "Too much input")
		Exit 0
	EndIf

	Global $event[$input]; 0 = code, 1 = time, 2 = date, 3 = tech 4 = hprep,

	;parses Input
	Call("newEvent", $request, $input)
	Global $innCode = $event[0]

	;reparses time into a second array (fix later)
	Global $timeDate = $event[1] & " " & $event[2]
	Global $timeSlot = Call("getTime", $timeDate) ;parses the time and returns an array

	If $input > 4 Then
		;MsgBox(0, "Switch One", "Switch One")
		;MsgBox (0, "lower", StringLower($event[3]))
		Switch StringLower($event[3])
			Case "gerard"
				$dhRep = "Gerard"
				$HPRep = $event[4]
			Case "dan"
				$dhRep = "Dan"
				$HPRep = $event[4]
			Case "martin"
				$dhRep = "Martin"
				$HPRep = $event[4]
			Case "tabatha"
				$dhRep = "Tabatha"
				$HPRep = $event[4]
			Case "victor"
				$dhRep = "Victor "
				$HPRep = $event[4]
			Case Else
				;MsgBox(0, "Switch Two", "Switch Two")
				Switch StringLower($event[4])
					Case "gerard"
						$dhRep = "Gerard"
						$HPRep = $event[3]
					Case "dan"
						$dhRep = "Dan"
						$HPRep = $event[3]
					Case "martin"
						$dhRep = "Martin"
						$HPRep = $event[3]
					Case "tabatha"
						$dhRep = "Tabatha"
						$HPRep = $event[3]
					Case "victor"
						$dhRep = "Victor "
						$HPRep = $event[3]
					Case Else
						$HPRep = InputBox("Inn Crowd", "HP Rep: ", "")
						$dhRep = Call("dhRep", $timeSlot);selects the agent to process the ticket using values from the $timeSlot array
				EndSwitch
		EndSwitch

	ElseIf ($input > 3) And ($HPRep <> "") Then
		Switch StringLower($event[3])
			Case "gerard"
				$dhRep = "Gerard"
			Case "dan"
				$dhRep = "Dan"
			Case "martin"
				$dhRep = "Martin"
			Case "tabatha"
				$dhRep = "Tabatha"
			Case "victor"
				$dhRep = "Victor "
			Case Else
				$HPRep = $event[3]
				$dhRep = Call("dhRep", $timeSlot);selects the agent to process the ticket using values from the $timeSlot array
		EndSwitch

	ElseIf ($input > 3) And ($HPRep == "") Then

		Switch StringLower($event[3])
			Case "gerard"
				$dhRep = "Gerard"
				$HPRep = InputBox("Inn Crowd", "HP Rep: ", "")
			Case "dan"
				$dhRep = "Dan"
				$HPRep = InputBox("Inn Crowd", "HP Rep: ", "")
			Case "martin"
				$dhRep = "Martin"
				$HPRep = InputBox("Inn Crowd", "HP Rep: ", "")
			Case "tabatha"
				$dhRep = "Tabatha"
				$HPRep = InputBox("Inn Crowd", "HP Rep: ", "")
			Case "victor"
				$dhRep = "Victor "
				$HPRep = InputBox("Inn Crowd", "HP Rep: ", "")
			Case Else ;assigns the last argument to HPRep and automatically selects the default tech
				$HPRep = $event[3]
				$dhRep = Call("dhRep", $timeSlot);selects the agent to process the ticket using values from the $timeSlot array
		EndSwitch

	ElseIf ($HPRep == "") Then
		$HPRep = InputBox("Inn Crowd", "HP Rep: ", "")
		$dhRep = Call("dhRep", $timeSlot);selects the agent to process the ticket using values from the $timeSlot array
	Else
		$dhRep = Call("dhRep", $timeSlot);selects the agent to process the ticket using values from the $timeSlot array
	EndIf

	;searches for the innCode
	Call("search", $innCode)

	;selects a WO from a list of cases to clone
	Call("selectCase", $innCode)

	;selects a WO from a list of cases to clone
	Call("cloneCase", $user, $dhRep, $timeSlot)

	;adds the appointment to the calendar
	$Case = Call("schedule", $timeSlot, $innCode, $dhRep, $HPRep)

	;display case number and start over
	MsgBox(0, "Case Number", "Case number: " & $Case)
	$i = 0

Next

Func newEvent($request, $input)

	For $i = 0 To ($input - 1)

		;returns the location of the space in $timeSlot and uses it to find the time
		If (StringInStr($request, " ")) Then
			$space = StringInStr($request, " ")

			;uses $space to parse the next entry into the array
			$event[$i] = StringTrimRight($request, StringLen($request) - ($space - 1))
			$request = StringTrimLeft($request, StringLen($event[$i]) + 1)

			;adds a tech to the request if present
		ElseIf $request <> $event[$i - 1] Then
			$event[$i] = $request

		EndIf
	Next
	Return $event
EndFunc   ;==>newEvent

Func dhRep($timeSlot)

	;converts standard time to military time
	$timeMod = ""
	If (StringLower($timeSlot[0]) == "a") And ($timeSlot[1] == 12) Then
		$timeMod = 0
	ElseIf (StringLower($timeSlot[0]) == "p") And ($timeSlot[1] < 12) Then
		$timeMod = $timeSlot[1] + 12
	Else
		$timeMod = $timeSlot[1]
	EndIf

	;Selects the user that will work the case
	Switch $timeMod
		Case 0 < $timeMod And $timeMod < 7
			$dhRep = "Martin"
		Case 7 <= $timeMod And $timeMod <= 15
			$dhRep = "Victor "
		Case 16 <= $timeMod And $timeMod < 23
			$dhRep = "Tabatha"
		Case Else
			$dhRep = "Dan"
	EndSwitch
	Return $dhRep

EndFunc   ;==>dhRep

Func getTime($timeSlot)

	;returns the location of the space in $timeSlot and uses it to find the time
	Local $space = StringInStr($timeSlot, " ")
	Local $hour[3]; 0 = hour, 1 = ampm, 2 = day

	;uses $space to find the hour portion of input
	$hour[0] = StringTrimRight($timeSlot, StringLen($timeSlot) - ($space - 1))
	$ampm = StringLen($hour[0]) - 1
	$hour[0] = StringTrimLeft($hour[0], $ampm)

	;uses $space to find the $ampm portion of input
	$hour[1] = StringTrimRight($timeSlot, StringLen($timeSlot) - ($space - 2))

	;uses $space to find the day portion of input
	$hour[2] = StringTrimLeft($timeSlot, $space)
	;_ArrayDisplay($hour, "Array Values of $hour")

	Return $hour

EndFunc   ;==>getTime

Func search($sbstr)

	;Creates referance to the search field & submits with the innCode
	$innSearch = _IEGetObjById($oIE, "sbstr")
	_IEFormElementSetValue($innSearch, $innCode)

	;may remove this piece here
	Local $go = _IEGetObjByName($oIE, "search")
	_IEAction($go, "click")
	Sleep(2000)

EndFunc   ;==>search

Func selectCase($innCode)

	;gets the number of accounts from the top of the accounts table and saves the numeric value to $acctCount
	$oTable = _IETableWriteToArray(_IETableGetCollection($oIE, 2), True)
	$acctCount = StringTrimRight(StringTrimLeft($oTable[0][0], 10), 2)

	;checks the table header and reads Cases[] to a variable
	$oTable = _IETableWriteToArray(_IETableGetCollection($oIE, 6), True)
	If(StringInStr($oTable[0][0],"Contacts"))Then
		$oTable = _IETableWriteToArray(_IETableGetCollection($oIE, 12), True)
		$caseCount = StringTrimRight(StringTrimLeft($oTable[0][0], 7), 2)
		$oTable = _IETableGetCollection($oIE, 15)
	Else
		$caseCount = StringTrimRight(StringTrimLeft($oTable[0][0], 7), 2)
		$oTable = _IETableGetCollection($oIE, 9)
	EndIf

	;reads the whole list of cases for the loop
	$cases = _IETableWriteToArray($oTable, True)
	_ArrayDisplay($cases)

	Local $stored = "False"
	Local $closed = "Closed"
	Local $conflictCount = 0
	Local $conflict = "On Hold - Scheduled"
	Local $wo = "Work Order Deployment"

	For $j = 1 To $caseCount
		If ($cases[$j][4] Not $closed) Then
			$conflictCount++
			MsgBox(0, "Schedule Conflict", $cases[$j][2] & " " & $cases[$j][3] & " is " & $cases[$j][4])
			$selectCase = $cases[$j][2]
			Exit 0
		ElseIf (($cases[$j][7] == $wo) And ($stored <> "True")) Then
			$selectCase = $cases[$j][2]
			$stored = "True"
		Else
			;do nothing
		EndIf
	Next

	If $stored == "False" Then
		;no previous WO to clone
		MsgBox(0, "WO Not Found", "No exsisting work orders for InnCode: " & $innCode)
		Exit 0
		;need to select the primary account from the list to view history
	Else
		;click the link and wait for the page to load
		_IELinkClickByText($oIE, $selectCase)
		_IELoadWait($oIE)
		Sleep(2000)
	EndIf

EndFunc   ;==>selectCase

Func cloneCase($user, $dhRep, $timeSlot)

	;Creates referance to the Clone button & submits
	Local $clone = _IEGetObjByName($oIE, "clone")
	_IEAction($clone, "click")
	Sleep(2000)

	;creates a referance to the "Site Type" and selects "HW Refresh"
	$siteType = _IEGetObjById($oIE, "00N80000002yuyo")
	$type = "HW Refresh"
	_IEFormElementOptionSelect($siteType, $type, 1, "byText")

	;creates a referance to the "Site Type" and selects "HW Refresh"
	$status = _IEGetObjById($oIE, "cas7")
	$scheduled = "On hold - Scheduled"
	_IEFormElementOptionSelect($status, $scheduled, 1, "byText")

	;Creates referance to the "Subject" and updates it with PMS Refresh
	$subject = _IEGetObjById($oIE, "cas14")
	$pmsSubject = $innCode & " PMS Refresh"
	_IEFormElementSetValue($subject, $pmsSubject)

	;Creates referance to the "Description" and updates it with PMS Refresh
	$description = _IEGetObjById($oIE, "cas15")
	$pmsDescription = $innCode & " PMS Refresh"
	_IEFormElementSetValue($description, $pmsDescription)

	;creates a reference to the "FastConnect Installation Date" and updates it with the request
	$dateTime = _IEGetObjById($oIE, "00N800000031upt")
	_IEFormElementSetValue($dateTime, $timeSlot[2] & " " & $timeSlot[1] & ":00" & " " & $timeSlot[0] & "M")

	;creates a referance to the "Internal comments field and notes the date/time & Rep Scheduled with
	$comments = _IEGetObjById($oIE, "cas16")
	$schduleRep = "Scheduled with " & $HPRep & " for " & $timeSlot[2] & " " & $timeSlot[1] & ":00" & " " & $timeSlot[0] & "M" ;this needs to be looked at
	_IEFormElementSetValue($comments, $schduleRep)

	;Creates referance to the save button and calls submit()
	Local $save = _IEGetObjByName($oIE, "save")
	Call("submit", $save)

	If ($dhRep <> $user) Then
		;MsgBox(0, "Equality", $dhRep & " is not " & $user & ", assigning " & $dhRep & " to case.")
		;clicks the "Change" link & modifies the user assigned to the ticket
		_IELinkClickByText($oIE, "[Change]")
		Sleep(2000)

		;creates a referance to the user field and updates it to the assigned user
		$owner = _IEGetObjById($oIE, "newOwn")
		_IEFormElementSetValue($owner, $dhRep)

		;Creates referance to the save button and calls submit()
		Local $save = _IEGetObjByName($oIE, "save")
		Call("submit", $save)
	Else
		;Do Nothing
		;MsgBox(0, "Current User", $dhRep & " equals " & $user)
		;Assigns event[4] to the ticket
		;$owner = _IEGetObjById($oIE, "newOwn")
		;_IEFormElementSetValue($owner, $event[4])

		;this will change the owner of the case when we change the input
		;Local $cancel = _IEGetObjByName($oIE, "cancel")
		;Call("submit", $cancel)
	EndIf

	Return $schduleRep
EndFunc   ;==>cloneCase

Func schedule($timeSlot, $innCode, $dhRep, $HPRep)

	;gets the Case number of the current ticket
	$caseNum = _IEPropertyGet($oIE, "title")
	$count = StringLen($caseNum) - 14
	$caseNum = StringTrimRight(StringTrimLeft($caseNum, 6), $count)

	$newEventURL = "https://na6.website_instance.com/00U/e?retURL=/cal/daily.apexp?cal_lkid=02330000000LIG7&cal_lkold=IHG+Fast+Connect+Schedule&md0=2013&md2=45&md3=311&RecordType=012300000000hfu&&aid=02330000000LIG7&anm=IHG+Fast+Connect+Schedule&ent=Event&evt13="
	$hour = $timeSlot[1] & ":00+"
	$ampm = $timeSlot[0] & "M"
	$date = "&evt4=" & $timeSlot[2]

	;returns the website_instance event code with the date "8:00+AM&evt4=11/8/2013"
	$time = $hour & $ampm & $date
	$scheduleSlot = $newEventURL & $time
	_IENavigate($oIE, $scheduleSlot, 1)
	Sleep(2000)

	;creates a referance to the "Subject" and inserts the InnCode & "PMS Refresh"
	$subject = _IEGetObjById($oIE, "evt5")
	$space = StringInStr($dhRep, " ")
	;uses $space to parse the next entry into the array
	$firstName = StringTrimRight($dhRep, StringLen($dhRep) - ($space - 1))
	$pmsSubject = $innCode & " PMS Refresh - " & $firstName
	_IEFormElementSetValue($subject, $pmsSubject)

	;creates a referance to the "Related To" dropdown and sets it to "Case"
	$relatedTo = _IEGetObjById($oIE, "evt3_mlktp")
	$Case = "Case"
	_IEFormElementOptionSelect($relatedTo, $Case, 1, "byText")

	;creates a ref to the case field and sets it to the case value
	$relatedTo = _IEGetObjById($oIE, "evt3")
	_IEFormElementSetValue($relatedTo, $caseNum)

	;add schedule Rep
	;creates a ref to the FastConnect Work Order field and sets it to the case value
	$wo = _IEGetObjById($oIE, "00N800000031riG")
	_IEFormElementSetValue($wo, $caseNum)

	;selects the button to search for a DH representative
	$invite = _IEGetObjByName($oIE, "new")
	_IEAction($invite, "click")
	_IELoadWait($oIE)
	Sleep(2000)

	;moves navigation to the new window and enters the representative's name
	$oIE = _IEAttach("Select event invitees")
	$srch = _IEGetObjByName($oIE, "srch")
	_IEFormElementSetValue($srch, $dhRep)

	;searches for the rep
	$go = _IEGetObjByName($oIE, "go")
	_IEAction($go, "click")
	Send("{Enter}")
	_IELoadWait($oIE)

	;selects the website_instance ID for the representative searched for
	$oForm = _IEGetObjById($oIE, "srchfrm")

	;selects the website_instance ID associated with DH Techs and uses it to select check boxes
	Switch $dhRep
		Case "Gerard"
			$dhRep = "00530000000xbs8"
		Case "Dan"
			$dhRep = "00530000000i5eJ"
		Case "Martin"
			$dhRep = "00580000003MlSf"
		Case "Tabatha"
			$dhRep = "00580000004ZhBJ"
		Case Else
			$dhRep = "00580000005qRGS" ;Victor
	EndSwitch

	_IEFormElementCheckBoxSelect($oForm, $dhRep, "", 1, "byValue")
	$insert = _IEGetObjByName($oIE, "insert")
	_IEAction($insert, "click")
	_IELoadWait($oIE)

	;add IHG_Scheduling
	$srch = _IEGetObjByName($oIE, "srch")
	_IEFormElementSetValue($srch, "IHG_Scheduling")
	$go = _IEGetObjByName($oIE, "go")
	_IEAction($go, "click")
	Send("{Enter}")
	_IELoadWait($oIE)
	$oForm = _IEGetObjById($oIE, "srchfrm")
	_IEFormElementCheckBoxSelect($oForm, "0038000001EhFmr", "", 1, "byValue")
	$insert = _IEGetObjByName($oIE, "insert")
	_IEAction($insert, "click")
	_IELoadWait($oIE)

	;closes the window and moves navigation back to the main screen
	$done = _IEGetObjByName($oIE, "insert", 1)
	_IEAction($done, "click")
	$oIE = _IEAttach("website_instance")

	;Creates referance to the save button and calls submit()
	Local $save = _IEGetObjByName($oIE, "save")
	Call("submit", $save)
	Return $caseNum
EndFunc   ;==>schedule

Func submit($button)

	;Clicks save and waits for the page to load
	_IEAction($button, "click")
	_IELoadWait($oIE)
	Sleep(2000)
EndFunc   ;==>submit

Func TogglePause()
	$Paused = Not $Paused
	While $Paused
		Sleep(100)
		ToolTip('Script is "Paused"', 0, 0)
	WEnd
	ToolTip("")
EndFunc   ;==>TogglePause

Func Terminate()
	Exit 0
EndFunc   ;==>Terminate
