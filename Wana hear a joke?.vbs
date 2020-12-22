OPTION EXPLICIT
on error resume next
on error goto 0

DIM dct
SET dct = CreateObject("Scripting.Dictionary")
SET dct = nothing

'SUB window_onload()
	call fnJoke("What Do you do to a bacon hair in jailbreak?", answer())
'END SUB

Function answer()
  answer = "You FRY them."
END FUNCTION

SUB fnJoke(p_strJoke, p_strPunchline)
  DIM strName, strPrompt
  strName = inputbox("What's your name?", "important thing", "Anonymous")
  strPrompt = "Hello " & strName & ", wanna hear a joke?"
  IF msgbox(strPrompt, vbOKCancel)=vbOK THEN
    msgbox p_strJoke & vbCRLF & vbCRLF & p_strPunchline
  ELSEIF strName="Anonymous" THEN
    ' tell it anyway
    msgbox "i'll tell ya anyway." & vbCRLF & _
      p_strJoke & vbCRLF & vbCRLF & p_strPunchline
  ELSE	' doesn't wanna hear it
    call msgbox("you are a PARTY POOPER, Mr. " & strName & "!")
  END IF
END Sub

CONST testing = 2
SELECT CASE testing
	CASE 1
		msgbox("Made By AWIRE9966 On Github.")
	CASE 2: msgbox("Made By AWIRE9966 On Github.")
	CASE 3
		msgbox("Made By AWIRE9966 On Github.")
END SELECT
