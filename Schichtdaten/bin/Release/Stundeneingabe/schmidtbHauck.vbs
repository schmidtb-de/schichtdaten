Public Function WaitForWindow(WindowTitle)

Set WshShell = WScript.CreateObject("WScript.Shell")

success=0

I = 0

Do

I = I + 1

WScript.Sleep 300

success=WshShell.AppActivate(WindowTitle)

If I = 50 Then

x=MsgBox ("Fenster "+Chr(34)& WindowTitle &Chr(34)+" konnte nicht gefunden werden.",48,"Warnmeldung")

WaitForWindow = False

asyncConnection.Disconnect(2)

WScript.quit

Exit Do

End If

Loop Until success

If success Then

WaitForWindow = True

End If

End Function

'Deklaration

set shell = CreateObject("WScript.Shell")

'Ausführen des Programmes

WScript.Sleep 500

shell.AppActivate "SINOx-Info"

'Tastendruck simulieren

WaitForWindow("Anmelden")

WScript.Sleep 100

shell.SendKeys "schmidtb"

WScript.Sleep 100

shell.SendKeys "{TAB}"

shell.SendKeys "test1234"

WScript.Sleep 100

shell.SendKeys "{ENTER}"

WaitForWindow("SQL Server Login")

WScript.Sleep 100

shell.SendKeys "{TAB}"

WScript.Sleep 100

shell.SendKeys "schmidtb"

WScript.Sleep 100

shell.SendKeys "{TAB}"

WScript.Sleep 100

shell.SendKeys "test1234"

WScript.Sleep 100

shell.SendKeys "{TAB}"

WScript.Sleep 100

shell.SendKeys "{ENTER}"

WaitForWindow("Anmelden")

WScript.Sleep 100

shell.SendKeys "schmidtb"

WScript.Sleep 100

shell.SendKeys "{TAB}"

WScript.Sleep 100

shell.SendKeys "test1234"

WScript.Sleep 100

shell.SendKeys "{ENTER}"

WaitForWindow("SQL Server Login")

WScript.Sleep 100

shell.SendKeys "{TAB}"

WScript.Sleep 100

shell.SendKeys "schmidtb"

WScript.Sleep 100

shell.SendKeys "{TAB}"

WScript.Sleep 100

shell.SendKeys "test1234"

WScript.Sleep 100

shell.SendKeys "{TAB}"

WScript.Sleep 100

shell.SendKeys "{ENTER}"

WScript.Sleep 5000

shell.SendKeys "test1234"

WScript.Sleep 100

shell.SendKeys "{ENTER}"

WaitForWindow("SQL Server Login")

WScript.Sleep 100

shell.SendKeys "{TAB}"
WScript.Sleep 100
shell.SendKeys "schmidtb"

WScript.Sleep 100
shell.SendKeys "{TAB}"
WScript.Sleep 100
shell.SendKeys "test1234"

WScript.Sleep 100
shell.SendKeys "{TAB}"
WScript.Sleep 100
shell.SendKeys "{ENTER}"
WScript.Sleep 1500
shell.SendKeys "{TAB}"
shell.SendKeys "{RIGHT}"
shell.SendKeys "{RIGHT}"
shell.SendKeys " "
