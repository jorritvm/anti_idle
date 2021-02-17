Dim objResult

Set objShell = WScript.CreateObject("WScript.Shell")    

Do While True
  objResult = objShell.sendkeys("{F15}")

  Wscript.Sleep (4000)
Loop