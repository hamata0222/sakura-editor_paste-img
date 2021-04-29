Dim exec_script
Dim filename

filename = Replace(Replace(Replace(Now(), "/", ""), ":", ""), " ", "")

exec_script = "powershell.exe -sta -WindowStyle Hidden -Command Add-Type -Assembly System.Windows.Forms;" _
& " if (!([Windows.Forms.Clipboard]::ContainsImage())) {exit} ;" _
& " [System.Windows.Forms.Clipboard]::GetImage().Save('./" & filename & ".png');" _
& " Echo '![./" & filename & ".png](./" & filename & ".png)'"

Set exec = CreateObject("WScript.Shell").Exec(exec_script)
Editor.InsText(exec.StdOut.ReadAll)
