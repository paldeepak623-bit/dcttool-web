Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Filter = "Excel Workbooks (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"
$dialog.Title = "Select PACS Master Excel Sheet"
$dialog.RestoreDirectory = $true
$dialog.ShowHelp = $false
$form = New-Object System.Windows.Forms.Form
$form.TopMost = $true
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.Opacity = 0
$form.Show()
$res = $dialog.ShowDialog($form)
$form.Close()
if ($res -eq [System.Windows.Forms.DialogResult]::OK) {
    Write-Output $dialog.FileName
}
