Write-Host -ForegroundColor Cyan "Preparing Data For Export"
$initialDirectory = "C:\"
$mailboxes = Get-Mailbox
$MBData = foreach($mailbox in $mailboxes.Alias){Get-MailboxStatistics -Identity $mailbox | Select DisplayName, TotalItemSize}
Function Save-File ([string]$initialDirectory) {
    $SaveInitialPath = "C:\"
    $SaveFileName = "Result.csv"
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $SaveInitialPath
    $OpenFileDialog.FileName = $SaveFileName
    $OpenFileDialog.ShowDialog() | Out-Null
    return $OpenFileDialog.filename
}
$SaveMyFile = Save-File
$MBData | Export-CSV -Path $SaveMyFile
