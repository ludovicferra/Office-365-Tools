#Ludovic Ferra
#https://github.com/ludovicferra

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$TeamsCC                         = New-Object system.Windows.Forms.Form
$TeamsCC.ClientSize              = New-Object System.Drawing.Point(373,173)
$TeamsCC.text                    = "MS Teams Clear Cache"
$TeamsCC.TopMost                 = $false
$TeamsCC.icon                    = "$env:LOCALAPPDATA\Microsoft\Teams\app.ico"
$TeamsCC.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#8f8dfe")

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Utilitaire de nettoyage du cache de l`'application Teams"
$Label1.AutoSize                 = $true
$Label1.width                    = 229
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(15,13)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ButtonClear                     = New-Object system.Windows.Forms.Button
$ButtonClear.text                = "Nettoyer MS Teams"
$ButtonClear.width               = 229
$ButtonClear.height              = 61
$ButtonClear.enabled             = $true
$ButtonClear.location            = New-Object System.Drawing.Point(62,82)
$ButtonClear.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',16,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$ButtonClear.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("#4b53bc")

$checkBoxReOpen                  = New-Object system.Windows.Forms.CheckBox
$checkBoxReOpen.text             = "Réouvrir Teams à la fin du traitement"
$checkBoxReOpen.AutoSize         = $false
$checkBoxReOpen.width            = 343
$checkBoxReOpen.height           = 20
$checkBoxReOpen.location         = New-Object System.Drawing.Point(13,53)
$checkBoxReOpen.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$TeamsCC.controls.AddRange(@($Label1,$ButtonClear,$checkBoxReOpen))


$TeamsCC.FormBorderStyle = 'Fixed3D'
$TeamsCC.MaximizeBox = $false

$checkBoxReOpen.Checked = $true #Etat par défaut de la case réouverture de Teams

function Clear-Teams {
    Write-Output "Arrêt du processus Teams"
    try{
        Get-Process -ProcessName Teams -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 3
        Write-Output "Processus Teams arrêté avec succès"
    } catch { Return $_ }
    Write-Output "Nettoyage du cache Teams"
    Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\application cache\cache" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\blob_storage" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\databases" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\cache" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\gpucache" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Indexeddb" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Local Storage" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\tmp" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    Write-Output "Cache Teams nettoyé"
}

$ButtonClear.Add_Click({ 
    "======= Lancement du nettoyage du cache Teams - $(Get-Date) ======" | Out-File "$env:TEMP\ClearCacheTeams.log" -Encoding utf8
    $ButtonClear.enabled = $false
    $ButtonClear.text = "Nettoyage en cours..."
    Clear-Teams | Out-File "$env:TEMP\ClearCacheTeams.log" -Append -Encoding utf8
    if ($checkBoxReOpen.Checked) { 
        $ButtonClear.text = "Lancement de Teams..."
        Powershell.exe -Command 'Start-Process "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe"'
        Write-Output  "Lancement de Teams" | Out-File "$env:TEMP\ClearCacheTeams.log" -Append -Encoding utf8
        Start-Sleep -Seconds 5
    }
    "======= Processus terminé ======" | Out-File "$env:TEMP\ClearCacheTeams.log" -Append -Encoding utf8
    $ButtonClear.text = "Nettoyage terminé"
})
[void]$TeamsCC.ShowDialog()
