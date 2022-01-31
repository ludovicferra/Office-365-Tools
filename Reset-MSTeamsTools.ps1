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
$TeamsCC.ClientSize              = New-Object System.Drawing.Point(373,245)
$TeamsCC.text                    = "MS Teams Reset Tools"
$TeamsCC.TopMost                 = $false
$TeamsCC.icon                    = "$env:LOCALAPPDATA\Microsoft\Teams\app.ico"
$TeamsCC.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#8f8dfe")

$ButtonClear                     = New-Object system.Windows.Forms.Button
$ButtonClear.text                = "Nettoyer le cache Teams"
$ButtonClear.width               = 338
$ButtonClear.height              = 48
$ButtonClear.enabled             = $true
$ButtonClear.location            = New-Object System.Drawing.Point(3,19)
$ButtonClear.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',16,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$ButtonClear.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("#4b53bc")

$checkBoxReOpen                  = New-Object system.Windows.Forms.CheckBox
$checkBoxReOpen.text             = "Réouvrir Teams à la fin du traitement"
$checkBoxReOpen.AutoSize         = $false
$checkBoxReOpen.width            = 313
$checkBoxReOpen.height           = 16
$checkBoxReOpen.location         = New-Object System.Drawing.Point(14,78)
$checkBoxReOpen.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ButtonUninstall                 = New-Object system.Windows.Forms.Button
$ButtonUninstall.text            = "Désinstaller Teams"
$ButtonUninstall.width           = 337
$ButtonUninstall.height          = 36
$ButtonUninstall.enabled         = $true
$ButtonUninstall.location        = New-Object System.Drawing.Point(4,26)
$ButtonUninstall.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$ButtonUninstall.BackColor       = [System.Drawing.ColorTranslator]::FromHtml("#4b53bc")

$Groupbox1                       = New-Object system.Windows.Forms.Groupbox
$Groupbox1.height                = 77
$Groupbox1.width                 = 346
$Groupbox1.text                  = "Désinstallation de l`'application Teams"
$Groupbox1.location              = New-Object System.Drawing.Point(15,152)

$Groupbox2                       = New-Object system.Windows.Forms.Groupbox
$Groupbox2.height                = 105
$Groupbox2.width                 = 344
$Groupbox2.text                  = "Nettoyage du cache de l`'application Teams"
$Groupbox2.location              = New-Object System.Drawing.Point(15,21)

$Groupbox2.controls.AddRange(@($ButtonClear,$checkBoxReOpen))
$Groupbox1.controls.AddRange(@($ButtonUninstall))
$TeamsCC.controls.AddRange(@($Groupbox1,$Groupbox2))


#Sources: https://lazyadmin.nl/powershell/microsoft-teams-uninstall-reinstall-and-cleanup-guide-scripts/
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

function Uninstall-Teams {
    ##Désinstallation de Teams Wide Installer
    Write-Output "Désinstallation de 'Teams Machine-wide Installer'" -ForegroundColor Yellow
    #Si par administrateur, demande d'utilisateur élevé
    if ( !$(New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) ) {
        $CredAdmin = $host.ui.PromptForCredential("Besoin d'un compte administrateur", "Merci de renseigner un compte administrateur.", "", "NetBiosUserName")
        Start-Process -ArgumentList { Get-WmiObject -Class Win32_Product | Where-Object{ $_.Name -eq "Teams Machine-Wide Installer" } | Foreach-Object { $_.Uninstall() } } -Credential $CredAdmin -Wait
    }
    else { Get-WmiObject -Class Win32_Product | Where-Object{ $_.Name -eq "Teams Machine-Wide Installer" } | Foreach-Object { $_.Uninstall() } }
    #Désinstallation de Teams Utilisateur
    $TeamLocalAppData = Join-path -Path $env:LOCALAPPDATA -ChildPath "\Microsoft\Teams"
    $TeamsProgramData = Join-path -Path "$env:ProgramData\$env:USERNAME" -ChildPath "\Microsoft\Teams"
    if ( Test-Path "$TeamLocalAppData\Current\Teams.exe" ) { $ClientInstaller = Join-path -Path $TeamLocalAppData -ChildPath "\Update.exe" }
    else { 
        if ( Test-Path "$TeamsProgramData\Current\Teams.exe" ) { $ClientInstaller = Join-path -Path $TeamsProgramData -ChildPath "\Update.exe" }
        else { $ClientInstaller = $null }
    }
    if ($ClientInstaller) {
        try { 
            $process = Start-Process -FilePath "$ClientInstaller" -ArgumentList "--uninstall /s" -PassThru -Wait -ErrorAction STOP
            if ($process.ExitCode -ne 0) { Write-Output "La désinstallation a échouée avec le code  $($process.ExitCode)." }
            else {Write-Output "La désinstallation a réussi" }
        }
        catch { Write-Output $_.Exception.Message }
    }
    else { Write-Output "Teams n'a pas été trouvé dans le profil" }
}

function Start-Teams {
    Powershell.exe -Command 'Start-Process "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe"'
    Write-Output  "Lancement de Teams" | Out-File "$env:TEMP\ClearCacheTeams.log" -Append -Encoding utf8
    Start-Sleep -Seconds 5
}

$ButtonClear.Add_Click({ 
    "======= Lancement du nettoyage du cache Teams - $(Get-Date) ======" | Out-File "$env:TEMP\ClearCacheTeams.log" -Encoding utf8
    $ButtonClear.enabled = $false
    $ButtonClear.text = "Nettoyage en cours..."
    Clear-Teams | Out-File "$env:TEMP\ClearCacheTeams.log" -Append -Encoding utf8
    if ($checkBoxReOpen.Checked) { 
        $ButtonClear.text = "Lancement de Teams..."
        Start-Teams
    }
    "======= Processus terminé ======" | Out-File "$env:TEMP\ClearCacheTeams.log" -Append -Encoding utf8
    $ButtonClear.text = "Nettoyage terminé"
})

$ButtonUninstall.Add_Click({ 
    "======= Lancement de la désinstallation de Teams - $(Get-Date) ======" | Out-File "$env:TEMP\ClearCacheTeams.log" -Encoding utf8
    $ButtonClear.enabled = $false
    $ButtonClear.visible = $false
    $ButtonUninstall.enabled = $false
    $ButtonUninstall.text = "Nettoyage en cours..."
    Clear-Teams | Out-File "$env:TEMP\ClearCacheTeams.log" -Append -Encoding utf8
    $ButtonUninstall.text = "Désinstallation en cours..."
    Uninstall-Teams | Out-File "$env:TEMP\ClearCacheTeams.log" -Append -Encoding utf8
    "======= Processus terminé ======" | Out-File "$env:TEMP\ClearCacheTeams.log" -Append -Encoding utf8
    $ButtonUninstall.text = "Désinstallation terminée"
})

[void]$TeamsCC.ShowDialog()
