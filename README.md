# Office-365-Tools
Outils autour des produits d'Office 365

https://github.com/ludovicferra/Office-365-Tools/blob/main/Reset-MSTeamsTools.ps1

![image](https://user-images.githubusercontent.com/57104517/136505747-a87ad8bb-e28a-4610-9c5c-8af6472ae7eb.png)

#### Utilisation live du script, executer ceci depuis une console Powershell : 
```PowerShell
Invoke-WebRequest "https://raw.githubusercontent.com/ludovicferra/Office-365-Tools/main/Reset-MSTeamsTools.ps1" -outfile "$Env:Temp\Reset-MSTeamsTools.ps1"
Unblock-File -Path "$Env:Temp\Reset-MSTeamsTools.ps1" #Permet de débloquer le code non signé sans toucher aux règles par défaut.
. "$Env:Temp\Reset-MSTeamsTools.ps1" -Force
```
