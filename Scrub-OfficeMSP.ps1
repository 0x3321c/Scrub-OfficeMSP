<#
.SYNOPSIS
This script uninstalls the MS Office MSP.
.DESCRIPTION
Use this script to remove silently the MS Office update patch for MSI-based installer from a computer. Run this PowerShell script for a specific KB and Microsoft Office version (from 2010 to 2016).
#>

function Remove-officeMsp{

$KBtargeted = "KB01234567"
Write-Output $KBtargeted

$Producttargeted = "_Office10.*_"
Write-Output $Producttargeted

$installKeys = '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', '\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'


foreach ($intKey in $installKeys){

$Numberinstalled=0

$KBinfo = @()

Get-childitem HKLM:$intKey | ? { $_.Name -like "*$Producttargeted*"} | % {

    $regKeys = Get-ItemProperty registry::$_ 
    $displayname = $regKeys.DisplayName
    $uninstallstring = ''

    if ($regKeys | select -Property UninstallString) { $uninstallstring = $regKeys.UninstallString }
    
    $tab = ''
      
    if ($uninstallstring.contains("{")) {$tab = $uninstallstring.split("{")}
   
    if ($tab.count -eq 3 ) {

        $product = $tab[1].split("}")[0]
        $patch = $tab[2].split("}")[0]

         $item = New-Object -type pscustomobject -Property @{
                    "DisplayName" = $displayname
                    "Product" = "{" + $product + "}"
                    "Patch" = "{" + $patch + "}"
                    }
        
        $KBinfo += $item  
    }
}


$KBinfo | ? { $_.DisplayName -like "*$KBtargeted*" } | % {

    #$_.product
    #$_.Patch
    
   
   "msiexec /package"  + $_.product +  " /uninstall " +  $_.Patch

   
  $Numberinstalled=$Numberinstalled+1  
  $cmd =  " /package " + $_.product + " /uninstall " + $_.Patch  + " /passive /norestart /l*v " + "C:\windows\temp\" + $KBtargeted + "_" + $Numberinstalled +"-uninstall.log"
  
  start-process  msiexec -ArgumentList $cmd  -Wait
   

} 

}
}