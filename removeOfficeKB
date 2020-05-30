$KBtargeted = "KB4018313"

$Producttargeted = "_Office10.PROPLUS_"

$Numberinstalled=0

$KBinfo = @()

Get-childitem HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | ? { $_.Name -like "*$Producttargeted*"} | % {

    $kb = Get-ItemProperty registry::$_ 
    $displayname = $kb.DisplayName
    $uninstallstring = ''

    if ($kb | select -Property UninstallString) { $uninstallstring = $kb.UninstallString }
    
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
  $cmd =  " /package " + $_.product + " /uninstall " + $_.Patch  + " /passive /norestart /l*v " + "C:\windows\logs\PKG\" + $KBtargeted + "_" + $Numberinstalled +"-uninstall.log"
  
  start-process  msiexec -ArgumentList $cmd  -Wait
   

} 

