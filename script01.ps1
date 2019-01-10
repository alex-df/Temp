# Auteur: AldF
# Date: 08.01.2019

Import-Module ActiveDirectory

#ERASE ALL THIS AND PUT XAML BELOW between the @" "@
$inputXML = @"
<Window x:Class="GetSecurityGroupMembersInfo.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="CONTOSO Share Tool" Height="450" Width="410">
    <Grid>
        <TextBox Name="serviceTB" HorizontalAlignment="Left" Height="23" Margin="91,23,0,0" TextWrapping="Wrap" Text="sdi" VerticalAlignment="Top" Width="291"/>
        <Label Name="labelStd2" Content="Service" HorizontalAlignment="Left" Margin="10,19,0,0" VerticalAlignment="Top" Width="76" Height="27"/>
        <CheckBox Name="writeCB" Content="R,W" HorizontalAlignment="Left" Margin="44,96,0,0" VerticalAlignment="Top"/>
        <TextBox Name="folderTB" HorizontalAlignment="Left" Height="23" Margin="91,50,0,0" TextWrapping="Wrap" Text="TestFolder" VerticalAlignment="Top" Width="291"/>
        <Label Name="labelStd1" Content="Répertoire" HorizontalAlignment="Left" Margin="10,46,0,0" VerticalAlignment="Top" Width="76" Height="27"/>
        <Button Name="checkBTN" Content="Check" HorizontalAlignment="Left" Margin="161,96,0,0" VerticalAlignment="Top" Width="221"/>
        <Button Name="createBTN" Content="Créer" HorizontalAlignment="Left" Margin="10,361,0,0" VerticalAlignment="Top" Width="75"/>

        <Rectangle Fill="#FFF4F4F5" HorizontalAlignment="Left" Height="217" Margin="10,128,0,0" Stroke="Black" VerticalAlignment="Top" Width="372"/>
        <Label Name="labelTest" Content="" HorizontalAlignment="Left" Margin="11,128,0,0" VerticalAlignment="Top" Width="372" Height="217"/>
    </Grid>
</Window>

"@       
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
 
#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================
 
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{
    if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
    write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
    get-variable WPF*
}
 
Get-FormVariables
 
#===========================================================================
# Actually make the objects work
#===========================================================================

Import-Module Adaxes

# Fonction pour retourner le chemin UNC du service
function getGPath($service){
    
    $group = "$service`_InformationsTechniques"

    $groupInfo = get-admgroup -Identity $group -Properties extensionAttribute2
    
    $attr2 = $groupInfo.extensionAttribute2.split(";")[2]

    $chemin = $attr2.split(",")[1]
    
    return $chemin
    
}

function createRGroup-OLD{

    #$WPFLabelTest.Content = "Patientez, création des groupes et des ACL..."

    # Création du groupe AD
    New-ADGroup -name "$($WPFserviceTB.Text)£groupe£$($WPFfolderTB.Text)£R" -path $ouData -GroupScope Global -ErrorAction Stop

    # Modification de la description
    $description = "G:\$($WPFfolderTB.Text) en lecture"
    Set-ADGroup -identity "$($WPFserviceTB.Text)£groupe£$($WPFfolderTB.Text)£R" -Description $description
    
    sleep(10)

    # Set des ACL
    $path = getGPath $WPFserviceTB.Text
    $Acl = Get-ACL -Path $path
    $ADObject = New-Object System.Security.Principal.NTAccount("CONTOSO\$($WPFserviceTB.Text)£groupe£$($WPFfolderTB.Text)£R")
    $permissions = 'ReadAndExecute, Synchronize'
    $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule($ADObject, $permissions, $InheritanceDefaultFlag, $PropagationFlag, $objType)
    $Acl.AddAccessRule($objACE)

    ################# CRASH LE SCRIPT ??????????????????????????????
    Set-Acl -path $path -AclObject $Acl

}

function createRWGroup-OLD{
    
    #$WPFLabelTest.Content = "Patientez, création des groupes et des ACL..."

    # Création du groupe AD
    New-ADGroup -name "$($WPFserviceTB.Text)£groupe£$($WPFfolderTB.Text)£RW" -path $ouData -GroupScope Global -ErrorAction Stop 

    # Modification de la description
    $description = "G:\$($WPFfolderTB.Text) en écriture"
    Set-ADGroup -identity "$($WPFserviceTB.Text)£groupe£$($WPFfolderTB.Text)£RW" -Description $description

    # Set des ACL
    $path = getGPath $WPFserviceTB.Text
    $Acl = Get-ACL -Path $path
    $ADObject = New-Object System.Security.Principal.NTAccount("CONTOSO\$($WPFserviceTB.Text)£groupe£$($WPFfolderTB.Text)£RW")
    $permissions = 'DeleteSubdirectoriesAndFiles, Write, ReadAndExecute, Synchronize'
    $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule($ADObject, $permissions, $InheritanceDefaultFlag, $PropagationFlag, $objType)
    $Acl.AddAccessRule($objACE)

    ################# CRASH LE SCRIPT ??????????????????????????????
    Set-Acl $path $Acl

}

function Create-RGroup{

    param (
        [Parameter(Mandatory = $true)]
        $GroupName,
        [Parameter(Mandatory = $true)]
        $Domain,
        [Parameter(Mandatory = $true)]
        $FolderPath,
        [Parameter(Mandatory = $true)]
        $Description,
        [Parameter(Mandatory = $true)]
        $OUPath
    )

    
    $ErrorActionPreference = 'Stop'
    <#$domain = "CONTOSO"
    $folderpath = getGPath $WPFserviceTB.Text
    $OUPath = $ouData
    $groupname = "$($WPFserviceTB.Text)£groupe£$($WPFfolderTB.Text)£R"
    $description = "G:\$($WPFfolderTB.Text) en lecture"
    #>

    Try{
        
        New-ADGroup -name $GroupName -Path $OUPath -GroupScope Global -Description $description
        
        $acl = Get-Acl -Path $FolderPath
        
        $ntAccount = New-Object System.Security.Principal.NTAccount("$Domain\$GroupName")
        $ace = New-Object System.Security.AccessControl.FileSystemAccessRule($ntAccount, 'ReadAndExecute,Synchronize', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $acl.AddAccessRule($ace)
        
        Set-Acl -path $FolderPath -AclObject $acl
    }
    Catch {
        Throw $_
    }
    
}

function Create-RWGroup{

    param (
        [Parameter(Mandatory = $true)]
        $GroupName,
        [Parameter(Mandatory = $true)]
        $Domain,
        [Parameter(Mandatory = $true)]
        $FolderPath,
        [Parameter(Mandatory = $true)]
        $Description,
        [Parameter(Mandatory = $true)]
        $OUPath
    )

    
    $ErrorActionPreference = 'Stop'

    Try{
        
        New-ADGroup -name $GroupName -Path $OUPath -GroupScope Global -Description $description
        
        $acl = Get-Acl -Path $FolderPath
        
        $ntAccount = New-Object System.Security.Principal.NTAccount("$Domain\$GroupName")
        $ace = New-Object System.Security.AccessControl.FileSystemAccessRule($ntAccount, 'DeleteSubdirectoriesAndFiles, Write, ReadAndExecute, Synchronize', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $acl.AddAccessRule($ace)
        
        Set-Acl -path $FolderPath -AclObject $acl
    }
    Catch {
        Throw $_
    }
    
}

#$folderpath = getGPath $WPFserviceTB.Text

#Create-RGroup -GroupName "sdi£groupe£testfolder£R" -Domain CONTOSO -FolderPath "\\ua-sdi\VOL1\sdi\groupe\TestFolder\" -Description "L:\TestFolder W access" -OUPath "OU=Data,OU=Groups,OU=CONTOSO,DC=CONTOSO,DC=TEMP,DC=ch"




# liste des variables WPF
<#
serviceTB
readCB
writeCB
folderTB
checkBTN
createBTN
logRTB
labelTest
#>

$ouData = 'OU=Data,OU=Groups,OU=CONTOSO,DC=CONTOSO,DC=TEMP,DC=ch'
$InheritanceDefaultFlag = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
$PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None
$objType =[System.Security.AccessControl.AccessControlType]::Allow

$WPFcheckBTN.Add_Click({

    $WPFLabelTest.Content = ""
    $action = ""

    Try{
        $path = getGPath $WPFserviceTB.Text
        
        if($WPFwriteCB.IsChecked -eq $true){
            $action = "lecture+écriture"
        }
        else{
            $action = "lecture uniquement"
        }

        $content = "Créer le dossier $($WPFfolderTB.Text) dans `n$path `nCréer le/les security group suivants `n$($WPFserviceTB.Text)£groupe£$($WPFfolderTB.Text)£ -> $action"
    }Catch{
        $content = "Erreur"
    }
    $WPFlabelTest.Content = $content

})

$WPFcreateBTN.Add_Click({

    $WPFLabelTest.Content = ""

    $path = getGPath $WPFserviceTB.Text
    $folderpath = "$path\$($WPFfolderTB.Text)"
    $OUPath = $ouData
    
    

    
    
    if(Test-Path $path){
        
        Try{
            # Créer le dossier
            $WPFlabelTest.Content =  "Création du dossier $path\$($WPFfolderTB.text)"
            New-Item -ItemType directory -path $folderpath -ErrorAction Stop
            $tampon = $WPFLabelTest.Content
            $WPFLabelTest.Content = "$tampon`nok"
            $tampon = $WPFLabelTest.Content

            # Créer les security group
            $WPFLabelTest.Content = "$tampon`nCréation du/des security group et des ACL"
            $tampon = $WPFLabelTest.Content
            if($WPFwriteCB.IsChecked -eq $true){

                $groupname = "$($WPFserviceTB.Text)£groupe£$($WPFfolderTB.Text)£RW"
                $description = "G:\$($WPFfolderTB.Text) en écriture"

                write-host("$groupname,$folderpath,$description,$oudata")

                create-RWGroup -GroupName $groupname -Domain CONTOSO -FolderPath $folderpath -Description $description -OUPath $ouData
                $tampon = $WPFLabelTest.Content
                $WPFLabelTest.Content = "$tampon`n£RW ACL et groupe ok"

                $groupname = "$($WPFserviceTB.Text)£groupe£$($WPFfolderTB.Text)£R"
                $description = "G:\$($WPFfolderTB.Text) en lecture"

                write-host("$groupname,$folderpath,$description,$oudata")
                Create-RGroup -GroupName $groupname -Domain CONTOSO -FolderPath $folderpath -Description $description -OUPath $ouData

                $tampon = $WPFLabelTest.Content
                $WPFLabelTest.Content = "$tampon`n£R ACL et groupe ok"
                               
            }
            else{
                
                $groupname = "$($WPFserviceTB.Text)£groupe£$($WPFfolderTB.Text)£R"
                $description = "G:\$($WPFfolderTB.Text) en lecture"

                write-host("$groupname,$folderpath,$description,$oudata")
                Create-RGroup -GroupName $groupname -Domain CONTOSO -FolderPath $folderpath -Description $description -OUPath $ouData
                
                $tampon = $WPFLabelTest.Content
                $WPFLabelTest.Content = "$tampon`n£R ACL et groupe ok"
                
            }

        }
        Catch [System.IO.IOException]{
            $tampon = $WPFLabelTest.Content
            $WPFLabelTest.Content = "$tampon`nLe dossier existe déjà"
        }
        Catch [Microsoft.ActiveDirectory.Management.ADException]{
            $tampon = $WPFLabelTest.Content
            $WPFLabelTest.Content = "$tampon`nErreur à la création du groupe"
        }
        Catch{
            $tampon = $WPFLabelTest.Content
            $WPFLabelTest.Content = "$tampon`nErreur inconnue"
        }       
        
    } else{
        $tampon = $WPFLabelTest.Content
        $WPFlabelTest.Content = "$tampon`nErreur le chemin n'existe pas"
    }
    

})


#===========================================================================
# Shows the form
#===========================================================================
write-host "To show the form, run the following" -ForegroundColor Cyan
#$Form.ShowDialog() | out-null



$async = $Form.Dispatcher.InvokeAsync({
    $Form.ShowDialog() | out-null
})
$async.Wait() | Out-Null
