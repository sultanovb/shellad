
############# Reference to file with variables ##############
. C:\kabi_automata_v-3\O365adVariables.ps1
#############################################################

##########################################################
###############c##### FUNCTIONS ###########################
##########################################################


##########################################################
# function "isUserinAD" check if the user exist in AD KABI
##########################################################

function isUserinAD { 

Clear-Host
write-host "`n*****************************************************`n"
   # Check if file exist and is empty

  

            if (Test-Path $file_user_list_path)
            {
                Write-Host "File exist. Proceeding..."

                            # Check if file is empty

                            if($pathtoUserList -ne $null)
                            {
                                Write-Host "File has user data. Continue...`n"
                            }
                            elseif($pathtoUserList -eq $null){
                    
                                Write-Host "File is empty. Please enter user data `
                                `n Program terminated." -ForegroundColor Yellow
                                exit
                            }

  
              # Checking each line of "kabiUserList.txt"
               foreach($text_line in $pathtoUserList)
            {
                $count_lines++
                # Calling function to check values in the user data string
                check_user_data_string  
            }

    foreach($user in $pathtoUserList){
        
		# Extract email prefix to verify if user exist in AD KABI
        # Split email and take prefix
		[string]$user_to_create_aftercheck = $user
        $user_ad_check = $user.Split("@")[0]
		$user = $user_ad_check
		
		
        Write-Host "`n*****************************************************`n"

    # Return True if user exist False if it does NOT 
        $user_upn = [bool](Get-ADUser -Server $kabi_ad_server -SearchBase $searchThisOU -Filter * -Properties SamAccountName | `
	                                                        select UserPrincipalName | `
															where UserPrincipalName -Like "$user*" | `
															ft -HideTableHeaders | `
															Out-String).Trim()

        # If user EXIST - QUIT script
        # If it DOES - CONTNUE script and create a user in AD KABI
        if ($user_upn -eq $True) {
                        
			Write-Host "`n`n*****************************************************" -ForegroundColor Yellow
			Write-Host "`n`tUser `"$user`" exist in AD KABI" -ForegroundColor Red

            choose_email_type_if_user_exist ($user)
            
            

        }
        elseif ($user_upn -eq $False) {
            Write-Host "`n`n`t`t`tUser $user DOES NOT exist in AD KABI`n `
			Creating new AD KABI user...`n " -ForegroundColor Green
			Write-Host "Please wait..."
			# Call function to create a user account
			. create_ad_user ($user_to_create_aftercheck)
			  Start-Sleep -Seconds 15
            # Call function to move user account to the right OU
             
             
			. choose_email_type ($user_to_create_aftercheck)
			
			
    }
  }
  Write-Host "`n*****************************************************"
  Write-host "Opening file with user`'s password..."
  Start notepad $passfile
  Write-Host "`t`t`t`tProgram completed." -BackgroundColor Green -ForegroundColor Magenta
  Write-Host "*****************************************************`n"
  Clear-Content $pathtoUserList_CLEANUP
  Read-Host "Press any key to exit..."
            continue

}
            else{
                write-host "Ooooops. File does not exist or the path to it is wrong."
            }
}

##########################################################
# function "attrib_mail_cl" to fill up .CL AD attributes
##########################################################

function attrib_mail_cl ($user_to_create_aftercheck) 
{

    # Checking if the user is located in Concepción or Santiago

       
     if($company -ne "CCPU")
     {
     Write-Host "Setting up `"mail`"  and `"mailNickName`" and other AD user attributes for $user_to_create_aftercheck" -ForegroundColor Green
	  $user_ad_set = $user_to_create_aftercheck.Split("@")[0]
      Set-aduser -Identity $user_ad_set -Server $kabi_ad_server -Replace   `
                   @{mail="$($user_ad_set)@$postfixCL"; `
	                    mailNickname="$user_ad_set"; `
                        wWWHomePage = "$wWWHomePage"; `
                        physicalDeliveryOfficeName = $physicalDeliveryOfficeName;`
                        postalCode = $postalCode; `
                        postOfficeBox = $postOfficeBox; `
                        scriptPath = "$scriptPath"; `
                        streetAddress = $streetAddress; `
                        l = $l; `
                        st =$st; `
                        co = "$co"; `
                        c = "$c"; `
                        extensionAttribute1="$attribExtension1"; `
                        company = "$companyOfficial" 
                      }

        }else
        {
            Write-Host "Setting up `"mail`"  and `"mailNickName`" and other AD user attributes for $user_to_create_aftercheck" -ForegroundColor Green
	         $user_ad_set = $user_to_create_aftercheck.Split("@")[0]
             Set-aduser -Identity $user_ad_set -Server $kabi_ad_server -Replace   `
                   @{mail="$($user_ad_set)@$postfixCL"; `
	                    mailNickname="$user_ad_set"; `
                        wWWHomePage = "$wWWHomePage"; `
                        physicalDeliveryOfficeName = $physicalDeliveryOfficeName_ccp;`
                        postalCode = $postalCode_ccp; `
                        postOfficeBox = $postOfficeBox_ccp; `
                        scriptPath = "$scriptPath"; `
                        streetAddress = $streetAddress_ccp; `
                        l = $l_ccp; `
                        st =$st_ccp; `
                        co = "$co"; `
                        c = "$c"; `
                        extensionAttribute1="$attribExtension1"; `
                        company = "$companyOfficial" 
                      }
        }

 	Write-host "`"ZIMBRA - .CL`" - AD KABI attributes were set." -ForegroundColor Green
 
    Write-host "Adding user to the AD groups" -ForegroundColor Green
    $adgroupZimbra, $adgroupOUUsers | Add-ADGroupMember -Members $user_ad_set

}


##########################################################
# function "attrib_mail_com" to fill up .COM AD attributes
##########################################################

function attrib_mail_com ($user_to_create_aftercheck) {
    
    Write-Host "Setting up AD `"O365`" attributes  for $user_to_create_aftercheck" -ForegroundColor Magenta
	$user_ad_set = $user_to_create_aftercheck.Split("@")[0]
    
	$SamAccountName = (Get-ADUser -Identity $user_ad_set -Server $kabi_ad_server -Properties SamAccountName | `
					  select SamAccountName | `
					  Format-Table -HideTableHeaders | `
					  Out-String).Trim()


# Set user attrib and address depending on location

$proxy = "$($user_ad_set)@$postfixCOM"

# Checking if the user is located in Concepción or Santiago
                                                    if ($company -eq  "CCPU")
                                                    {
                                                        Write-Host "User is located in Concepcion. Assigning CCP address..."
                                                        Set-aduser -Identity $sam -Server $kabi_ad_server -Replace `
										                                                        @{ mail="$($user_ad_set)@$postfixCOM"; `                                                                                                      mailNickname="$SamAccountName"; `
                                                                                                      #userPrincipalName="$user_to_create_aftercheck"; 
										                                                              targetAddress="$proxy "; `
										                                                              proxyAddresses="SMTP:$proxy "; `
										                                                              extensionAttribute14="$attribExtension14"; `
										                                                              extensionAttribute15="$attribExtension15"; `
                                                                                                      wWWHomePage = "$wWWHomePage"; `                                                                                                      scriptPath = "$scriptPath"; `                                                                                                      l = $l_ccp; `                                                                                                      st =$st_ccp; `
                                                                                                      postalCode = $postalCode_ccp; `
                                                                                                      postOfficeBox = $postOfficeBox_ccp; `
                                                                                                      streetAddress = $streetAddress_ccp; `
                                                                                                      physicalDeliveryOfficeName = $physicalDeliveryOfficeName_ccp;`
                                                                                                      co = $co; `                                                                                                      c = $c; `
                                                                                                      extensionAttribute1="$attribExtension1"; `
                                                                                                      company = "$companyOfficial"  
                                                                                                    }
                                                         }
                                                     else
                                                     {
                                                               Write-Host "User is located in Santiago."
                                                                if(($company) -eq  ("PUG" -or "PUUA"))
                                                                {
                                                                    Set-aduser -Identity $sam -Server $kabi_ad_server -Replace `
                                                                    @{mail="$($user_ad_set)@$postfixCOM"; `
										                                  mailNickname="$SamAccountName"; `
										                                  targetAddress="$proxy"; `
										                                  proxyAddresses="SMTP:$proxy"; `
										                                  extensionAttribute14="$attribExtension14"; `
										                                  extensionAttribute15="$attribExtension15"; `
                                                                          wWWHomePage = "$wWWHomePage"; `
                                                                          physicalDeliveryOfficeName = "$physicalDeliveryOfficeName"; `
                                                                          postalCode = "$postalCode"; `
                                                                          postOfficeBox = "$postOfficeBox"; `
                                                                          scriptPath = "$scriptPath"; `
                                                                          streetAddress = "$streetAddress";`
                                                                          l = "$l"; `
                                                                          st = "$st"; `
                                                                          co = "$co"; `
                                                                          c = "$c"; ` 
										                                  userPrincipalName="$proxy"; `
                                                                          extensionAttribute1="$attribExtension1"; `  
                                                                          company = "$companyOfficial"          
										                                }
			
                                                                }
                                                                else
                                                                {
                                                                    Set-aduser -Identity $sam -Server $kabi_ad_server -Replace `
										                                                       @{mail="$($user_ad_set)@$postfixCOM"; `
										                                                              mailNickname="$SamAccountName"; `
										                                                              targetAddress="$proxy"; `
										                                                              proxyAddresses="SMTP:$proxy"; `
										                                                              extensionAttribute14="$attribExtension14"; `
										                                                              extensionAttribute15="$attribExtension15"; `
                                                                                                      wWWHomePage = "$wWWHomePage"; `
                                                                                                      physicalDeliveryOfficeName = "$physicalDeliveryOfficeName"; `
                                                                                                      postalCode = "$postalCode"; `
                                                                                                      postOfficeBox = "$postOfficeBox"; `
                                                                                                      scriptPath = "$scriptPath"; `
                                                                                                      streetAddress = "$streetAddress";`
                                                                                                      l = "$l"; `
                                                                                                      st = "$st"; `
                                                                                                      co = "$co"; `
                                                                                                      c = "$c"; ` 
										                                                              userPrincipalName="$proxy"; `
                                                                                                      extensionAttribute1="$attribExtension1"; `
                                                                                                      company = "$companyOfficial"          
										                                                           }                                 
                                                                }

                                                     }
    
    
    Write-host "`"O365 - .COM`" - AD KABI attributes were set." -ForegroundColor Magenta
    Write-host "Adding user to the AD groups" -ForegroundColor Magenta
    $adgroupOUUsers, $adgroupAzure | Add-ADGroupMember -Members $user_ad_set
    
}


##########################################################
# function "choose_email_type" ask .CL or .COM account
##########################################################

function choose_email_type ($user_to_create_aftercheck){
        
    if (($user_to_create_aftercheck) -match $match_cl) {
        Write-Host "Setting up AD attributes --- $user_to_create_aftercheck --- `".CL`"  for ZIMBRA" -ForegroundColor Green
        
        # Call function to create .CL attributes
		Start-Sleep -Seconds 10
        attrib_mail_cl ($user_to_create_aftercheck)


    }
    elseif (($user_to_create_aftercheck) -match $match_com) {
        Write-Host "Setting up KABI AD attributes --- $user_to_create_aftercheck --- `".COM`"  for O365" -ForegroundColor Magenta
        Start-Sleep -Seconds 10
        # Call function to create .COM attributes
        attrib_mail_com ($user_to_create_aftercheck)
    }
   }


##########################################################
# function "create_ad_user" in AD KABI
##########################################################

$pass_file_header + "`n" >> $pass_file

function create_ad_user ($user_to_create_aftercheck){
		
	# Capitalaize first letter of user first and last name	
	$TextInfo = (Get-Culture).TextInfo
   
	# Get user prefix from - first.last@fresenius-kabi.com/cl
	
	$user_id_prefix = $user_to_create_aftercheck.Split("@")[0]
	$user_id_prefix
	# Split prefix and capitalize user first name
	$user_first_name = $user_id_prefix.Split(".")[0]
	$user_first_name = $TextInfo.ToTitleCase("$user_first_name")
	
	# Split prefix and capitalize user last name
	$user_last_name = $user_id_prefix.Split(".")[1]
	$user_last_name = $TextInfo.ToTitleCase("$user_last_name")
	
	
	# Put capitalized first and last user names together
	$user_full_name = $user_first_name + " " + $user_last_name
 	Write-Host "`n`tCreating a new AD user --- $user_full_name ---`n" -ForegroundColor Green
	
	# Create SamAccountName
	$script:sam = $user_first_name + "." + $user_last_name
	
	##Create UserPrincipalName
	$upn = $sam + "@" + $user_upn_postfix

    # Creating Attributes
    # tor.lik@fresenius-kabi.com,MUAF,12345,+56224627066,11234569-7,andres.obaid,IT,Admi.de Sistemas Senior
   
                    
                  $company =  $user_to_create_aftercheck.Split(",")[1]     # "MU/PU ??"
                  $script:attribExtension1 = $user_to_create_aftercheck.Split(",")[2]  # "enter number Cost Center"
                  $telephoneNumber = $user_to_create_aftercheck.Split(",")[3]
                  $employeeNumber =  $user_to_create_aftercheck.Split(",")[4]  # "userRUT"
                  $manager = $user_to_create_aftercheck.Split(",")[5]
                  #$title = $user_to_create_aftercheck.Split(",")[6] # "IT"
                  $department = $user_to_create_aftercheck.Split(",")[6] # "IT"
                  $title = $user_to_create_aftercheck.Split(",")[7]  #"Admi.de Sistemas Senior  ???"
                  $description = $user_to_create_aftercheck.Split(",")[7]  #"Admi.de Sistemas Senior  ???"
                  

                  # Choosing and assigning value to "company" attribute
                  foreach($unit in $belongs_to_company.Keys)
                        {
                            #$belongs_to_company.$unit
                            if($company  -eq $unit)
                            {
                                $unit_assign_to = $belongs_to_company.$unit
                            }
                        }
                   
                  
 
	
	# Create, enable and assign random password for a new user account in specific OU 
	New-ADUser -Name $user_full_name -GivenName $user_first_name -Surname $user_last_name `
	           -Server $kabi_ad_server `
               -ChangePasswordAtLogon $true `
	           -SamAccountName $sam.ToLower() -UserPrincipalName $upn.ToLower()`
			   -AccountPassword (ConvertTo-SecureString -AsPlainText $newPassword -Force ) `
			   -Enabled $true `
			   -Path $specific_ou `
               -DisplayName $user_full_name `
               -Department $department `               -OfficePhone $telephoneNumber `
               -EmployeeNumber $employeeNumber `
               -Manager $manager `
               -Title $title `
               -Description $description `
               -Company $company 

#Start-Sleep -Seconds 5

    #Set-ADUser -Identity $sam -Replace @{extensionAttribute1="$attribExtension1";company="$companyOfficial"}
               
          
                    # Loop goes through hash table and selects the right OU to move user in
                    # Detects if user is in Santiago or Concepcion
                    foreach($ou in $ou_to_move.Keys)
                    {

                        # Chosen OU from the STRING
                       # $company
                        if($company -eq  $ou)
                        {
                            # Assign hash value of selected key from string OU 
                               $company_assign = $ou_to_move.$ou 
                                Write-Host "Moving user to $company"
                            # Moving user....
                                Get-ADUser $sam -Server $kabi_ad_server | Move-ADObject -TargetPath $company_assign

                         }       
                    }
                
	
	# Write username and password into text file 
	Write-Host "Check $user_full_name password in $pass_file"
    Write-Host "Program execution is in progress.... PLEASE WAIT..."
	$user_full_name + " -------> " + $newPassword + "      " +  $timestamp >> $pass_file
	# Reset_user_pass
	$newPassword = "empty now"
    . generate_random_password
    Start-Sleep -Seconds 10
}

##########################################################
# function "choose_email_type_if_user_exist" in AD KABI when the user exist
##########################################################

function choose_email_type_if_user_exist ($user)
{
        
       $user_if_has_email = (Get-ADUser -Identity $user -Server $kabi_ad_server -Properties mail | `
                                                                   select mail | ft -HideTableHeaders | Out-String).Trim() 
      
 if (($user_if_has_email) -match $match_cl) 
    {

        Write-host "User  `"$user`" already has `".CL`" account set" -ForegroundColor Yellow
        $choice = Read-Host "`t`tIf you want to migrate $user to email `".COM`", please enter `"y`" `
        Otherwise enter `"n`" and the program will be terminated "
              
                       switch($choice)
                       {
                                "y"
                                    {
                                       Write-Host "You have chosen to migrate `".CL`"  to `".COM`"" -ForegroundColor Magenta
                                       Write-Host "Proceeding..." -ForegroundColor Magenta

                                       # Calling function to replace .CL with .COMdi
                                       migrate_cl_to_com ($user)

                                       Write-host "`"O365 - .COM`" - AD KABI attributes were set." -ForegroundColor Magenta
                                       Write-Host "*****************************************************"
                                       Write-Host "`t`t`t`tChecking next user...." -BackgroundColor Green -ForegroundColor Magenta
                                       Write-Host "*****************************************************"
                                       #exit
                                       continue 
                                    
                                    }
                                "n"
                                    {
                                        Write-Host "You have chosen to exit the program." -ForegroundColor Yellow
                                        Write-Host "No changes will be done to user account" -ForegroundColor Yellow
                                        Write-Host "*****************************************************" -ForegroundColor Yellow   
                                        Write-Host "`t`t`t`tProgram terminated" -ForegroundColor Red
                                        Write-Host "*****************************************************" -ForegroundColor Yellow
                                        exit
                                    }
                                default
                                    {
                                        Write-Host "Invalid entry. You have NOT entered `"y`"   or `"n`"." -ForegroundColor Yellow
                                        Write-Host "*****************************************************" -ForegroundColor Yellow   
                                        Write-Host "`t`t`t`tProgram terminated" -ForegroundColor Red
                                        Write-Host "*****************************************************" -ForegroundColor Yellow
                                        exit 
                                    }
                       }


        
    }
    elseif (($user_if_has_email) -match $match_com) 
    {
        Write-host "User already has `".COM`" account set" -ForegroundColor Yellow
        Write-Host "`tSkipping the user... NO change done to this account.`n`n" -ForegroundColor Red
        continue
        # exit
    }
    else
    {
        Write-host "User account email settings are not set or incorrect.`nPlease check." -ForegroundColor Yellow
    }
}



##########################################################
# function "migrate_cl_to_com($user)"replace .CL attributes with .COM
##########################################################
function migrate_cl_to_com ($user)
{
    
     # Capitalaize first letter of user first and last name	
	 $TextInfo = (Get-Culture).TextInfo

     $SamAccountName = (Get-ADUser -Identity $user -Server $kabi_ad_server -Properties sAMAccountName | select sAMAccountName | ft -HideTableHeaders | Out-String).Trim()
     $upn = $user + "@" + "$user_upn_postfix"
     
     # If user has only one-word "UPN"




     $user_first_name = $user.Split(".")[0]
     $user_f_name = $TextInfo.ToTitleCase($user_first_name)
     $user_last_name = $user.Split(".")[1] 
     $user_l_name = $TextInfo.ToTitleCase($user_last_name)
     $displayName = $user_f_name + " " + $user_l_name
     
      $user_if_has_email = (Get-ADUser -Identity $user -Server $kabi_ad_server -Properties mail | `
                                                                   select mail | ft -HideTableHeaders | Out-String).Trim() 



     Set-aduser -Identity $user -Server $kabi_ad_server  -Replace @{mail="$($user)@$postfixCOM";mailNickname="$SamAccountName";userPrincipalName="$upn";`
     targetAddress="$($user)@$postfixCOM";displayName=$displayName;proxyAddresses="SMTP:$($user)@$postfixCOM";extensionAttribute14="$attribExtension14";extensionAttribute15="$attribExtension15";extensionAttribute1="$attribExtension1";company="$companyOfficial"}
		

								                                                                                     
}

##########################################################
# function "generate_password()" in AD KABI
##########################################################

function generate_random_password
{
	param([string]$newPassword="")
	
    #########
    function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    return [String]$characters[$random]
}
 function Scramble-String([string]$inputString){     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}

 
$newPassword = Get-RandomCharacters -length 10 -characters 'abcdefghiklmnoprstuvwxyz'
$newPassword += Get-RandomCharacters -length 2 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
$newPassword += Get-RandomCharacters -length 2 -characters '1234567890'
$newPassword += Get-RandomCharacters -length 1 -characters '!"§$%&/()=?}][{@#*+'
 
# Write-Host $newPassword
 
$newPassword = Scramble-String $newPassword
 
#Write-Host $newPassword

}
. generate_random_password	



##########################################################
# function "check_user_data_string" to verify user data string in "KabiUserList.txt" file
##########################################################

function check_user_data_string(){    $error_display =  " ------------------------------------------------------------------------------`                                            ATENCIÓN !!! `n ------------------------------------------------------------------------------ `n`  Line - $count_lines 

 Una arroba - `"@`" o/y coma - `",`" (una o más) que separan los valores fáltan  `
  Por favor verificar datos en `"KabiUserList.txt`" `n `
  Orden de entrada de datos en `"KabiUserList.txt`"  `n  `
    `t1) email `n
    `t2) OU `n
    `t3) cost-center `n
    `t4) phone `n
    `t5) rut `n
    `t6) gerente como - `"UserID`"  `n
    `t7) title `n
    `t8) Descripción del nuevo usuario (título del trabajo)  `n
  Todos los valores DEBEN ESTAR EN UNA LÍNEA separados por comas - `",`"`n
  Ejemplo: `n `
    tor.lik@fresenius-kabi.com,MUAF,12345,+56224627066,11234569-7,andres.obaid,`n    IT,Admi.de Sistemas Senior`n`n---------------------------------------------------------------------------- `n"        $good_message = "User data string is ok, ready to go !!!"        # Count number of "@"  and ","     $commas_num = [regex]::Matches($text_line,",").count    $at_sign_num = [regex]::Matches($text_line,"@").count    if(($commas_num -eq 7) -and ($at_sign_num -eq 1))    {        Write-Host "Line - $count_lines " $good_message -ForegroundColor Green    }else    {        Write-Host   $error_display -ForegroundColor Yellow        Write-Host "                              Program terminated" -ForegroundColor Red        Write-host "`n*****************************************************`n"        exit    }}



##########################################################
# Function Remove-ScriptVariables
##########################################################

function Remove-ScriptVariables($path) {
    $result = Get-Content $path |
    ForEach  {

    if ( $_ -match ‘(\$.*?)\s*=’) {
    $matches[1]  | ? { $_ -notlike ‘*.*’ -and $_ -notmatch ‘result’ -and $_ -notmatch ‘env:’}
    }

    }

    ForEach ($v in ($result | Sort-Object | Get-Unique))  {
    Write-Host “Removing” $v.replace(“$”,””)
    Remove-Variable ($v.replace(“$”,””)) -ErrorAction SilentlyContinue
    }

    }
    # end function Get-ScriptVariables
   # Remove-ScriptVariables($path) 

# generate a new random password
# dot infront of function will make internal function variable available outside
#. generate_random_password

# START HERE !!!!!!!!!!!#Start-Sleep -Seconds 600

isUserinAD

# START HERE !!!!!!!!!!!


##########################################################
#################### END OF FUNCTIONS ####################
##########################################################




# SIG # Begin signature block
# MIID+QYJKoZIhvcNAQcCoIID6jCCA+YCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUZRB0CyCDV1QpCcukdxkBR01n
# ptOgggIWMIICEjCCAX+gAwIBAgIQisNZh1LR/75OMapoFF9I3jAJBgUrDgMCHQUA
# MBkxFzAVBgNVBAMTDkJ1bGF0IFN1bHRhbm92MB4XDTE4MTIyMTEzMTg1NFoXDTM5
# MTIzMTIzNTk1OVowGTEXMBUGA1UEAxMOQnVsYXQgU3VsdGFub3YwgZ8wDQYJKoZI
# hvcNAQEBBQADgY0AMIGJAoGBAPlYhyl6giVH73gxe28+jBXFt368ChET3/eIvNqN
# B9l71z5Re2+Ud/cDPp4oXI6uFOsOfbLznBa68FEVjEd56eRj1XVbw016gQEyjeZn
# fKxAPPsbaUCUOVmUtOPbdpT8Q+K4yoGodju5LQP82xBwE1Q8XOtywmbPcHs9aPYa
# nk5tAgMBAAGjYzBhMBMGA1UdJQQMMAoGCCsGAQUFBwMDMEoGA1UdAQRDMEGAELd+
# 3oTufimsd0/dUpf5TBShGzAZMRcwFQYDVQQDEw5CdWxhdCBTdWx0YW5vdoIQisNZ
# h1LR/75OMapoFF9I3jAJBgUrDgMCHQUAA4GBAIMc++JyM2qxFYVAjm4cTk0Ux8EB
# yo6IGc8/WPmIOjh129rpmrEVH7IosJpvDoZNYRjE6Lq/IZ38KlVJGiQg2jE55uDE
# iS1D3rVsOVLL435SAI2GR1waiwmpkdgnSWwiSDVsy9NYg0SGIAfY4aF54VIL/RkT
# 6l+dOjlGKGg4NoCbMYIBTTCCAUkCAQEwLTAZMRcwFQYDVQQDEw5CdWxhdCBTdWx0
# YW5vdgIQisNZh1LR/75OMapoFF9I3jAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU1QiKiyjZ+jzV
# D1YW1fyVuexqCwIwDQYJKoZIhvcNAQEBBQAEgYC4450GOhfa0n14uo3VQ8BgmFAS
# ymamlDY7Shqb1RY373utJNPHODbGB6ttXzwH9y8xoeEE2Vb036VFM1dnnhIHa+/f
# Be4YDHLxD4m8XvqQt5N3wei+QR9AVfGzphp+q2nNmOtOeAIS0YOWd8W3u8jz6f8e
# ucRklNdRFjAqcuwaTA==
# SIG # End signature block
