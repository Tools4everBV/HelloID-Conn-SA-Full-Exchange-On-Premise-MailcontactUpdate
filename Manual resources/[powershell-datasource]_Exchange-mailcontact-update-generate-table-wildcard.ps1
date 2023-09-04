try{
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername,$adminSecurePassword)
    $searchValue = ($dataSource.searchvalue).trim()
    $searchQuery = "*$searchValue*"  

    $sessionOptionParams = @{
        SkipCACheck = $true
        SkipCNCheck = $true
        SkipRevocationCheck = $true
    }

    $sessionOption = New-PSSessionOption  @SessionOptionParams 

    $sessionParams = @{        
        Authentication = 'Kerberos' 
        ConfigurationName = 'Microsoft.Exchange' 
        ConnectionUri = $ExchangeConnectionUri 
        Credential = $adminCredential        
        SessionOption = $sessionOption       
    }

    $exchangeSession = New-PSSession @SessionParams

    Write-Information "Search query is '$searchQuery'" 
    
    $getMailcontactsParams = @{
        RecipientType = "MailContact"
        ResultSize = "Unlimited"
    }
   
    
     $invokecommandParams = @{
        Session = $exchangeSession
        Scriptblock = [scriptblock] { Param ($Params)Get-Recipient @Params}
        ArgumentList = $getMailcontactsParams
    } 

    Write-Information "Successfully connected to Exchange '$ExchangeConnectionUri'"  
    
    $mailBoxes =  Invoke-Command @invokeCommandParams | Where-Object { $_.EmailAddresses -like "$searchQuery" -or $_.Alias -like "$searchQuery" -or $_.Name -like "$searchQuery" -or $_.DisplayName -like "$searchQuery"} 
    $resultCount = $mailboxes.Identity.Count
    Write-Information "Successfully queried Exchange for contact '$searchValue'. Result count: $resultCount"

    #Write-Information ($mailBoxes[0] | ConvertTo-Json)
    $resultContacts = [System.Collections.Generic.List[PSCustomObject]]::New()

    foreach ($contact in $mailBoxes) {        
        $resultContact = @{            
            DisplayName = $contact.DisplayName
            FirstName = $contact.FirstName
            LastName = $contact.LastName
            Initials = $contact.Initials
            Alias = $contact.Alias
            Name = $contact.Name
            Emailaddress      = $contact.PrimarySMTPAddress
            DN                = $contact.DistinguishedName
            HiddenFromAddressListsEnabled = $contact.HiddenFromAddressListsEnabled



        }
        $resultContacts.add($resultContact)

    }
    $resultContacts    
    
    Remove-PSSession($exchangeSession)
  
} catch {
    Write-Error "Error connecting to Exchange using the URI '$exchangeConnectionUri', Message '$($_.Exception.Message)'"
}

