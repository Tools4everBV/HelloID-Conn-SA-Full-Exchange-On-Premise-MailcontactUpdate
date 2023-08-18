# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Variables configured in form
$DN = $form.selectedcontact.DN
$Alias = $form.alias
$ExternalEmailAddress = $form.externalEmailAddress
$DisplayName = $form.displayname
$Name = $form.Name
$initials = $form.initials
$FirstName = $form.firstname
$LastName = $form.lastname
$HiddenFromAddressListsBoolean = $form.hidefromaddresslist

function Resolve-HTTPError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline
        )]
        [object]$ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            FullyQualifiedErrorId = $ErrorObject.FullyQualifiedErrorId
            MyCommand             = $ErrorObject.InvocationInfo.MyCommand
            RequestUri            = $ErrorObject.TargetObject.RequestUri
            ScriptStackTrace      = $ErrorObject.ScriptStackTrace
            ErrorMessage          = ''
        }
        if ($ErrorObject.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') {
            $httpErrorObj.ErrorMessage = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            $httpErrorObj.ErrorMessage = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
        }
        Write-Output $httpErrorObj
    }
}

function Remove-EmptyValuesFromHashtable {
    param(
        [parameter(Mandatory = $true)][Hashtable]$Hashtable
    )

    $newHashtable = @{}
    foreach ($Key in $Hashtable.Keys) {
        if (-not[String]::IsNullOrEmpty($Hashtable.$Key)) {
            $null = $newHashtable.Add($Key, $Hashtable.$Key)
        }
    }
    
    return $newHashtable
}

<#----- Exchange On-Premises: Start -----#>
# Connect to Exchange
try {
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop 
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber
    Write-Information "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" 
    
    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
catch {
    Write-Error "Error connecting to Exchange using the URI [$exchangeConnectionUri]. Error: $($_.Exception.Message)"
    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Failed to connect to Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}


# Update Mail Contact
try {
    $mailContact = Get-MailContact -Filter "ExternalEmailAddress -eq '$ExternalEmailAddress'"

    Write-Verbose "Updating mail contact '$($DisplayName)' with ExternalEmailAddress '$($ExternalEmailAddress)'"

    $contactUpdateParams = @{
        Identity             = $($mailContact.Identity)
        Name                 = $Name
        DisplayName          = $DisplayName
        FirstName            = $FirstName
        Initials             = $Initials
        LastName             = $LastName
    }

    $mailContact = Set-Contact @contactUpdateParams -ErrorAction Stop    

    Write-Information "Successfully updated mail contact with the following parameters: $($contactUpdateParams|ConvertTo-Json)"
    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Successfully updated mail contact with the following parameters: $($contactUpdateParams|ConvertTo-Json)" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $DisplayName # optional (free format text) 
        TargetIdentifier  = $($mailContact.Identity) # optional (free format text) 
    }
    #send result back
    Write-Information -Tags "Audit" -MessageData $log

    $mailContact = Get-MailContact -Filter "ExternalEmailAddress -eq '$ExternalEmailAddress'"

     $mailContactUpdateParams = @{         
        Identity             = $($mailContact.Identity)   
        Alias                = $Alias
        ExternalEmailAddress = $ExternalEmailAddress            
    }

    if ($HiddenFromAddressListsBoolean -eq 'true') {
        $mailContactUpdateParams.Add('HiddenFromAddressListsEnabled', $true)
    }
    if ($HiddenFromAddressListsBoolean -eq 'false') {
        $mailContactUpdateParams.Add('HiddenFromAddressListsEnabled', $false)
    }

    $mailContact = Set-MailContact @mailContactUpdateParams -ErrorAction Stop        

    Write-Information "Successfully updated mail contact with the following parameters: $($mailContactUpdateParams|ConvertTo-Json)"
    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Successfully updated mail contact with the following parameters: $($mailContactUpdateParams|ConvertTo-Json)" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $DisplayName # optional (free format text) 
        TargetIdentifier  = $($mailContact.Identity) # optional (free format text) 
    }
    #send result back
    Write-Information -Tags "Audit" -MessageData $log
}
catch {
    $ex = $PSItem
    if ( $($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObject = Resolve-HTTPError -Error $ex
        $verboseErrorMessage = $errorObject.ErrorMessage
        $auditErrorMessage = $errorObject.ErrorMessage
    }

    # If error message empty, fall back on $ex.Exception.Message
    if ([String]::IsNullOrEmpty($verboseErrorMessage)) {
        $verboseErrorMessage = $ex.Exception.Message
    }
    if ([String]::IsNullOrEmpty($auditErrorMessage)) {
        $auditErrorMessage = $ex.Exception.Message
    }

    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Error updating mail contact with the following parameters: $($mailContactParams|ConvertTo-Json). Error Message: $auditErrorMessage" # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $([string]$mailContactParams.Name) # optional (free format text) 
        TargetIdentifier  = $([string]$mailContactParams.ExternalEmailAddress) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log

    Write-Verbose "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($verboseErrorMessage)"
    throw "Error updating mail contact with the following parameters: $($mailContactParams|ConvertTo-Json). Error Message: $auditErrorMessage"

    # Clean up error variables
    Remove-Variable 'verboseErrorMessage' -ErrorAction SilentlyContinue
    Remove-Variable 'auditErrorMessage' -ErrorAction SilentlyContinue
}
    

# Disconnect from Exchange
try {
    Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
    Write-Information "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]"     
    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
catch {
    Write-Error "Error disconnecting from Exchange.  Error: $($_.Exception.Message)"
    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Failed to disconnect from Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
<#----- Exchange On-Premises: End -----#>
