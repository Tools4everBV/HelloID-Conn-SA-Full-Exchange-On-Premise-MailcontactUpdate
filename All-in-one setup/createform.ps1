# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

#HelloID variables
#Note: when running this script inside HelloID; portalUrl and API credentials are provided automatically (generate and save API credentials first in your admin panel!)
$portalUrl = "https://CUSTOMER.helloid.com"
$apiKey = "API_KEY"
$apiSecret = "API_SECRET"
$delegatedFormAccessGroupNames = @("") #Only unique names are supported. Groups must exist!
$delegatedFormCategories = @("Exchange On-Premise") #Only unique names are supported. Categories will be created if not exists
$script:debugLogging = $false #Default value: $false. If $true, the HelloID resource GUIDs will be shown in the logging
$script:duplicateForm = $false #Default value: $false. If $true, the HelloID resource names will be changed to import a duplicate Form
$script:duplicateFormSuffix = "_tmp" #the suffix will be added to all HelloID resource names to generate a duplicate form with different resource names

#The following HelloID Global variables are used by this form. No existing HelloID global variables will be overriden only new ones are created.
#NOTE: You can also update the HelloID Global variable values afterwards in the HelloID Admin Portal: https://<CUSTOMER>.helloid.com/admin/variablelibrary
$globalHelloIDVariables = [System.Collections.Generic.List[object]]@();

#Global variable #1 >> ExchangeConnectionUri
$tmpName = @'
ExchangeConnectionUri
'@ 
$tmpValue = @'
http://exchangeserver/powershell
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #2 >> ExchangeAdminUsername
$tmpName = @'
ExchangeAdminUsername
'@ 
$tmpValue = @'
username@domain.com
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #3 >> ExchangeAdminPassword
$tmpName = @'
ExchangeAdminPassword
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True"});


#make sure write-information logging is visual
$InformationPreference = "continue"

# Check for prefilled API Authorization header
if (-not [string]::IsNullOrEmpty($portalApiBasic)) {
    $script:headers = @{"authorization" = $portalApiBasic}
    Write-Information "Using prefilled API credentials"
} else {
    # Create authorization headers with HelloID API key
    $pair = "$apiKey" + ":" + "$apiSecret"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $key = "Basic $base64"
    $script:headers = @{"authorization" = $Key}
    Write-Information "Using manual API credentials"
}

# Check for prefilled PortalBaseURL
if (-not [string]::IsNullOrEmpty($portalBaseUrl)) {
    $script:PortalBaseUrl = $portalBaseUrl
    Write-Information "Using prefilled PortalURL: $script:PortalBaseUrl"
} else {
    $script:PortalBaseUrl = $portalUrl
    Write-Information "Using manual PortalURL: $script:PortalBaseUrl"
}

# Define specific endpoint URI
$script:PortalBaseUrl = $script:PortalBaseUrl.trim("/") + "/"  

# Make sure to reveive an empty array using PowerShell Core
function ConvertFrom-Json-WithEmptyArray([string]$jsonString) {
    # Running in PowerShell Core?
    if($IsCoreCLR -eq $true){
        $r = [Object[]]($jsonString | ConvertFrom-Json -NoEnumerate)
        return ,$r  # Force return value to be an array using a comma
    } else {
        $r = [Object[]]($jsonString | ConvertFrom-Json)
        return ,$r  # Force return value to be an array using a comma
    }
}

function Invoke-HelloIDGlobalVariable {
    param(
        [parameter(Mandatory)][String]$Name,
        [parameter(Mandatory)][String][AllowEmptyString()]$Value,
        [parameter(Mandatory)][String]$Secret
    )

    $Name = $Name + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl + "api/v1/automation/variables/named/$Name")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
    
        if ([string]::IsNullOrEmpty($response.automationVariableGuid)) {
            #Create Variable
            $body = @{
                name     = $Name;
                value    = $Value;
                secret   = $Secret;
                ItemType = 0;
            }    
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl + "api/v1/automation/variable")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $variableGuid = $response.automationVariableGuid

            Write-Information "Variable '$Name' created$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        } else {
            $variableGuid = $response.automationVariableGuid
            Write-Warning "Variable '$Name' already exists$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        }
    } catch {
        Write-Error "Variable '$Name', message: $_"
    }
}

function Invoke-HelloIDAutomationTask {
    param(
        [parameter(Mandatory)][String]$TaskName,
        [parameter(Mandatory)][String]$UseTemplate,
        [parameter(Mandatory)][String]$AutomationContainer,
        [parameter(Mandatory)][String][AllowEmptyString()]$Variables,
        [parameter(Mandatory)][String]$PowershellScript,
        [parameter()][String][AllowEmptyString()]$ObjectGuid,
        [parameter()][String][AllowEmptyString()]$ForceCreateTask,
        [parameter(Mandatory)][Ref]$returnObject
    )
    
    $TaskName = $TaskName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl +"api/v1/automationtasks?search=$TaskName&container=$AutomationContainer")
        $responseRaw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false) 
        $response = $responseRaw | Where-Object -filter {$_.name -eq $TaskName}
    
        if([string]::IsNullOrEmpty($response.automationTaskGuid) -or $ForceCreateTask -eq $true) {
            #Create Task

            $body = @{
                name                = $TaskName;
                useTemplate         = $UseTemplate;
                powerShellScript    = $PowershellScript;
                automationContainer = $AutomationContainer;
                objectGuid          = $ObjectGuid;
                variables           = (ConvertFrom-Json-WithEmptyArray($Variables));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/automationtasks/powershell")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $taskGuid = $response.automationTaskGuid

            Write-Information "Powershell task '$TaskName' created$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        } else {
            #Get TaskGUID
            $taskGuid = $response.automationTaskGuid
            Write-Warning "Powershell task '$TaskName' already exists$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        }
    } catch {
        Write-Error "Powershell task '$TaskName', message: $_"
    }

    $returnObject.Value = $taskGuid
}

function Invoke-HelloIDDatasource {
    param(
        [parameter(Mandatory)][String]$DatasourceName,
        [parameter(Mandatory)][String]$DatasourceType,
        [parameter(Mandatory)][String][AllowEmptyString()]$DatasourceModel,
        [parameter()][String][AllowEmptyString()]$DatasourceStaticValue,
        [parameter()][String][AllowEmptyString()]$DatasourcePsScript,        
        [parameter()][String][AllowEmptyString()]$DatasourceInput,
        [parameter()][String][AllowEmptyString()]$AutomationTaskGuid,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $DatasourceName = $DatasourceName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    $datasourceTypeName = switch($DatasourceType) { 
        "1" { "Native data source"; break} 
        "2" { "Static data source"; break} 
        "3" { "Task data source"; break} 
        "4" { "Powershell data source"; break}
    }
    
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/datasource/named/$DatasourceName")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
      
        if([string]::IsNullOrEmpty($response.dataSourceGUID)) {
            #Create DataSource
            $body = @{
                name               = $DatasourceName;
                type               = $DatasourceType;
                model              = (ConvertFrom-Json-WithEmptyArray($DatasourceModel));
                automationTaskGUID = $AutomationTaskGuid;
                value              = (ConvertFrom-Json-WithEmptyArray($DatasourceStaticValue));
                script             = $DatasourcePsScript;
                input              = (ConvertFrom-Json-WithEmptyArray($DatasourceInput));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
      
            $uri = ($script:PortalBaseUrl +"api/v1/datasource")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
              
            $datasourceGuid = $response.dataSourceGUID
            Write-Information "$datasourceTypeName '$DatasourceName' created$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        } else {
            #Get DatasourceGUID
            $datasourceGuid = $response.dataSourceGUID
            Write-Warning "$datasourceTypeName '$DatasourceName' already exists$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        }
    } catch {
      Write-Error "$datasourceTypeName '$DatasourceName', message: $_"
    }

    $returnObject.Value = $datasourceGuid
}

function Invoke-HelloIDDynamicForm {
    param(
        [parameter(Mandatory)][String]$FormName,
        [parameter(Mandatory)][String]$FormSchema,
        [parameter(Mandatory)][Ref]$returnObject
    )
    
    $FormName = $FormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/forms/$FormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }
    
        if(([string]::IsNullOrEmpty($response.dynamicFormGUID)) -or ($response.isUpdated -eq $true)) {
            #Create Dynamic form
            $body = @{
                Name       = $FormName;
                FormSchema = (ConvertFrom-Json-WithEmptyArray($FormSchema));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/forms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
    
            $formGuid = $response.dynamicFormGUID
            Write-Information "Dynamic form '$formName' created$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        } else {
            $formGuid = $response.dynamicFormGUID
            Write-Warning "Dynamic form '$FormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        }
    } catch {
        Write-Error "Dynamic form '$FormName', message: $_"
    }

    $returnObject.Value = $formGuid
}


function Invoke-HelloIDDelegatedForm {
    param(
        [parameter(Mandatory)][String]$DelegatedFormName,
        [parameter(Mandatory)][String]$DynamicFormGuid,
        [parameter()][Array][AllowEmptyString()]$AccessGroups,
        [parameter()][String][AllowEmptyString()]$Categories,
        [parameter(Mandatory)][String]$UseFaIcon,
        [parameter()][String][AllowEmptyString()]$FaIcon,
        [parameter()][String][AllowEmptyString()]$task,
        [parameter(Mandatory)][Ref]$returnObject
    )
    $delegatedFormCreated = $false
    $DelegatedFormName = $DelegatedFormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$DelegatedFormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }
    
        if([string]::IsNullOrEmpty($response.delegatedFormGUID)) {
            #Create DelegatedForm
            $body = @{
                name            = $DelegatedFormName;
                dynamicFormGUID = $DynamicFormGuid;
                isEnabled       = "True";
                useFaIcon       = $UseFaIcon;
                faIcon          = $FaIcon;
                task            = ConvertFrom-Json -inputObject $task;
            }
            if(-not[String]::IsNullOrEmpty($AccessGroups)) { 
                $body += @{
                    accessGroups    = (ConvertFrom-Json-WithEmptyArray($AccessGroups));
                }
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
    
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Information "Delegated form '$DelegatedFormName' created$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
            $delegatedFormCreated = $true

            $bodyCategories = $Categories
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$delegatedFormGuid/categories")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $bodyCategories
            Write-Information "Delegated form '$DelegatedFormName' updated with categories"
        } else {
            #Get delegatedFormGUID
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Warning "Delegated form '$DelegatedFormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
        }
    } catch {
        Write-Error "Delegated form '$DelegatedFormName', message: $_"
    }

    $returnObject.value.guid = $delegatedFormGuid
    $returnObject.value.created = $delegatedFormCreated
}


<# Begin: HelloID Global Variables #>
foreach ($item in $globalHelloIDVariables) {
	Invoke-HelloIDGlobalVariable -Name $item.name -Value $item.value -Secret $item.secret 
}
<# End: HelloID Global Variables #>


<# Begin: HelloID Data sources #>
<# Begin: DataSource "Exchange-mailcontact-update-check-names" #>
$tmpPsScript = @'
try{
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername,$adminSecurePassword)
    $searchValue = ($dataSource.externalemailaddress).trim()
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
        ResultSize = "Unlimited"
    }
   
    
     $invokecommandParams = @{
        Session = $exchangeSession
        Scriptblock = [scriptblock] { Param ($Params)Get-Recipient @Params}
        ArgumentList = $getMailcontactsParams
    } 

    Write-Information "Successfully connected to Exchange '$ExchangeConnectionUri'"  
    
    $mailBoxes =  Invoke-Command @invokeCommandParams | Where-Object { $_.EmailAddresses -match "$searchValue" } 

    $resultCount = $mailboxes.Identity.Count

        Write-Information "Successfully queried Exchange for email address '$searchValue'. Result count: $resultCount"


        # Return info message if mailaddress is available or already in use
        if($resultCount -eq 0){
            $result = "Email address $emailAddress is free to use"
        }else{
            #$result = "Email address $emailAddress is already in use by another mailbox: $($mailboxes.DisplayName) ($($mailboxes.PrimarySmtpAddress))"
            $result = "Email address $emailAddress is already in use by another mailuser: $($mailboxes.DisplayName) ($($mailboxes.PrimarySmtpAddress)) of type $($mailboxes.RecipientTypeDetails). Mailcontact will be updated!"
        }

        $returnObject = @{Result=$result}
        Write-Output $returnObject
    
    Remove-PSSession($exchangeSession)
  
} catch {
    Write-Error "Error connecting to Exchange using the URI '$exchangeConnectionUri', Message '$($_.Exception.Message)'"
}

'@ 
$tmpModel = @'
[{"key":"Result","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"externalemailaddress","type":0,"options":1}]
'@ 
$dataSourceGuid_1 = [PSCustomObject]@{} 
$dataSourceGuid_1_Name = @'
Exchange-mailcontact-update-check-names
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_1_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_1) 
<# End: DataSource "Exchange-mailcontact-update-check-names" #>

<# Begin: DataSource "Exchange-mailcontact-update-hidefromaddresslist" #>
$tmpPsScript = @'
$status = $datasource.hidefromaddresslist.HiddenFromAddressListsEnabled
$result = $false
if($status -eq 'true'){
    $result = $true
}
$returnObject = @{Result=$result}
Write-Output $returnObject
'@ 
$tmpModel = @'
[{"key":"Result","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"hidefromaddresslist","type":0,"options":1}]
'@ 
$dataSourceGuid_2 = [PSCustomObject]@{} 
$dataSourceGuid_2_Name = @'
Exchange-mailcontact-update-hidefromaddresslist
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_2_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_2) 
<# End: DataSource "Exchange-mailcontact-update-hidefromaddresslist" #>

<# Begin: DataSource "Exchange-mailcontact-update-generate-table-wildcard" #>
$tmpPsScript = @'
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

'@ 
$tmpModel = @'
[{"key":"DisplayName","type":0},{"key":"HiddenFromAddressListsEnabled","type":0},{"key":"Initials","type":0},{"key":"LastName","type":0},{"key":"Name","type":0},{"key":"DN","type":0},{"key":"FirstName","type":0},{"key":"Emailaddress","type":0},{"key":"Alias","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"searchvalue","type":0,"options":1}]
'@ 
$dataSourceGuid_0 = [PSCustomObject]@{} 
$dataSourceGuid_0_Name = @'
Exchange-mailcontact-update-generate-table-wildcard
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_0_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_0) 
<# End: DataSource "Exchange-mailcontact-update-generate-table-wildcard" #>
<# End: HelloID Data sources #>

<# Begin: Dynamic Form "Exchange on-premise - Update Mailcontact" #>
$tmpSchema = @"
[{"label":"Search","fields":[{"key":"searchcontact","templateOptions":{"label":"Search"},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"selectedcontact","templateOptions":{"label":"Select contact","required":true,"grid":{"columns":[{"headerName":"Emailaddress","field":"Emailaddress"},{"headerName":"Display Name","field":"DisplayName"},{"headerName":"Alias","field":"Alias"},{"headerName":"Name","field":"Name"},{"headerName":"First Name","field":"FirstName"},{"headerName":"Initials","field":"Initials"},{"headerName":"Last Name","field":"LastName"}],"height":300,"rowSelection":"single"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_0","input":{"propertyInputs":[{"propertyName":"searchvalue","otherFieldValue":{"otherFieldKey":"searchcontact"}}]}},"useFilter":true,"useDefault":false,"searchPlaceHolder":"Search this data"},"type":"grid","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":true}]},{"label":"Edit","fields":[{"key":"externalEmailAddress","templateOptions":{"label":"External Email Address","required":true,"useDependOn":true,"dependOn":"selectedcontact","dependOnProperty":"Emailaddress"},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"emailUniqueInfo","templateOptions":{"label":"Info","rows":3,"placeholder":"Loading...","useDataSource":true,"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_1","input":{"propertyInputs":[{"propertyName":"externalemailaddress","otherFieldValue":{"otherFieldKey":"externalEmailAddress"}}]}},"displayField":"Result"},"className":"textarea-resize-vert","type":"textarea","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"displayname","templateOptions":{"label":"Display Name","required":true,"minLength":2,"useDependOn":true,"dependOn":"selectedcontact","dependOnProperty":"DisplayName"},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"Name","templateOptions":{"label":"Name","useDependOn":true,"dependOn":"selectedcontact","dependOnProperty":"Name"},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"alias","templateOptions":{"label":"Alias","useDataSource":false,"required":true,"minLength":2,"useDependOn":true,"dependOn":"selectedcontact","dependOnProperty":"Alias"},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"firstname","templateOptions":{"label":"First name","useDependOn":true,"dependOn":"selectedcontact","dependOnProperty":"FirstName"},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"initials","templateOptions":{"label":"Initials","useDependOn":true,"dependOn":"selectedcontact","dependOnProperty":"Initials"},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"lastname","templateOptions":{"label":"Last name","useDependOn":true,"dependOn":"selectedcontact","dependOnProperty":"LastName"},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"hidefromaddresslist","templateOptions":{"label":"Hide from Address Lists","useSwitch":true,"checkboxLabel":"Hide from Address Lists","useDataSource":true,"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_2","input":{"propertyInputs":[{"propertyName":"hidefromaddresslist","otherFieldValue":{"otherFieldKey":"selectedcontact"}}]}},"displayField":"Result"},"type":"boolean","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false}]}]
"@ 

$dynamicFormGuid = [PSCustomObject]@{} 
$dynamicFormName = @'
Exchange on-premise - Update Mailcontact
'@ 
Invoke-HelloIDDynamicForm -FormName $dynamicFormName -FormSchema $tmpSchema  -returnObject ([Ref]$dynamicFormGuid) 
<# END: Dynamic Form #>

<# Begin: Delegated Form Access Groups and Categories #>
$delegatedFormAccessGroupGuids = @()
if(-not[String]::IsNullOrEmpty($delegatedFormAccessGroupNames)){
    foreach($group in $delegatedFormAccessGroupNames) {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/groups/$group")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
            $delegatedFormAccessGroupGuid = $response.groupGuid
            $delegatedFormAccessGroupGuids += $delegatedFormAccessGroupGuid
            
            Write-Information "HelloID (access)group '$group' successfully found$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormAccessGroupGuid })"
        } catch {
            Write-Error "HelloID (access)group '$group', message: $_"
        }
    }
    if($null -ne $delegatedFormAccessGroupGuids){
        $delegatedFormAccessGroupGuids = ($delegatedFormAccessGroupGuids | Select-Object -Unique | ConvertTo-Json -Depth 100 -Compress)
    }
}

$delegatedFormCategoryGuids = @()
foreach($category in $delegatedFormCategories) {
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories/$category")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        $response = $response | Where-Object {$_.name.en -eq $category}
        
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
        
        Write-Information "HelloID Delegated Form category '$category' successfully found$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    } catch {
        Write-Warning "HelloID Delegated Form category '$category' not found"
        $body = @{
            name = @{"en" = $category};
        }
        $body = ConvertTo-Json -InputObject $body -Depth 100

        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories")
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid

        Write-Information "HelloID Delegated Form category '$category' successfully created$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    }
}
$delegatedFormCategoryGuids = (ConvertTo-Json -InputObject $delegatedFormCategoryGuids -Depth 100 -Compress)
<# End: Delegated Form Access Groups and Categories #>

<# Begin: Delegated Form #>
$delegatedFormRef = [PSCustomObject]@{guid = $null; created = $null} 
$delegatedFormName = @'
Exchange on-premise - Update Mailcontact
'@
$tmpTask = @'
{"name":"Exchange on-premise - Update Mailcontact","script":"# Set TLS to accept TLS, TLS 1.1 and TLS 1.2\r\n[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12\r\n\r\n$VerbosePreference = \"SilentlyContinue\"\r\n$InformationPreference = \"Continue\"\r\n$WarningPreference = \"Continue\"\r\n\r\n# Variables configured in form\r\n$DN = $form.selectedcontact.DN\r\n$Alias = $form.alias\r\n$ExternalEmailAddress = $form.externalEmailAddress\r\n$DisplayName = $form.displayname\r\n$Name = $form.Name\r\n$initials = $form.initials\r\n$FirstName = $form.firstname\r\n$LastName = $form.lastname\r\n$HiddenFromAddressListsBoolean = $form.hidefromaddresslist\r\n\r\nfunction Resolve-HTTPError {\r\n    [CmdletBinding()]\r\n    param (\r\n        [Parameter(Mandatory,\r\n            ValueFromPipeline\r\n        )]\r\n        [object]$ErrorObject\r\n    )\r\n    process {\r\n        $httpErrorObj = [PSCustomObject]@{\r\n            FullyQualifiedErrorId = $ErrorObject.FullyQualifiedErrorId\r\n            MyCommand             = $ErrorObject.InvocationInfo.MyCommand\r\n            RequestUri            = $ErrorObject.TargetObject.RequestUri\r\n            ScriptStackTrace      = $ErrorObject.ScriptStackTrace\r\n            ErrorMessage          = \u0027\u0027\r\n        }\r\n        if ($ErrorObject.Exception.GetType().FullName -eq \u0027Microsoft.PowerShell.Commands.HttpResponseException\u0027) {\r\n            $httpErrorObj.ErrorMessage = $ErrorObject.ErrorDetails.Message\r\n        }\r\n        elseif ($ErrorObject.Exception.GetType().FullName -eq \u0027System.Net.WebException\u0027) {\r\n            $httpErrorObj.ErrorMessage = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()\r\n        }\r\n        Write-Output $httpErrorObj\r\n    }\r\n}\r\n\r\nfunction Remove-EmptyValuesFromHashtable {\r\n    param(\r\n        [parameter(Mandatory = $true)][Hashtable]$Hashtable\r\n    )\r\n\r\n    $newHashtable = @{}\r\n    foreach ($Key in $Hashtable.Keys) {\r\n        if (-not[String]::IsNullOrEmpty($Hashtable.$Key)) {\r\n            $null = $newHashtable.Add($Key, $Hashtable.$Key)\r\n        }\r\n    }\r\n    \r\n    return $newHashtable\r\n}\r\n\r\n\u003c#----- Exchange On-Premises: Start -----#\u003e\r\n# Connect to Exchange\r\ntry {\r\n    $adminSecurePassword = ConvertTo-SecureString -String \"$ExchangeAdminPassword\" -AsPlainText -Force\r\n    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)\r\n    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck\r\n    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop \r\n    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber\r\n    Write-Information \"Successfully connected to Exchange using the URI [$exchangeConnectionUri]\" \r\n    \r\n    $Log = @{\r\n        Action            = \"UpdateAccount\" # optional. ENUM (undefined = default) \r\n        System            = \"Exchange On-Premise\" # optional (free format text) \r\n        Message           = \"Successfully connected to Exchange using the URI [$exchangeConnectionUri]\" # required (free format text) \r\n        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $exchangeConnectionUri # optional (free format text) \r\n        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) \r\n    }\r\n    #send result back  \r\n    Write-Information -Tags \"Audit\" -MessageData $log\r\n}\r\ncatch {\r\n    Write-Error \"Error connecting to Exchange using the URI [$exchangeConnectionUri]. Error: $($_.Exception.Message)\"\r\n    $Log = @{\r\n        Action            = \"UpdateAccount\" # optional. ENUM (undefined = default) \r\n        System            = \"Exchange On-Premise\" # optional (free format text) \r\n        Message           = \"Failed to connect to Exchange using the URI [$exchangeConnectionUri].\" # required (free format text) \r\n        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $exchangeConnectionUri # optional (free format text) \r\n        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) \r\n    }\r\n    #send result back  \r\n    Write-Information -Tags \"Audit\" -MessageData $log\r\n}\r\n\r\n\r\n# Update Mail Contact\r\ntry {\r\n    $mailContact = Get-MailContact -Filter \"ExternalEmailAddress -eq \u0027$ExternalEmailAddress\u0027\"\r\n\r\n    Write-Verbose \"Updating mail contact \u0027$($DisplayName)\u0027 with ExternalEmailAddress \u0027$($ExternalEmailAddress)\u0027\"\r\n\r\n    $contactUpdateParams = @{\r\n        Identity             = $($mailContact.Identity)\r\n        Name                 = $Name\r\n        DisplayName          = $DisplayName\r\n        FirstName            = $FirstName\r\n        Initials             = $Initials\r\n        LastName             = $LastName\r\n    }\r\n\r\n    $mailContact = Set-Contact @contactUpdateParams -ErrorAction Stop    \r\n\r\n    Write-Information \"Successfully updated mail contact with the following parameters: $($contactUpdateParams|ConvertTo-Json)\"\r\n    $Log = @{\r\n        Action            = \"UpdateAccount\" # optional. ENUM (undefined = default) \r\n        System            = \"Exchange On-Premise\" # optional (free format text) \r\n        Message           = \"Successfully updated mail contact with the following parameters: $($contactUpdateParams|ConvertTo-Json)\" # required (free format text) \r\n        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $DisplayName # optional (free format text) \r\n        TargetIdentifier  = $($mailContact.Identity) # optional (free format text) \r\n    }\r\n    #send result back\r\n    Write-Information -Tags \"Audit\" -MessageData $log\r\n\r\n    $mailContact = Get-MailContact -Filter \"ExternalEmailAddress -eq \u0027$ExternalEmailAddress\u0027\"\r\n\r\n     $mailContactUpdateParams = @{         \r\n        Identity             = $($mailContact.Identity)   \r\n        Alias                = $Alias\r\n        ExternalEmailAddress = $ExternalEmailAddress            \r\n    }\r\n\r\n    if ($HiddenFromAddressListsBoolean -eq \u0027true\u0027) {\r\n        $mailContactUpdateParams.Add(\u0027HiddenFromAddressListsEnabled\u0027, $true)\r\n    }\r\n    if ($HiddenFromAddressListsBoolean -eq \u0027false\u0027) {\r\n        $mailContactUpdateParams.Add(\u0027HiddenFromAddressListsEnabled\u0027, $false)\r\n    }\r\n\r\n    $mailContact = Set-MailContact @mailContactUpdateParams -ErrorAction Stop        \r\n\r\n    Write-Information \"Successfully updated mail contact with the following parameters: $($mailContactUpdateParams|ConvertTo-Json)\"\r\n    $Log = @{\r\n        Action            = \"UpdateAccount\" # optional. ENUM (undefined = default) \r\n        System            = \"Exchange On-Premise\" # optional (free format text) \r\n        Message           = \"Successfully updated mail contact with the following parameters: $($mailContactUpdateParams|ConvertTo-Json)\" # required (free format text) \r\n        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $DisplayName # optional (free format text) \r\n        TargetIdentifier  = $($mailContact.Identity) # optional (free format text) \r\n    }\r\n    #send result back\r\n    Write-Information -Tags \"Audit\" -MessageData $log\r\n}\r\ncatch {\r\n    $ex = $PSItem\r\n    if ( $($ex.Exception.GetType().FullName -eq \u0027Microsoft.PowerShell.Commands.HttpResponseException\u0027) -or $($ex.Exception.GetType().FullName -eq \u0027System.Net.WebException\u0027)) {\r\n        $errorObject = Resolve-HTTPError -Error $ex\r\n        $verboseErrorMessage = $errorObject.ErrorMessage\r\n        $auditErrorMessage = $errorObject.ErrorMessage\r\n    }\r\n\r\n    # If error message empty, fall back on $ex.Exception.Message\r\n    if ([String]::IsNullOrEmpty($verboseErrorMessage)) {\r\n        $verboseErrorMessage = $ex.Exception.Message\r\n    }\r\n    if ([String]::IsNullOrEmpty($auditErrorMessage)) {\r\n        $auditErrorMessage = $ex.Exception.Message\r\n    }\r\n\r\n    $Log = @{\r\n        Action            = \"UpdateAccount\" # optional. ENUM (undefined = default) \r\n        System            = \"Exchange On-Premise\" # optional (free format text) \r\n        Message           = \"Error updating mail contact with the following parameters: $($mailContactParams|ConvertTo-Json). Error Message: $auditErrorMessage\" # required (free format text) \r\n        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $([string]$mailContactParams.Name) # optional (free format text) \r\n        TargetIdentifier  = $([string]$mailContactParams.ExternalEmailAddress) # optional (free format text) \r\n    }\r\n    #send result back  \r\n    Write-Information -Tags \"Audit\" -MessageData $log\r\n\r\n    Write-Verbose \"Error at Line \u0027$($ex.InvocationInfo.ScriptLineNumber)\u0027: $($ex.InvocationInfo.Line). Error: $($verboseErrorMessage)\"\r\n    throw \"Error updating mail contact with the following parameters: $($mailContactParams|ConvertTo-Json). Error Message: $auditErrorMessage\"\r\n\r\n    # Clean up error variables\r\n    Remove-Variable \u0027verboseErrorMessage\u0027 -ErrorAction SilentlyContinue\r\n    Remove-Variable \u0027auditErrorMessage\u0027 -ErrorAction SilentlyContinue\r\n}\r\n    \r\n\r\n# Disconnect from Exchange\r\ntry {\r\n    Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop\r\n    Write-Information \"Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]\"     \r\n    $Log = @{\r\n        Action            = \"UpdateAccount\" # optional. ENUM (undefined = default) \r\n        System            = \"Exchange On-Premise\" # optional (free format text) \r\n        Message           = \"Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]\" # required (free format text) \r\n        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $exchangeConnectionUri # optional (free format text) \r\n        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) \r\n    }\r\n    #send result back  \r\n    Write-Information -Tags \"Audit\" -MessageData $log\r\n}\r\ncatch {\r\n    Write-Error \"Error disconnecting from Exchange.  Error: $($_.Exception.Message)\"\r\n    $Log = @{\r\n        Action            = \"UpdateAccount\" # optional. ENUM (undefined = default) \r\n        System            = \"Exchange On-Premise\" # optional (free format text) \r\n        Message           = \"Failed to disconnect from Exchange using the URI [$exchangeConnectionUri].\" # required (free format text) \r\n        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $exchangeConnectionUri # optional (free format text) \r\n        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) \r\n    }\r\n    #send result back  \r\n    Write-Information -Tags \"Audit\" -MessageData $log\r\n}\r\n\u003c#----- Exchange On-Premises: End -----#\u003e","runInCloud":false}
'@ 

Invoke-HelloIDDelegatedForm -DelegatedFormName $delegatedFormName -DynamicFormGuid $dynamicFormGuid -AccessGroups $delegatedFormAccessGroupGuids -Categories $delegatedFormCategoryGuids -UseFaIcon "True" -FaIcon "fa fa-pencil-square" -task $tmpTask -returnObject ([Ref]$delegatedFormRef) 
<# End: Delegated Form #>

