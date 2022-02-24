
#The resource URI
$resource = "https://westus2.api.loganalytics.io"
#Your Client ID and Client Secret obtained when registering your WebApp
$clientid = "e7987137-0f5b-46d1-addd-116e5586cee0"
$clientSecret = "MMO7Q~CXh6HUIvmrY.w0J4NCXH6c.Jg2XMLle"
#Your Reply URL configured when registering your WebApp
$redirectUri = "https://localhost"
#Scope
$scope = "Data.Read"
Add-Type -AssemblyName System.Web
#UrlEncode the ClientID and ClientSecret and URL's for special characters
$clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($clientid)
$clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($clientSecret)
$resourceEncoded = [System.Web.HttpUtility]::UrlEncode($resource)
$scopeEncoded = [System.Web.HttpUtility]::UrlEncode($scope)
#Refresh Token Path
$refreshtokenpath = "C:\LogAnalyticsRefresh.token"

#Functions
Function Get-AuthCode {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width = 440; Height = 640 }
    $web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width = 420; Height = 600; Url = ($url -f ($Scope -join "%20")) }
    $DocComp = {
        $Global:uri = $web.Url.AbsoluteUri        
        if ($Global:uri -match "error=[^&]*|code=[^&]*") { $form.Close() }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown( { $form.Activate() })
    $form.ShowDialog() | Out-Null
    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $Global:output = @{ }
    foreach ($key in $queryOutput.Keys) {
        $output["$key"] = $queryOutput[$key]
    }
    $output
}

function Get-AzureAuthN ($resource) {
    # Get Permissions (if the first time, get an AuthCode and Get a Bearer and Refresh Token
    # Get AuthCode
    $url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUri&client_id=$clientID&resource=$resourceEncoded&scope=$scopeEncoded"
    Get-AuthCode
    # Extract Access token from the returned URI
    $regex = '(?<=code=)(.*)(?=&)'
    
    $authCode = ($uri | Select-String -pattern $regex).Matches[0].Value
    Write-Output "Received an authCode, $authCode"
    #get Access Token
    $body = "grant_type=authorization_code&redirect_uri=$redirectUri&client_id=$clientId&client_secret=$clientSecretEncoded&code=$authCode&resource=$resource"
    $Authorization = Invoke-RestMethod https://login.microsoftonline.com/common/oauth2/token `
        -Method Post -ContentType "application/x-www-form-urlencoded" `
        -Body $body `
        -ErrorAction STOP
    Write-Output $Authorization.access_token
    $Global:accesstoken = $Authorization.access_token
    $Global:refreshtoken = $Authorization.refresh_token 
    if ($refreshtoken) { $refreshtoken | Out-File "$($refreshtokenpath)" }

    if ($Authorization.token_type -eq "Bearer" ) {
        Write-Host "You've successfully authenticated to $($resource) with authorization for $($Authorization.scope)"           
    }
    else {
        Write-Host "Check the console for errors. Chances are you provided the incorrect clientID and clientSecret combination for the API Endpoint selected"
    }
}

function Get-NewTokens {
    # We have a previous refresh token. 
    # use it to get a new token
    $refreshtoken = Get-Content "$($refreshtokenpath)"
    # Refresh the token
    #get Access Token
    $body = "grant_type=refresh_token&refresh_token=$refreshtoken&redirect_uri=$redirectUri&client_id=$clientId&client_secret=$clientSecretEncoded"
    $Global:Authorization = Invoke-RestMethod https://login.microsoftonline.com/common/oauth2/token `
        -Method Post -ContentType "application/x-www-form-urlencoded" `
        -Body $body `
        -ErrorAction STOP
    $Global:accesstoken = $Authorization.access_token
    $Global:refreshtoken = $Authorization.refresh_token
    if ($refreshtoken) {
        $refreshtoken | Out-File "$($refreshtokenpath)"    
        Write-Host "Updated tokens" 
        $Authorization    
        $Global:headerParams = @{'Authorization' = "$($Authorization.token_type) $($Authorization.access_token)" }
    }
} 

$logAnalyticsWorkspace = "f8592915-335e-49c6-9b2c-e081fd1a9bbb"
$logAnalyticsBaseURI = "https://westus2.api.loganalytics.io/v1/workspaces"

# Get the Log Analytics Data

$WorkspaceName = 'MIMHybrid'
$WorkspaceResourceGroupName = 'sec-lab'
$Workspace = Get-AzOperationalInsightsWorkspace -ResourceGroupName $WorkspaceResourceGroupName -Name $WorkspaceName
$QueryResults = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query 'MIMHybrid_CL 
| project CreatedTime_t, CommittedTime_t, Type_s, RequestStatus_s, TargetObjectType_s, Operation_s,   
 DisplayName_s, 
Target_DisplayName_s, Target_AccountName_s, Target_DepartmentNumber_s, Target_EmployeeType_s,
Creator_HybridObjectID_g, Creator_DisplayName_s, Creator_AccountName_s, 
Approver_DisplayName_s, Approver_AccountName_s, Reason_s, Justification_s, RequestParameter_s, TimeGenerated, EventId_d  
| where TargetObjectType_s == "Person" and Operation_s == "Create" or Operation_s == "Put" and RequestStatus_s == "Completed"'

# Output Columns for CSV
$headerRow = $null
$headerRow = $QueryResults.tables.columns | Select-Object name
$columnsCount = $headerRow.Count

# Format the Report
$logData = @()
foreach ($row in $QueryResults.tables.rows) {
    $data = new-object PSObject
    for ($i = 0; $i -lt $columnsCount; $i++) {
        $data | add-member -membertype NoteProperty -name $headerRow[$i].name -value $row[$i]
    }
    $logData += $data
    $data = $null
}

$QueryResults.results | Export-CSV -Path "C:\LogFile\Results.csv" -NoTypeInformation

   