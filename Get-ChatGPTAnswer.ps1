[CmdletBinding()]
param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]
    $Question,
 
    [Parameter()]
    [Switch]
    $ChangeAPIKey
)
# Import / Install Credential Manager Module
try {
    Import-Module CredentialManager -ErrorAction Stop
}
catch {
    Write-Host "Credential Manager Module not installed, installing..."
    Install-Module CredentialManager -Force -Scope CurrentUser
    Import-Module CredentialManager
}

# Switch ChangeAPIKey
if ($ChangeAPIKey -eq $true) {
    Write-Host "Please Provide an OpenAI API Key. you can generate one at: https://platform.openai.com/account/api-keys "
    $NewAPIKey = Read-Host "OpenAI API Key: "
    New-StoredCredential -Target "OpenAI API Key" -UserName "No Username" -Password $NewAPIKey -Persist LocalMachine | Out-Null
    Exit
}

# Get API Key
try {
    # Try to get the key From Microsoft Credential Store
    $OpenAIAPIKey = Get-StoredCredential -Target "OpenAI API Key" -ErrorAction Stop
    if ($null -eq $OpenAIAPIKey) {
        Throw("API Key is null")
    }   
}
catch {
    # No Credetinals Found, asking to provide some
    Write-Host ("NO OPENAI API KEY FOUND! Error: " + $_.exception.message) -ForegroundColor Red
    Write-Host "Please Provide an OpenAI API Key. you can generate one at: https://platform.openai.com/account/api-keys "
    $NewAPIKey = Read-Host "OpenAI API Key: "
    New-StoredCredential -Target "OpenAI API Key" -UserName "No Username" -Password $NewAPIKey -Persist LocalMachine | Out-Null
    $OpenAIAPIKey = Get-StoredCredential -Target "OpenAI API Key"
}

# Convert Secure String API Key to Plain text
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($OpenAIAPIKey.Password)
$OpenAIAPIKeyPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

# Query OpenAI
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/json")
$headers.Add("Authorization", "Bearer $OpenAIAPIKeyPlainText")

$body = "{
`n  `"model`": `"text-davinci-003`",
`n  `"prompt`": `"$Question`",
`n  `"temperature`": 0,
`n  `"max_tokens`": 2000,
`n  `"top_p`": 1,
`n  `"frequency_penalty`": 0,
`n  `"presence_penalty`": 0
`n}
`n"

try {
    $response = Invoke-RestMethod 'https://api.openai.com/v1/completions' -Method 'POST' -Headers $headers -Body $body
    $response.choices[0].text
}
catch {
    Write-Host ("Error during Reuest: " + $_.exception.message)
}