[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $true)]
    [string]$ContainerAppName,

    [string]$ContainerName,

    [string]$ConfigPath,

    [string]$AzCliPath = "az",

    [switch]$DryRun,

    [switch]$Dev
)

$scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

if ($Dev.IsPresent) {
    $ConfigPath = Join-Path (Join-Path $scriptDirectory "..\config") "settings.dev.json"
} elseif (-not $PSBoundParameters.ContainsKey("ConfigPath") -or [string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path (Join-Path $scriptDirectory "..\config") "settings.json"
}

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Config file not found at '$ConfigPath'."
}

$rawContent = Get-Content -LiteralPath $ConfigPath -Raw
if (-not $rawContent.Trim()) {
    throw "Config file '$ConfigPath' is empty."
}

try {
    $convertParams = @{ InputObject = $rawContent }
    if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey("Depth")) {
        $convertParams["Depth"] = 10
    }
    $config = ConvertFrom-Json @convertParams
} catch {
    throw "Failed to parse JSON from '$ConfigPath': $($_.Exception.Message)"
}

try {
    & $AzCliPath --version | Out-Null
} catch {
    throw "Azure CLI executable '$AzCliPath' could not be invoked. Ensure the CLI is installed and the path is correct."
}

if ($LASTEXITCODE -ne 0) {
    throw "Azure CLI check failed with exit code $LASTEXITCODE."
}

function Convert-ToEnvString {
    param(
        [Parameter(Mandatory = $true)]
        $Value
    )

    switch ($Value) {
        $null { return $null }
        { $_ -is [bool] } { return ($Value.ToString().ToLowerInvariant()) }
        { $_ -is [System.Array] } {
            $flattened = $Value | ForEach-Object { $_.ToString().Trim() } | Where-Object { $_ -ne "" }
            if ($flattened.Count -eq 0) { return "" }
            return ($flattened -join ",")
        }
        default { return ($Value.ToString()) }
    }
}

$quickReplies = $null
if ($config.features -and $config.features.quickReplyOptions) {
    $quickReplies = @($config.features.quickReplyOptions | Where-Object { $_ -and $_.ToString().Trim() -ne "" })
}

$mapping = [ordered]@{
    "SPEECH_RESOURCE_REGION" = $config.speech.region
    "SPEECH_RESOURCE_KEY" = $config.speech.apiKey
    "SPEECH_PRIVATE_ENDPOINT" = $config.speech.privateEndpoint
    "ENABLE_PRIVATE_ENDPOINT" = $config.speech.enablePrivateEndpoint
    "STT_LOCALES" = $config.speech.sttLocales
    "TTS_VOICE" = $config.speech.ttsVoice
    "CUSTOM_VOICE_ENDPOINT_ID" = $config.speech.customVoiceEndpointId
    "AZURE_AI_FOUNDRY_ENDPOINT" = $config.agent.endpoint
    "AZURE_AI_FOUNDRY_AGENT_ID" = $config.agent.agentId
    "AZURE_AI_FOUNDRY_PROJECT_ID" = $config.agent.projectId
    "AGENT_API_URL" = $config.agent.apiUrl
    "SYSTEM_PROMPT" = $config.agent.systemPrompt
    "ENABLE_ON_YOUR_DATA" = $config.search.enabled
    "COG_SEARCH_ENDPOINT" = $config.search.endpoint
    "COG_SEARCH_API_KEY" = $config.search.apiKey
    "COG_SEARCH_INDEX_NAME" = $config.search.indexName
    "AVATAR_CHARACTER" = $config.avatar.character
    "AVATAR_STYLE" = $config.avatar.style
    "AVATAR_CUSTOMIZED" = $config.avatar.customized
    "AVATAR_USE_BUILT_IN_VOICE" = $config.avatar.useBuiltInVoice
    "AVATAR_AUTO_RECONNECT" = $config.avatar.autoReconnect
    "AVATAR_USE_LOCAL_VIDEO_FOR_IDLE" = $config.avatar.useLocalVideoForIdle
    "SHOW_SUBTITLES" = $config.ui.showSubtitles
    "CONTINUOUS_CONVERSATION" = $config.conversation.continuous
    "ENABLE_QUICK_REPLY" = $config.features.quickReplyEnabled
    "QUICK_REPLY_OPTIONS" = if ($quickReplies) { $quickReplies } else { $null }
    "SERVICES_PROXY_PUBLIC_BASE_URL" = $config.servicesProxyBaseUrl
}

$secretKeys = @(
    "SPEECH_RESOURCE_KEY",
    "COG_SEARCH_API_KEY"
)

$envVarArgs = @()
$displayTable = @()

foreach ($key in $mapping.Keys) {
    $originalValue = $mapping[$key]

    if ($null -eq $originalValue) {
        continue
    }

    $rawValue = Convert-ToEnvString -Value $originalValue

    if ($null -eq $rawValue) {
        continue
    }

    $cliValue = if ($rawValue -eq "") { '""' } else { $rawValue }
    $envVarArgs += "$key=$cliValue"

    $maskedValue = if ($secretKeys -contains $key) { "***" } else { $rawValue }
    $displayTable += [pscustomobject]@{ Name = $key; Value = $maskedValue }
}

if ($envVarArgs.Count -eq 0) {
    Write-Warning "No environment variables were found in '$ConfigPath'."
    return
}

Write-Host "Preparing to update Azure Container App environment variables..." -ForegroundColor Cyan
$tableOutput = $displayTable | Format-Table -AutoSize | Out-String
Write-Host $tableOutput

$arguments = @("containerapp", "update", "--name", $ContainerAppName, "--resource-group", $ResourceGroupName)
if ($PSBoundParameters.ContainsKey("ContainerName") -and $ContainerName) {
    $arguments += @("--container", $ContainerName)
}
$arguments += "--set-env-vars"
$arguments += $envVarArgs

if ($DryRun.IsPresent) {
    Write-Host "Dry run enabled. Azure CLI command:" -ForegroundColor Yellow
    Write-Host "`n$AzCliPath $($arguments -join ' ')`n"
    return
}

Write-Host "Invoking Azure CLI to apply environment variables..." -ForegroundColor Cyan
& $AzCliPath @arguments

if ($LASTEXITCODE -ne 0) {
    throw "Azure CLI command failed with exit code $LASTEXITCODE."
}

Write-Host "Environment variables updated successfully." -ForegroundColor Green
