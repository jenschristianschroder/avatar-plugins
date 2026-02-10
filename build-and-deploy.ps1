Param(
    [string]$Registry = "avatarregistry.azurecr.io",
    [string]$ImageName = "avatar-app-hub",
    [string]$Tag = "latest",
    [string]$ResourceGroupName = "avatar-demo",
    [string]$ContainerAppName = "avatar-app-hub",
    [switch]$Dev,
    [switch]$CreateNewRevision,
    [switch]$WaitForRevision,
    [switch]$Local
)

$ErrorActionPreference = 'Stop'

function Invoke-Step {
    param(
        [Parameter(Mandatory)] [scriptblock]$Action,
        [string]$Description
    )

    if ($Description) {
        Write-Host "==> $Description" -ForegroundColor Cyan
    }

    & $Action
}

function Get-NestedConfigValue {
    param(
        [Parameter(Mandatory)] $Config,
        [Parameter(Mandatory)] [string[]]$Path
    )

    $current = $Config
    foreach ($segment in $Path) {
        if ($null -eq $current) {
            return $null
        }

        $hasProperty = $current.PSObject.Properties.Match($segment).Count -gt 0
        if (-not $hasProperty) {
            return $null
        }

        $current = $current.$segment
    }

    return $current
}

function Convert-ConfigValueToEnvString {
    param(
        [Parameter(Mandatory)] $Value
    )

    if ($null -eq $Value) {
        return $null
    }

    switch ($Value.GetType().Name) {
        "Boolean" { return $Value.ToString().ToLowerInvariant() }
        "String" { return $Value.Trim() }
        default {
            if ($Value -is [System.Array]) {
                return (($Value | ForEach-Object { $_.ToString() }) -join "|").Trim()
            }

            return $Value.ToString()
        }
    }
}

function Resolve-DevEnvironmentVariables {
    param(
        [Parameter(Mandatory)] [string]$SettingsPath
    )

    if (-not (Test-Path -LiteralPath $SettingsPath)) {
        throw "Dev settings file not found at '$SettingsPath'."
    }

    $settingsContent = Get-Content -LiteralPath $SettingsPath -Raw
    if (-not $settingsContent.Trim()) {
        throw "Dev settings file '$SettingsPath' is empty."
    }

    $config = $settingsContent | ConvertFrom-Json

    $mappings = @(
        @{ Env = "SPEECH_RESOURCE_REGION"; Path = @("speech", "region") }
        @{ Env = "SPEECH_RESOURCE_KEY"; Path = @("speech", "apiKey") }
        @{ Env = "SPEECH_PRIVATE_ENDPOINT"; Path = @("speech", "privateEndpoint") }
        @{ Env = "ENABLE_PRIVATE_ENDPOINT"; Path = @("speech", "enablePrivateEndpoint") }
        @{ Env = "STT_LOCALES"; Path = @("speech", "sttLocales") }
        @{ Env = "TTS_VOICE"; Path = @("speech", "ttsVoice") }
        @{ Env = "CUSTOM_VOICE_ENDPOINT_ID"; Path = @("speech", "customVoiceEndpointId") }
        @{ Env = "USE_MANAGED_IDENTITY"; Path = @("speech", "useManagedIdentity") }
        @{ Env = "SPEECH_RESOURCE_ID"; Path = @("speech", "speechResourceId") }
        @{ Env = "SPEECH_ENDPOINT"; Path = @("speech", "speechEndpoint") }
        @{ Env = "AZURE_AI_FOUNDRY_ENDPOINT"; Path = @("agent", "endpoint") }
        @{ Env = "AZURE_AI_FOUNDRY_AGENT_ID"; Path = @("agent", "agentId") }
        @{ Env = "AZURE_AI_FOUNDRY_PROJECT_ID"; Path = @("agent", "projectId") }
        @{ Env = "AGENT_API_URL"; Path = @("agent", "apiUrl") }
        @{ Env = "SYSTEM_PROMPT"; Path = @("agent", "systemPrompt") }
        @{ Env = "ENABLE_ON_YOUR_DATA"; Path = @("search", "enabled") }
        @{ Env = "COG_SEARCH_ENDPOINT"; Path = @("search", "endpoint") }
        @{ Env = "COG_SEARCH_API_KEY"; Path = @("search", "apiKey") }
        @{ Env = "COG_SEARCH_INDEX_NAME"; Path = @("search", "indexName") }
        @{ Env = "AVATAR_CHARACTER"; Path = @("avatar", "character") }
        @{ Env = "AVATAR_STYLE"; Path = @("avatar", "style") }
        @{ Env = "AVATAR_CUSTOMIZED"; Path = @("avatar", "customized") }
        @{ Env = "AVATAR_USE_BUILT_IN_VOICE"; Path = @("avatar", "useBuiltInVoice") }
        @{ Env = "AVATAR_AUTO_RECONNECT"; Path = @("avatar", "autoReconnect") }
        @{ Env = "AVATAR_USE_LOCAL_VIDEO_FOR_IDLE"; Path = @("avatar", "useLocalVideoForIdle") }
        @{ Env = "AVATAR_BACKGROUND_IMAGE"; Path = @("avatar", "backgroundImage") }
        @{ Env = "AVATAR_TRANSPARENT_BACKGROUND"; Path = @("avatar", "transparentBackground") }
        @{ Env = "SHOW_SUBTITLES"; Path = @("ui", "showSubtitles") }
        @{ Env = "CONTINUOUS_CONVERSATION"; Path = @("conversation", "continuous") }
        @{ Env = "ENABLE_QUICK_REPLY"; Path = @("features", "quickReplyEnabled") }
        @{ Env = "QUICK_REPLY_OPTIONS"; Path = @("features", "quickReplyOptions") }
        @{ Env = "SERVICES_PROXY_PUBLIC_BASE_URL"; Path = @("servicesProxyBaseUrl") }
        @{ Env = "BRANDING_PRIMARY_COLOR"; Path = @("branding", "primaryColor") }
        @{ Env = "BRANDING_BACKGROUND_IMAGE"; Path = @("branding", "backgroundImage") }
        @{ Env = "BRANDING_LOGO_URL"; Path = @("branding", "logoUrl") }
        @{ Env = "BRANDING_LOGO_ALT"; Path = @("branding", "logoAlt") }
    )

    $envVariables = @{}

    foreach ($mapping in $mappings) {
        $value = Get-NestedConfigValue -Config $config -Path $mapping.Path
        if ($null -eq $value) {
            continue
        }

        $stringValue = Convert-ConfigValueToEnvString -Value $value
        if ([string]::IsNullOrWhiteSpace($stringValue)) {
            continue
        }

        $envVariables[$mapping.Env] = $stringValue
    }

    return $envVariables
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $scriptRoot

try {
    if ($Registry -and $Registry.StartsWith('-')) {
        throw "Invalid registry value '$Registry'. Use '-CreateNewRevision' (single dash) to enable new revision creation."
    }

    # Validate incompatible switch combinations
    if ($Local.IsPresent -and ($CreateNewRevision.IsPresent -or $WaitForRevision.IsPresent)) {
        throw "The -Local switch cannot be used with -CreateNewRevision or -WaitForRevision."
    }

    $timestampTag = Get-Date -Format "yyMMddHHmm"

    if ($Dev.IsPresent) {
        $Tag = $timestampTag
        $ImageName = "avatar-app-dev"
        $ContainerAppName = "avatar-app-dev"
    }

    if ($Local.IsPresent) {
        $Tag = "local-$timestampTag"
        $ImageName = "avatar-app-local"
    }

    $shouldCreateRevision = $CreateNewRevision.IsPresent -or $Dev.IsPresent

    if ($WaitForRevision.IsPresent -and -not $shouldCreateRevision) {
        throw "WaitForRevision can only be used together with CreateNewRevision."
    }

    Invoke-Step -Description "Restoring npm dependencies if needed" -Action {
        if (-not (Test-Path node_modules)) {
            npm install
        }
    }

    Invoke-Step -Description "Building front-end bundle with webpack" -Action {
        npx webpack --config .\webpack.config.js
    }

    # For local builds, don't use registry prefix
    if ($Local.IsPresent) {
        $imageTag = "${ImageName}:${Tag}"
    } else {
        $imageTag = "$Registry/${ImageName}:${Tag}"
    }

    $dockerBuildArgs = @("build", "-t", $imageTag, "-f", ".\Dockerfile")
    if ($Dev.IsPresent) {
        $dockerBuildArgs += @("--build-arg", "CONFIG_FILE=settings.dev.json")
    } elseif ($Local.IsPresent) {
        $dockerBuildArgs += @("--build-arg", "CONFIG_FILE=settings.local.json")
    }
    $dockerBuildArgs += "."

    Invoke-Step -Description "Building container $imageTag" -Action {
        & docker @dockerBuildArgs

        if ($LASTEXITCODE -ne 0) {
            throw "Docker build failed with exit code $LASTEXITCODE."
        }
    }

    if ($Local.IsPresent) {
        # Stop and remove any existing container with the same name
        $containerName = "avatar-local"
        
        Invoke-Step -Description "Cleaning up existing local container if present" -Action {
            $existingContainer = docker ps -a --filter "name=$containerName" --format "{{.Names}}" 2>$null
            if ($existingContainer -eq $containerName) {
                Write-Host "Stopping and removing existing container '$containerName'..." -ForegroundColor Yellow
                docker stop $containerName 2>$null | Out-Null
                docker rm $containerName 2>$null | Out-Null
            }
        }

        # Read environment variables from settings.local.json
        $localSettingsPath = Join-Path (Join-Path $scriptRoot "config") "settings.local.json"
        $envArgs = @()
        
        if (Test-Path $localSettingsPath) {
            Invoke-Step -Description "Loading environment variables from settings.local.json" -Action {
                $envVars = Resolve-DevEnvironmentVariables -SettingsPath $localSettingsPath
                
                if ($envVars.Count -gt 0) {
                    foreach ($kvp in $envVars.GetEnumerator() | Sort-Object Key) {
                        $envArgs += "-e"
                        $envArgs += "$($kvp.Key)=$($kvp.Value)"
                    }
                    Write-Host "Loaded $($envVars.Count) environment variables" -ForegroundColor Green
                }
            }
        } else {
            Write-Warning "settings.local.json not found at $localSettingsPath. Running with default configuration."
        }

        # Run the container locally
        $dockerRunArgs = @(
            "run", "-d",
            "--name", $containerName,
            "-p", "8080:8080",
            "-p", "4000:4000"
        )
        
        if ($envArgs.Count -gt 0) {
            $dockerRunArgs += $envArgs
        }
        
        $dockerRunArgs += $imageTag

        Invoke-Step -Description "Running container locally on ports 8080 (web) and 4000 (agent)" -Action {
            $containerId = & docker @dockerRunArgs

            if ($LASTEXITCODE -ne 0) {
                throw "Docker run failed with exit code $LASTEXITCODE."
            }

            Write-Host "`nContainer started successfully!" -ForegroundColor Green
            Write-Host "  Container ID: $containerId" -ForegroundColor Cyan
            Write-Host "  Web interface: http://localhost:8080" -ForegroundColor Cyan
            Write-Host "  Agent proxy: http://localhost:4000" -ForegroundColor Cyan
            Write-Host "`nTo view logs: docker logs -f $containerName" -ForegroundColor Yellow
            Write-Host "To stop: docker stop $containerName" -ForegroundColor Yellow
        }
    } else {
        Invoke-Step -Description "Pushing $imageTag" -Action {
            docker push $imageTag

            if ($LASTEXITCODE -ne 0) {
                throw "Docker push failed with exit code $LASTEXITCODE."
            }
        }

        if ($shouldCreateRevision) {
            if (-not $ResourceGroupName) {
                throw "ResourceGroupName is required when creating a new revision."
            }

            if (-not $ContainerAppName) {
                throw "ContainerAppName is required when creating a new revision."
            }

            $revisionSuffix = if ($Dev.IsPresent) { $Tag } else { Get-Date -Format "ddMMyyHHmm" }
            $revisionArgs = @(
                "containerapp", "update",
                "--name", $ContainerAppName,
                "--resource-group", $ResourceGroupName,
                "--image", $imageTag,
                "--revision-suffix", $revisionSuffix
            )

            if ($Dev.IsPresent) {
                $devSettingsPath = Join-Path (Join-Path $scriptRoot "config") "settings.dev.json"
                $envVars = Resolve-DevEnvironmentVariables -SettingsPath $devSettingsPath

                if ($envVars.Count -gt 0) {
                    $revisionArgs += "--set-env-vars"
                    $revisionArgs += ($envVars.GetEnumerator() | Sort-Object Key | ForEach-Object { "{0}={1}" -f $_.Key, $_.Value })
                } else {
                    Write-Warning "No environment variables were resolved from $devSettingsPath."
                }
            }

            Invoke-Step -Description "Creating new container app revision with suffix $revisionSuffix" -Action {
                & az @revisionArgs

                if ($LASTEXITCODE -ne 0) {
                    throw "Azure CLI command failed with exit code $LASTEXITCODE."
                }
            }

            if ($WaitForRevision.IsPresent) {
            Invoke-Step -Description "Waiting for revision ${ContainerAppName}--${revisionSuffix} to become active" -Action {
                $revisionName = "$ContainerAppName--$revisionSuffix"
                $timeoutSeconds = 600
                $pollIntervalSeconds = 10
                $deadline = (Get-Date).AddSeconds($timeoutSeconds)
                $showArgs = @(
                    "containerapp", "revision", "show",
                    "--name", $ContainerAppName,
                    "--resource-group", $ResourceGroupName,
                    "--revision", $revisionName,
                    "--only-show-errors"
                )

                while ($true) {
                    $resultJson = & az @showArgs

                    if ($LASTEXITCODE -ne 0) {
                        throw "Azure CLI command failed with exit code $LASTEXITCODE while polling revision status."
                    }

                    try {
                        $revision = $resultJson | ConvertFrom-Json
                    } catch {
                        throw "Failed to parse revision status response: $($_.Exception.Message)"
                    }

                    $provisioningState = $revision.properties.provisioningState
                    $runningState = $revision.properties.runningState
                    $healthState = if ($revision.properties.PSObject.Properties.Name -contains 'healthState') { $revision.properties.healthState } else { $revision.healthState }

                    $activeFlag = $revision.active
                    if ($null -eq $activeFlag -and $revision.properties.PSObject.Properties.Name -contains 'active') {
                        $activeFlag = $revision.properties.active
                    }
                    $isActive = [bool]$activeFlag

                    $hasSucceeded = $provisioningState -eq "Succeeded"
                    $isRunning = $runningState -eq "Running"
                    $isHealthy = $healthState -eq "Healthy"

                    if ($isActive -and ($hasSucceeded -or $isRunning -or $isHealthy)) {
                        Write-Host "Revision $revisionName is active (provisioning=$provisioningState, running=$runningState, health=$healthState)." -ForegroundColor Green
                        return
                    }

                    if ((Get-Date) -ge $deadline) {
                        throw "Revision $revisionName did not become active within $timeoutSeconds seconds. Last known state: provisioning=$provisioningState, running=$runningState, health=$healthState, active=$isActive."
                    }

                    Write-Host "Current state: provisioning=$provisioningState, running=$runningState, health=$healthState, active=$isActive. Checking again in $pollIntervalSeconds seconds..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $pollIntervalSeconds
                }
            }
        }
        }
    }

    Write-Host "Build and deployment steps completed." -ForegroundColor Green
} finally {
    Pop-Location
}
