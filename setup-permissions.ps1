#Requires -Version 5.1

<#
.SYNOPSIS
    Interactive permission configuration script for Office 365 MCP Server

.DESCRIPTION
    This script helps configure Microsoft Graph API permissions for the Office 365 MCP Server.
    It reads the current configuration from appsettings.json, allows interactive permission
    selection by service, and generates admin consent URLs when needed.

.EXAMPLE
    .\setup-permissions.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

# Define available permissions grouped by service
$PermissionGroups = @{
    "User Profile" = @(
        @{ Scope = "User.Read"; Description = "Read user profile"; AdminRequired = $false }
    )
    "Outlook Mail" = @(
        @{ Scope = "Mail.ReadBasic"; Description = "Read basic mail properties"; AdminRequired = $false }
        @{ Scope = "Mail.Read"; Description = "Read user's mail"; AdminRequired = $false }
        @{ Scope = "Mail.ReadWrite"; Description = "Read and write user's mail"; AdminRequired = $false }
        @{ Scope = "Mail.Send"; Description = "Send mail as the user"; AdminRequired = $false }
    )
    "Outlook Calendar" = @(
        @{ Scope = "Calendars.Read"; Description = "Read user's calendars"; AdminRequired = $false }
        @{ Scope = "Calendars.ReadWrite"; Description = "Read and write user's calendars"; AdminRequired = $false }
    )
    "Microsoft Teams" = @(
        @{ Scope = "Team.ReadBasic.All"; Description = "Read basic team info"; AdminRequired = $true }
        @{ Scope = "Channel.ReadBasic.All"; Description = "Read basic channel info"; AdminRequired = $true }
        @{ Scope = "ChannelMessage.Read.All"; Description = "Read channel messages (enables search)"; AdminRequired = $true }
        @{ Scope = "ChannelMessage.Send"; Description = "Send channel messages"; AdminRequired = $true }
        @{ Scope = "Chat.Read"; Description = "Read user's chats (enables search)"; AdminRequired = $false }
        @{ Scope = "Chat.ReadWrite"; Description = "Read and write user's chats"; AdminRequired = $false }
    )
    "OneNote" = @(
        @{ Scope = "Notes.Read"; Description = "Read user's OneNote notebooks"; AdminRequired = $false }
        @{ Scope = "Notes.ReadWrite"; Description = "Read and write OneNote notebooks"; AdminRequired = $false }
        @{ Scope = "Notes.Create"; Description = "Create new OneNote pages"; AdminRequired = $false }
    )
}

$ConfigPath = Join-Path $PSScriptRoot "src\Office365MCP\appsettings.json"

function Read-Configuration {
    if (-not (Test-Path $ConfigPath)) {
        Write-Error "Configuration file not found: $ConfigPath"
        exit 1
    }

    try {
        $json = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        return $json
    }
    catch {
        Write-Error "Failed to parse configuration file: $_"
        exit 1
    }
}

function Write-Configuration {
    param($Config)

    try {
        $json = $Config | ConvertTo-Json -Depth 10
        Set-Content -Path $ConfigPath -Value $json -Encoding UTF8
        Write-Host "✓ Configuration saved to $ConfigPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to save configuration: $_"
        exit 1
    }
}

function Show-Header {
    Clear-Host
    Write-Host "╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   Office 365 MCP Server - Permission Configuration       ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Show-PermissionGroup {
    param(
        [string]$GroupName,
        [array]$Permissions,
        [array]$CurrentScopes
    )

    Write-Host "  $GroupName" -ForegroundColor Yellow
    Write-Host ("  " + ("-" * 60))

    foreach ($perm in $Permissions) {
        $isSelected = $CurrentScopes -contains $perm.Scope
        $checkbox = if ($isSelected) { "[X]" } else { "[ ]" }
        $adminBadge = if ($perm.AdminRequired) { " (Admin)" } else { "" }

        $color = if ($isSelected) { "Green" } else { "Gray" }
        Write-Host "    $checkbox $($perm.Scope)$adminBadge" -ForegroundColor $color
        Write-Host "         $($perm.Description)" -ForegroundColor DarkGray
    }

    Write-Host ""
}

function Select-Permissions {
    param([array]$CurrentScopes)

    $selectedScopes = [System.Collections.ArrayList]::new()
    $selectedScopes.AddRange($CurrentScopes)

    while ($true) {
        Show-Header

        Write-Host "  Current Permissions:" -ForegroundColor Cyan
        Write-Host ""

        foreach ($group in $PermissionGroups.GetEnumerator() | Sort-Object Name) {
            Show-PermissionGroup -GroupName $group.Key -Permissions $group.Value -CurrentScopes $selectedScopes
        }

        Write-Host "  Options:" -ForegroundColor Cyan
        Write-Host "    [A] Add permission"
        Write-Host "    [R] Remove permission"
        Write-Host "    [S] Save and exit"
        Write-Host "    [Q] Quit without saving"
        Write-Host ""

        $choice = Read-Host "  Select option"

        switch ($choice.ToUpper()) {
            "A" {
                $scope = Read-Host "Enter scope name to add (e.g., Mail.Read)"
                if ($scope -and $selectedScopes -notcontains $scope) {
                    $selectedScopes.Add($scope) | Out-Null
                    Write-Host "  ✓ Added: $scope" -ForegroundColor Green
                    Start-Sleep -Seconds 1
                }
            }
            "R" {
                $scope = Read-Host "Enter scope name to remove"
                if ($scope -and $selectedScopes -contains $scope) {
                    $selectedScopes.Remove($scope) | Out-Null
                    Write-Host "  ✓ Removed: $scope" -ForegroundColor Green
                    Start-Sleep -Seconds 1
                }
            }
            "S" {
                return $selectedScopes.ToArray()
            }
            "Q" {
                Write-Host "  Exiting without saving..." -ForegroundColor Yellow
                exit 0
            }
        }
    }
}

function Get-AdminConsentUrl {
    param(
        [string]$ClientId,
        [array]$Scopes
    )

    $scopeString = ($Scopes | ForEach-Object { $_ }) -join " "
    return "https://login.microsoftonline.com/organizations/v2.0/adminconsent?client_id=$ClientId&scope=$([uri]::EscapeDataString($scopeString))"
}

function Main {
    Show-Header

    Write-Host "  Loading configuration..." -ForegroundColor Cyan
    $config = Read-Configuration

    $clientId = $config.Graph.ClientId
    $currentScopes = $config.Graph.Scopes

    Write-Host "  ✓ Loaded $($currentScopes.Count) current permissions" -ForegroundColor Green
    Write-Host ""
    Start-Sleep -Seconds 1

    $newScopes = Select-Permissions -CurrentScopes $currentScopes

    Show-Header
    Write-Host "  Selected Permissions:" -ForegroundColor Cyan
    Write-Host ""

    $adminRequired = $false
    foreach ($scope in $newScopes | Sort-Object) {
        Write-Host "    • $scope" -ForegroundColor Green

        # Check if admin consent is required
        foreach ($group in $PermissionGroups.Values) {
            $perm = $group | Where-Object { $_.Scope -eq $scope }
            if ($perm -and $perm.AdminRequired) {
                $adminRequired = $true
            }
        }
    }

    Write-Host ""
    Write-Host "  Total: $($newScopes.Count) permissions" -ForegroundColor Cyan

    if ($adminRequired) {
        Write-Host ""
        Write-Host "  ⚠ Some permissions require admin consent" -ForegroundColor Yellow
    }

    Write-Host ""
    $confirm = Read-Host "  Save configuration? (Y/N)"

    if ($confirm -eq "Y" -or $confirm -eq "y") {
        $config.Graph.Scopes = $newScopes
        Write-Configuration -Config $config

        if ($adminRequired) {
            Write-Host ""
            Write-Host "  Admin Consent Required" -ForegroundColor Yellow
            Write-Host "  ═══════════════════════" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  To grant admin consent, have your tenant administrator visit:" -ForegroundColor White
            Write-Host ""

            $consentUrl = Get-AdminConsentUrl -ClientId $clientId -Scopes $newScopes
            Write-Host "  $consentUrl" -ForegroundColor Cyan
            Write-Host ""

            $openBrowser = Read-Host "  Open consent URL in browser? (Y/N)"
            if ($openBrowser -eq "Y" -or $openBrowser -eq "y") {
                Start-Process $consentUrl
            }
        }

        Write-Host ""
        Write-Host "  ✓ Setup complete!" -ForegroundColor Green
    }
    else {
        Write-Host "  Configuration not saved." -ForegroundColor Yellow
    }
}

# Run the script
Main
