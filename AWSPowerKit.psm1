<#
.SYNOPSIS
    AWS PowerKit module for interacting with AWS Cloud REST API.
.DESCRIPTION
    AWS PowerKit module for interacting with AWS Cloud REST API.
    - Dependencies: AWSPowerKit-Shared
    - Functions:
      - Use-AWSPowerKit: Interactive function to run any function in the module.
    - Debug output is enabled by default. To disable, set $DisableDebug = $true before running functions.
.EXAMPLE
    Use-AWSPowerKit
    This example lists all functions in the AWSPowerKit module.
.EXAMPLE
    Use-AWSPowerKit
    Simply run the function to see a list of all functions in the module and nested modules.
.EXAMPLE
    Get-DefinedPowerKitVariables
    This example lists all variables defined in the AWSPowerKit module.
.LINK
    GitHub:

#>
$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'
Push-Location $PSScriptRoot

function Get-RequisitePowerKitModules {
    $LOCAL_MODULES = $(Get-ChildItem -Path . -Recurse -Depth 2 -Include *.psd1 -Exclude 'AWSPowerKit.psd1')
    # Find list of module in subdirectories and import them
    $LOCAL_MODULE_LIST = $LOCAL_MODULES | ForEach-Object {
        Write-Debug "Importing nested module: $($_.FullName)"
        Import-Module $_ -Force
        # Validate the module is imported
        if (-not (Get-Module -Name $_.BaseName)) {
            Write-Error "Module $($_.BaseName) not found. Exiting."
            throw "Nested module $($_.BaseName) not found. Exiting."
        }
        return $_.BaseName
    }
    $LOCAL_MODULE_LIST
}

# function to run provided functions with provided parameters (as hash table)
function Invoke-AWSPowerKitFunction {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FunctionName,
        [Parameter(Mandatory = $false)]
        [string]$AWSProfile,
        [Parameter(Mandatory = $false)]
        [hashtable]$Parameters,
        [Parameter(Mandatory = $false)]
        [switch]$SkipNestedModuleImport = $false

    )
    # if AWSProfile is provided, set the profile
    $FUNCTION_PARAMS = @{}
    if ($AWSProfile) {
        $FUNCTION_PARAMS.add('AWS_PROFILE', $AWSProfile)
    }
    elseif ($env:AWSPowerKit_PROFILE_NAME) {
        $FUNCTION_PARAMS.add('AWS_PROFILE', $env:AWSPowerKit_PROFILE_NAME)
    }
    if ($Parameters) {
        # add the @params to the $FUNCTION_PARAMS hash table
        $FUNCTION_PARAMS += $Parameters
    }
    if (! $FUNCTION_PARAMS.AWS_PROFILE) {
        Write-Error 'No AWS profile provided. Exiting...'
        return $null
    }
    if (! $SkipNestedModuleImport) { 
        Write-Debug 'Importing nested modules...'
        Import-NestedModules 
        Write-Debug 'Done: Importing nested modules...'
    }   
    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        # Invoke expression to run the function, splatting the parameters
        $stopwatch.Start()
        Write-Debug "Running function: $FunctionName"
        Write-Debug "Parameters: $($FUNCTION_PARAMS | Out-String)"
        Invoke-Expression "$FunctionName @FUNCTION_PARAMS"
        $stopwatch.Stop()
        Write-Debug "Function: $FunctionName completed in $($stopwatch.Elapsed.TotalSeconds) seconds"
    }
    catch {
        # Write all errors to the console as debug messages
        Write-Debug "Error: $($_.Exception.Message)"
        Write-Error "Failed to run function: $FunctionName"
    }
}

# Function display console interface to run any function in the module
function Show-AWSPowerKitFunctions {
    # List nested modules and their exported functions to the console in a readable format, grouped by module
    $colors = @('Green', 'Cyan', 'Red', 'Magenta', 'Yellow')
    $nestedModules = Get-RequisitePowerKitModules

    $colorIndex = 0
    $functionReferences = @{}
    $nestedModules | ForEach-Object {
        $MODULE = Get-Module -Name $_
        # Select a color from the list
        $color = $colors[$colorIndex % $colors.Count]
        $spaces = ' ' * (52 - $MODULE.Name.Length)
        Write-Host '' -BackgroundColor Black
        Write-Host "Module: $($MODULE.Name)" -BackgroundColor $color -ForegroundColor White -NoNewline
        Write-Host $spaces  -BackgroundColor $color -NoNewline
        Write-Host ' ' -BackgroundColor Black
        $spaces = ' ' * 41
        Write-Host " Exported Commands:$spaces" -BackgroundColor "Dark$color" -ForegroundColor White -NoNewline
        Write-Host ' ' -BackgroundColor Black
        $MODULE.ExportedCommands.Keys | ForEach-Object {
            # Assign a letter reference to the function
            $functRefNum = $colorIndex
            $functionReferences[$functRefNum] = $_

            Write-Host ' ' -NoNewline -BackgroundColor "Dark$color"
            Write-Host '   ' -NoNewline -BackgroundColor Black
            Write-Host "$functRefNum -> " -NoNewline -BackgroundColor Black
            Write-Host "$_" -NoNewline -BackgroundColor Black -ForegroundColor $color
            # Calculate the number of spaces needed to fill the rest of the line
            $spaces = ' ' * (50 - $_.Length)
            Write-Host $spaces -NoNewline -BackgroundColor Black
            Write-Host ' ' -NoNewline -BackgroundColor "Dark$color"
            Write-Host ' ' -BackgroundColor Black
            # Increment the color index for the next function
            $colorIndex++
        }
        $spaces = ' ' * 60
        Write-Host $spaces -BackgroundColor "Dark$color" -NoNewline
        Write-Host ' ' -BackgroundColor Black
    }
    Write-Host 'Note: You can run functions without this interface by calling them directly.' 
    Write-Host "Example: Invoke-AWSPowerKitFunction -FunctionName 'FunctionName' -Parameters @{ 'ParameterName' = 'ParameterValue' }" 
    # Write separator for readability
    Write-Host "`n" -BackgroundColor Black
    Write-Host '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++' -BackgroundColor Black -ForegroundColor DarkGray
    # Ask the user which function they want to run
    $selectedFunction = Read-Host -Prompt "`nSelect a function to run by ID, or FunctionName [parameters] (or hit enter to exit):"
    # if the user enters a function name, run it with the provided parameters as a hash table
    if ($selectedFunction -match '(\w+)\s*\[(.*)\]') {
        $functionName = $matches[1]
        $parameters = $matches[2] -split '\s*,\s*' | ForEach-Object {
            $key, $value = $_ -split '\s*=\s*'
            @{ $key = $value }
        }
        Invoke-AWSPowerKitFunction -FunctionName $functionName -Parameters $parameters -SkipNestedModuleImport
    }
    elseif ($selectedFunction -match '(\d+)') {
        $selectedFunction = [int]$selectedFunction
        Invoke-AWSPowerKitFunction -FunctionName $functionReferences[$selectedFunction] -SkipNestedModuleImport
    }
    elseif ($selectedFunction -eq '') {
        return $null
    }
    else {
        Write-Host 'Invalid selection. Please try again.' -ForegroundColor Red
        Show-AWSPowerKitFunctions
    }
    # Ask the user if they want to run another function
    $runAnother = Read-Host -Prompt 'Run another function? (Y / any key to exit)'
    if ($runAnother -eq 'Y') {
        Show-AWSPowerKitFunctions
    }
    else {
        Write-Host 'Have a great day!'
        return $null
    }
}

# Function to directly invoke a function in the module


# Function to list availble profiles with number references for interactive selection or 'N' to create a new profile
function Show-AWSPowerKitProfileList {
    $profileIndex = 0
    $env:AWSPowerKit_PROFILE_LIST = aws configure list-profiles
    if (!$env:AWSPowerKit_PROFILE_LIST) {
        Write-Host 'No profiles found. Please create a new profile using: aws configure --profile profileName or aws sso configure --profile profileName'
        Write-Debug "Profile List: $(Get-AWSPowerKitProfileList)"
        Show-AWSPowerKitProfileList
    } 
    else {
        Write-Debug "Profile list: $env:AWSPowerKit_PROFILE_LIST"
        $PROFILE_LIST = $env:AWSPowerKit_PROFILE_LIST.split()
        Write-Debug "Profile list array $PROFILE_LIST"
        $PROFILE_LIST | ForEach-Object {
            Write-Host "[$profileIndex] $_"
            $profileIndex++
        }
    }   
    Write-Host '[N] Create a new profile'
    Write-Host '[Q / Return] Quit'
    Write-Host '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++' -ForegroundColor DarkGray
    try {
        # read input from the user and just break with no error if the input is not a number, 'N', 'R' or 'Q'
        $selectedProfile = Read-Host 'Select a profile number or action'
    }
    catch {
        return $null
    }
    if ((!$selectedProfile) -or ($selectedProfile -eq 'Q')) {
        return $null
    }
    elseif ($selectedProfile -eq 'N') {
        New-AWSPowerKitProfile
    } 
    else {
        $selectedProfile = [int]$selectedProfile
        Write-Debug "Selected profile index: $selectedProfile"
        Write-Debug "Selected profile name: $($PROFILE_LIST[$selectedProfile])"
    }
    return "$($PROFILE_LIST[$selectedProfile])"
}

function Use-AWSPowerKit {
    param (
        [Parameter(Mandatory = $false)]
        [string] $ProfileName
    )
    $env:AWSPowerKit_PROFILE_NAME = $null
    Write-Debug 'Running: Get-RequisitePowerKitModules'
    Get-RequisitePowerKitModules
    Write-Debug 'Done: Get-RequisitePowerKitModules'
    #Write-Debug "Profile List: $(Get-AWSPowerKitProfileList)"
    if (!$ProfileName) {
        if (!$FunctionName) {
            Write-Host 'No profile name provided. Check the profiles available.'
            try {
                $env:AWSPowerKit_PROFILE_NAME = Show-AWSPowerKitProfileList
            }
            catch {
                Write-Host 'No profile selected. Exiting...'
                return $null
            }
        }
        else {
            Write-Debug "Example: Use-AWSPowerKit -ProfileName 'profileName' -FunctionName 'functionName'"
            Write-Error 'No -ProfileName provided with FunctionName, Exiting...'
        }
    } 
    else {
        try {
            $ProfileName = $ProfileName.Trim().ToLower()
            Write-Debug "Setting provided profile: $ProfileName"
            $env:AWSPowerKit_PROFILE_NAME = $ProfileName
            if (!$env:AWSPowerKit_PROFILE_NAME -or ($env:AWSPowerKit_PROFILE_NAME -ne $ProfileName)) {
                Throw 'Profile not loaded! Exiting...'
            }
        }
        catch {
            Write-Error "Unable to set profile $ProfileName. Exiting..."
        }
    }
    if ($env:AWSPowerKit_PROFILE_NAME) {
        Write-Host "Profile loaded: $($env:AWSPowerKit_PROFILE_NAME)"
        if (!$FunctionName) {
            Show-AWSPowerKitFunctions
        }
        else {
            try {
                Write-Debug "Running function: $FunctionName"
                Invoke-Expression $FunctionName
            }
            catch {
                Write-Error "Function $FunctionName failed. Exiting..."
            }
        }
        Show-AWSPowerKitFunctions
    }
}
Pop-Location