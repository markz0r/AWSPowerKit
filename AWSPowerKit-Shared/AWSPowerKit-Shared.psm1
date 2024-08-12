# Set exit on error
$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'




# Function: Install-SharedDependencies
# Description: This function installs the shared dependencies for the M365PowerKit module.
function Install-SharedDependencies {
    function Get-PSModules {
        $REQUIRED_MODULES = @('ExchangeOnlineManagement')
        $REQUIRED_MODULES | ForEach-Object {
            if (-not (Get-InstalledModule -Name $_)) {
                try {
                    Install-Module -Name $_
                    Write-Debug "$_ module installed successfully"
                }
                catch {
                    Write-Error "Failed to install $_ module"
                }
            }
            else {
                Write-Debug "$_ module already installed"
            }
            try {
                Import-Module -Name $_
                Write-Debug "Loading the $_ module..."
                Write-Debug "$_ module loaded successfully"
            }
            catch {
                Write-Error "Failed to import $_ module"
            }
        }
        Write-Debug ' All required modules imported successfully'
    }
    function Test-PowerShellVersion {
        $MIN_PS_VERSION = (7, 3)
        if ($PSVersionTable.PSVersion.Major -lt $MIN_PS_VERSION[0] -or ($PSVersionTable.PSVersion.Major -eq $MIN_PS_VERSION[0] -and $PSVersionTable.PSVersion.Minor -lt $MIN_PS_VERSION[1])) { Write-Host "Please install PowerShell $($MIN_PS_VERSION[0]).$($MIN_PS_VERSION[1]) or later, see: https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows" -ForegroundColor Red; exit }
    }
    # Check for AWScli in powershell
    function Test-AWSCLI {
        $AWSCLI = Get-Command -Name aws -ErrorAction SilentlyContinue
        if (!$AWSCLI) {
            Write-Host 'AWS CLI not found. Please install AWS CLI, see: https://docs.aws.amazon.com/cli/latest/userguide/install-cliv2.html' -ForegroundColor Red
            exit
        }
    }
    Write-Debug 'Installing required PS modules...'
    Get-PSModules
    Write-Debug 'Required modules installed successfully...'
}

# Function to create a new profile
function New-AWSPowerKitProfile {
    # Ask user to enter the profile name
    $ProfileName = Read-Host 'Enter a profile name:'
    $ProfileName = $ProfileName.ToLower().Trim()
    if (!$ProfileName -or $ProfileName -eq '' -or $ProfileName.Length -gt 100) {
        Write-Error 'Profile name cannot be empty, or more than 100 characters, Please try again.'
        # Load the selected profile or create a new profile
        Write-Debug "Profile name entered: $ProfileName"
        Throw 'Profile name cannot be empty, taken or mor than 100 characters, Please try again.'
    }
    else {
        try {
            aws configure --profile $ProfileName
            Write-Debug "Profile $ProfileName created successfully"
        }
        catch {
            Write-Debug "Error: $($_.Exception.Message)"
            throw "New-AWSPowerKitProfile $ProfileName failed. Exiting."
        }
    }
}
