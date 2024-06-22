# Set exit on error
$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'

function Select-AWSProfile {
    # Show list of AWS profiles configured on the system with numbered selection then prompt for selection and set the $AWS_PROFILE variable to the corresponding profile name, not the profile number
    $AWS_PROFILES = aws configure list-profiles
    $AWS_PROFILES | ForEach-Object -Begin { $i = 1 } -Process { Write-Host "$i. $_"; $i++ }
    $AWS_PROFILE = Read-Host 'Select AWS Profile Number'
    $AWS_PROFILE = $AWS_PROFILES[$AWS_PROFILE - 1]
    Write-Debug "You selected profile $AWS_PROFILE"
    $AWS_PROFILE
}

function Get-Zonelist {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AWS_PROFILE
    )
    $ZONE_LIST = aws route53 list-hosted-zones --profile $AWS_PROFILE
    Write-Debug "ZONE_LIST: $ZONE_LIST"
    $ZONE_LIST_OBJECT = $ZONE_LIST | ConvertFrom-Json
    $ZONE_LIST_OBJECT
}

function Get-Route53Report {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SIMPLE_NAME,
        [Parameter(Mandatory = $false)]
        [string]$AWS_PROFILE,
        [Parameter(Mandatory = $false)]
        [System.Object]$ZONE_LIST
    )
    # Function that creates a new excel file and adds a worksheet
    function Export-ExcelFile {
        param(
            [Parameter(Mandatory = $true)]
            [string]$REPORT_PATH,
            [Parameter(Mandatory = $true)]
            [string]$SheetName,
            [Parameter(Mandatory = $true)]
            [System.Object]$R53_ZONE_LIST,
            [Parameter(Mandatory = $true)]
            [string]$AWS_PROFILE,
            [Parameter(Mandatory = $true)]
            [System.Object]$SIMPLE_NAME
        )
        $PULL_DATE = Get-Date -Format 'yyyyMMdd'
        $R53_ZONE_JSON = aws route53 list-hosted-zones --profile $AWS_PROFILE --output json
        $R53_ZONE_LIST = $R53_ZONE_JSON | ConvertFrom-Json
        # {@{Id=/hostedzone/Z1I5ZMLGJYA2FJ; Name=thenumberingsystem.com.au.; CallerReference=B948C08A-E6A8-78FD-9A69-85B0EAF478E7; Config=; ResourceRecordSetCount=35}, @{Id=/hostedzone/Z1WU1KB5ITOT0S; Name=num.local.; CallerReference=DA0CC1AA-F3B7-98FE-A71C-1D1E55D92129; Config=… 
        # For each hosted zone run aws route53 list-resource-record-sets --hosted-zone-id $_.Id | Out-File .\$AWS_PROFILE-R53Zone-$_.Name-$PULL_DATE -Encoding utf8
        $HOSTED_ZONE_ARRAY = @()
        $R53_ZONE_LIST.hostedZones | ForEach-Object {
            $ZONE = $_
            $R53_ZONE_ID = $ZONE.Id
            # Get the string following the last / in the ID
            $R53_ZONE_ID = $R53_ZONE_ID.Split('/')[-1]
            $R53_ZONE_NAME = $ZONE.Name
            $R53_ZONE_NAME = $R53_ZONE_NAME.TrimEnd('.')
            $R53_ZONE_NAME = $R53_ZONE_NAME.Replace('.', '-')

            $R53_ZONE_RR_JSON = aws route53 list-resource-record-sets --hosted-zone-id $R53_ZONE_ID --profile $AWS_PROFILE --output json 
            #$R53_ZONE_RR_JSON | Out-File ".\$AWS_PROFILE-R53Zone-$R53_ZONE_NAME-$PULL_DATE.json" -Encoding utf8
            $R53_ZONE_OBJECT = $R53_ZONE_RR_JSON | ConvertFrom-Json -Depth 10
            # BUILD A HOSTED_ZONE_ARRAY with the following columns: DNS Name, Record Type, TTL, Record Value/s, Zone Name, Zone ID, AWS Profile
            $R53_ZONE_OBJECT.ResourceRecordSets | ForEach-Object {
                # Check if AliasTarget is null and set it to 'N/A'
                if (!$_.AliasTarget) {
                    $RR_VALUES = ($_.ResourceRecords | ForEach-Object { $_.Value }) -join ','
                } else {
                    $RR_VALUES = $_.AliasTarget.DNSName
                }
                # if $_.TTL is null set it to 'N/A'
                if (!$_.TTL) {
                    $TTL = 'N/A'
                } else {
                    $TTL = $_.TTL
                }
                # If 
                $RR = @($_.Name, $_.Type, $TTL, $RR_VALUES,$R53_ZONE_NAME, $R53_ZONE_ID, $SIMPLE_NAME, $AWS_PROFILE)
                #Check if the RR element count is 7
                if ($RR.Count -ne 8) {
                    Write-Error "RR element count is not 8: $RR"
                }
                Write-Debug "Adding RR $RR"
                # Add the RR as an array to the HOSTED_ZONE_ARRAY
                $HOSTED_ZONE_ARRAY += ,$RR
            }
            #Write-Debug "HOSTED_ZONE_ARRAY: $HOSTED_ZONE_ARRAY"
            #Write-Debug "HOSTED_ZONE_ARRAY.Count: $($HOSTED_ZONE_ARRAY.Count)"
            #Write-Debug "R53_ZONE_OBJECT: $R53_ZONE_OBJECT"
        } 
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = $SheetName

        $Data = $workbook.Worksheets.Item(1)
        $Data.Name = $SheetName
        $Data.Cells.Item(1, 1) = 'DNS Name'
        $Data.Cells.Item(1, 2) = 'Record Type'
        $Data.Cells.Item(1, 3) = 'TTL'
        $Data.Cells.Item(1, 4) = 'Record Value/s'
        $Data.Cells.Item(1, 5) = 'Zone Name'
        $Data.Cells.Item(1, 6) = 'Zone ID'
        $Data.Cells.Item(1, 7) = 'Account Label'
        $Data.Cells.Item(1, 8) = 'AWS Profile'

        # Insert Data
        # function to add row with data as array parameter
        function Add-Row {
            param(
                [Parameter(Mandatory = $true)]
                [int]$ROW_NUM,
                # add mandadataory parameter for array of data to add to row validating it is not null and not empty and contains 12 items
                [Parameter(Mandatory = $true)]
                [ValidateCount(8, 8)]
                [array]$ROW_DATA
            )
            $col = 1
            foreach ($item in $ROW_DATA) {
                $Data.Cells.Item($ROW_NUM, $col) = $item
                $col++
            }
        }
        $ROW_NUM = 2
        foreach ($RR in $HOSTED_ZONE_ARRAY) {
            Write-Debug "Adding row $ROW_NUM with data $RR"
            Add-Row -ROW_NUM $ROW_NUM -ROW_DATA $RR
            $ROW_NUM++
        }
        # for each sheet in workbook
        foreach ($sheet in $workbook.Worksheets) {
            $sheet.Activate()
            $sheet.Application.ActiveWindow.SplitRow = 1
            $sheet.Application.ActiveWindow.FreezePanes = $true
            $sheet.Rows.Item(1).Font.Bold = $true
            $sheet.Rows.Item(1).Borders.Item(9).LineStyle = 1
            $sheet.Rows.Item(1).Borders.Item(9).Weight = 2
            # add filter
            $sheet.Cells.Item(1, 1).EntireRow.AutoFilter() | Out-Null
            # autofit the columns
            $sheet.UsedRange.EntireColumn.AutoFit() | Out-Null   
        }
        # Save and close
        $workbook.SaveAs($REPORT_PATH)
        $workbook.Close()
        $excel.Quit()
    }
    $PULL_DATE = Get-Date -Format 'yyyyMMdd'
    if (!$AWS_PROFILE) {
        $AWS_PROFILE = Select-AWSProfile
    }
    $REPORT_PATH = "${PWD}\Route53-Scan-Report-${SIMPLE_NAME}-${PULL_DATE}.xlsx"
    if (!$R53_ZONE_LIST) {
        $R53_ZONE_LIST = Get-Zonelist -AWS_PROFILE $AWS_PROFILE
    }
    Write-Debug "REPORT_PATH: $REPORT_PATH, SheetName: ${SIMPLE_NAME}-${PULL_DATE}, SIMPLE_NAME: $SIMPLE_NAME, AWS_PROFILE: $AWS_PROFILE, R53_ZONE_LIST: $R53_ZONE_LIST"
    Export-ExcelFile -REPORT_PATH "$REPORT_PATH" -SheetName "${SIMPLE_NAME}-${PULL_DATE}" -R53_ZONE_LIST ${R53_ZONE_LIST} -AWS_PROFILE $AWS_PROFILE -SIMPLE_NAME $SIMPLE_NAME
}
