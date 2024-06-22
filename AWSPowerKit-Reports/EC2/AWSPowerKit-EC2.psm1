# Set exit on error
$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'
$env:AWSPowerKitAWSProfile = ''

function Select-AWSProfile {
    # Show list of AWS profiles configured on the system with numbered selection then prompt for selection and set the $AWS_PROFILE variable to the corresponding profile name, not the profile number
    $AWS_PROFILES = aws configure list-profiles
    $AWS_PROFILES | ForEach-Object -Begin { $i = 1 } -Process { Write-Host "$i. $_"; $i++ }
    $AWS_PROFILE = Read-Host 'Select AWS Profile Number'
    $AWS_PROFILE = $AWS_PROFILES[$AWS_PROFILE - 1]
    Write-Debug "You selected profile $AWS_PROFILE"
    $AWS_PROFILE
}

function Get-AccountList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AWS_PROFILE
    )
    if (!(aws sts get-caller-identity --profile $AWS_PROFILE)) {
        Write-Debug "AWS_PROFILE $AWS_PROFILE is not valid attempting re-auth"
        aws configure sso --profile $AWS_PROFILE
        #$env:AWSPowerKitAWSProfile = Select-AWSProfile
    }
    $ACCOUNT_LIST = aws organizations list-accounts --profile $AWS_PROFILE
    Write-Debug "REPO_LIST: $ACCOUNT_LIST"
    $ACC_LIST_OBJECT = $ACCOUNT_LIST | ConvertFrom-Json
    Write-Debug "ACC_LIST_OBJECT: $ACC_LIST_OBJECT"
    $ACC_LIST_OBJECT
}

function Get-EC2Report {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$AWS_PROFILE,
        [Parameter(Mandatory = $false)]
        [System.Object]$ECR_REPO_LIST,
        [Parameter(Mandatory = $true)]
        [string]$SIMPLE_NAME
    )
    # Function that creates a new excel file and adds a worksheet
    function Export-ExcelFile {
        param(
            [Parameter(Mandatory = $true)]
            [string]$REPORT_PATH,
            [Parameter(Mandatory = $true)]
            [string]$SheetName,
            [Parameter(Mandatory = $true)]
            [System.Object]$ECR_REPO_LIST,
            [Parameter(Mandatory = $true)]
            [string]$AWS_PROFILE,
            [Parameter(Mandatory = $true)]
            [string]$SIMPLE_NAME
        )
        $PULL_DATE = Get-Date -Format 'yyyyMMdd'
        $ACCOUNT_LIST = Get-AccountList -AWS_PROFILE $AWS_PROFILE
        $ACCOUNT_LIST = $ACCOUNT_LIST | ConvertFrom-Json
        $CRIT_COUNT = 0
        $HIGH_COUNT = 0
        $MEDIUM_COUNT = 0
        $LOW_COUNT = 0
        $INFO_COUNT = 0
        $UNDEFINED_COUNT = 0
        $OLDEST_IMAGE_AGE = 0
        $IMAGE_COUNT = 0
        $CRIT_LIST = @()
        $HIGH_LIST = @()

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = $SheetName

        $Data = $workbook.Worksheets.Item(1)
        $Data.Name = $SheetName
        $Data.Cells.Item(1, 1) = 'Repo Name'
        $Data.Cells.Item(1, 2) = 'Latest Image'
        $Data.Cells.Item(1, 3) = 'Image Pushed'
        $Data.Cells.Item(1, 4) = 'Scan Status'
        $Data.Cells.Item(1, 5) = 'Image Scanned'
        $Data.Cells.Item(1, 6) = 'Latest Age (Days)'
        $Data.Cells.Item(1, 7) = 'Last Scan (Days)'
        $Data.Cells.Item(1, 8) = 'CRITICAL'
        $Data.Cells.Item(1, 9) = 'HIGH'
        $Data.Cells.Item(1, 10) = 'MEDIUM'
        $Data.Cells.Item(1, 11) = 'LOW'
        $Data.Cells.Item(1, 12) = 'INFO'
        $Data.Cells.Item(1, 13) = 'UNDEFINED'
        $Data.Cells.Item(1, 14) = 'Image count'
        $Data.Cells.Item(1, 15) = 'Description'

        # Insert Data
        # function to add row with data as array parameter
        function Add-Row {
            param(
                [Parameter(Mandatory = $true)]
                [int]$ROW_NUM,
                # add mandadataory parameter for array of data to add to row validating it is not null and not empty and contains 12 items
                [Parameter(Mandatory = $true)]
                [ValidateCount(15, 15)]
                [array]$ROW_DATA
            )
            $col = 1
            foreach ($item in $ROW_DATA) {
                $Data.Cells.Item($ROW_NUM, $col) = $item
                $col++
            }
        }
        $ROW_NUM = 2
        foreach ($REPO in $ECR_REPO_LIST.repositories) {
            #write-host "Processing $($REPO.repositoryName)"
            $image_list_json = aws ecr describe-images --repository-name $($REPO.repositoryName) --query 'sort_by(imageDetails,& imagePushedAt)[*]' --profile $AWS_PROFILE --output json
            $image_list = $image_list_json | ConvertFrom-Json -Dept 10
            # Write-Debug formatted image list with all properties
            #$image_list_json | ForEach-Object { Write-Debug $_ }

            # If image list is empty, set latest image to N/A
            if ($image_list.count -eq 0) {
                $LATEST_IMAGE = 'N/A'
            }
            else {
                $LATEST_IMAGE = $image_list[-1]
            }
            # If image count is set all values to N/A and add row to excel sheet
            if ($image_list.count -eq 0) {
                $rdata = @(
                    $($REPO.repositoryName), 'No images in repo', 
                    $(Get-Date -Format 'yyyyMMdd_hhmmss'), 'No images in repo', $(Get-Date -Format 'yyyyMMdd_hhmmss'), 0, 0, 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', $image_list.count, 'No images found'
                )
            }
            # IF IMAGE SCAN STATUS IS NOT COMPLETE
            else {
                $LATEST_IMAGE_SCAN_FINDINGS_JSON = aws ecr describe-image-scan-findings --repository-name $($REPO.repositoryName) --image-id imageDigest=$($LATEST_IMAGE.imageDigest) --profile $AWS_PROFILE
                $LATEST_IMAGE_SCAN_FINDINGS = $LATEST_IMAGE_SCAN_FINDINGS_JSON | ConvertFrom-Json -Depth 30
                # If image scan not complete/host no details
                if (!$LATEST_IMAGE_SCAN_FINDINGS.imageScanStatus) {
                    $rdata = @($($REPO.repositoryName), $($LATEST_IMAGE.imageTags[0]),
                        $(Get-Date -Format 'yyyyMMdd_hhmmss'), 'N/A', $(Get-Date -Format 'yyyyMMdd_hhmmss'), 0, 0, 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', $image_list.count, 'PowerKit unable to parse response - check in console'
                    )
                }
                elseif ($LATEST_IMAGE_SCAN_FINDINGS.imageScanStatus.Status -ne 'COMPLETE') {
                    $rdata = @(
                        $($REPO.repositoryName),
                        $($LATEST_IMAGE.imageTags[0]),
                        $LATEST_IMAGE.imagePushedAt,
                        $LATEST_IMAGE_SCAN_FINDINGS.imageScanStatus.Status,
                        $LATEST_IMAGE.imagePushedAt,
                        [math]::Round(([datetime]::Now - [datetime]$LATEST_IMAGE.imagePushedAt).TotalDays),
                        [math]::Round(([datetime]::Now - [datetime]$LATEST_IMAGE.imagePushedAt).TotalDays),
                        'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A',
                        $image_list.count,
                        "Image scan not completed: $($LATEST_IMAGE_SCAN_FINDINGS.imageScanStatus.Status) - $($LATEST_IMAGE_SCAN_FINDINGS.imageScanStatus.description)"
                    )
                } 
                else {
                    $rdata = @(
                        $($REPO.repositoryName),
                        $($LATEST_IMAGE.imageTags[0]),
                        $LATEST_IMAGE.imagePushedAt,
                        $LATEST_IMAGE_SCAN_FINDINGS.imageScanStatus.Status,
                        $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.imageScanCompletedAt,
                        [math]::Round(([datetime]::Now - [datetime]$LATEST_IMAGE.imagePushedAt).TotalDays),
                        [math]::Round(([datetime]::Now - [datetime]$LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.imageScanCompletedAt).TotalDays),
                        $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.CRITICAL,
                        $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.HIGH,
                        $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.MEDIUM,
                        $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.LOW,
                        $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.INFORMATIONAL,
                        $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.UNDEFINED,
                        $image_list.count,
                        "$($LATEST_IMAGE_SCAN_FINDINGS.imageScanStatus.Status) - $($LATEST_IMAGE_SCAN_FINDINGS.imageScanStatus.description)"
                    )
                    if ($LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.CRITICAL -gt 0 -or $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.HIGH -gt 0) {
                        if ($LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findings) {
                            $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findings | ForEach-Object {
                                if ($_.severity -eq 'CRITICAL') {
                                    $CRIT_FINDING = @($_.name, $_.description, $_.severity, $_.uri, '', '', "$($REPO.repositoryName):$($LATEST_IMAGE.imageTags[0])") 
                                    $CRIT_LIST += , $CRIT_FINDING
                                }
                                if ($_.severity -eq 'HIGH') {
                                    $HIGH_FINDING = @($_.name, $_.description, $_.severity, $_.uri, '', '', "$($REPO.repositoryName):$($LATEST_IMAGE.imageTags[0])")
                                    $HIGH_LIST += , $HIGH_FINDING
                                }
                            }
                        }
                        if ($LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.enhancedFindings) {
                            $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.enhancedFindings | ForEach-Object {
                                if ($_.severity -eq 'CRITICAL') {
                                    $SOURCE_URL = $_.packageVulnerabilityDetails.sourceUrl
                                    # join values for referenceUrls into a single string
                                    $REF_URLS = $_.packageVulnerabilityDetails.referenceUrls
                                    $REF_URLS = $REF_URLS -Join ', '
                                    $CRIT_FINDING = @($_.title, $_.description, $_.severity, $SOURCE_URL, $REF_URLS, $_.type, "$($REPO.repositoryName):$($LATEST_IMAGE.imageTags[0])") 
                                    $CRIT_LIST += , $CRIT_FINDING
                                }
                                if ($_.severity -eq 'HIGH') {
                                    $SOURCE_URL = $_.attributes | Where-Object { $_.key -eq 'sourceUrl' } | Select-Object -First 1 -ExpandProperty value
                                    # join values for referenceUrls into a single string
                                    $REF_URLS = $_.attributes | Where-Object { $_.key -eq 'referenceUrl' } | Select-Object -ExpandProperty value
                                    $REF_URLS = $REF_URLS -Join ', '
                                    $HIGH_FINDING = @($_.title, $_.description, $_.severity, $SOURCE_URL, $REF_URLS, $_.type, "$($REPO.repositoryName):$($LATEST_IMAGE.imageTags[0])")
                                    $HIGH_LIST += , $HIGH_FINDING
                                }
                            }
                        }

                    }
                } 
                if ($rdata.count -ne 15) {
                    Write-Host 'Array count is not 15, exiting script'
                    exit
                }
                # replace empty and null values with N/A
                $rdata = $rdata | ForEach-Object { if ($_ -eq $null -or $_ -eq '') { 0 } else { $_ } }
        
                # Write rdata array to console
                #$rdata | ForEach-Object { Write-Host $_ }
                Add-Row -ROW_NUM $ROW_NUM -ROW_DATA $rdata
                $ROW_NUM++
                $CRIT_COUNT += $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.CRITICAL
                $HIGH_COUNT += $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.HIGH
                $MEDIUM_COUNT += $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.MEDIUM
                $LOW_COUNT += $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.LOW
                $INFO_COUNT += $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.INFORMATIONAL
                $UNDEFINED_COUNT += $LATEST_IMAGE_SCAN_FINDINGS.imageScanFindings.findingSeverityCounts.UNDEFINED
                $IMAGE_COUNT += $image_list.count

                if ($image_list.count -ne 0) {
                    $OLDEST_IMAGE_AGE = [math]::Max($OLDEST_IMAGE_AGE, $rdata[5])
                }
            }
            # Add a summary sheet that counts the total number of findings by severity with CRITICAL in red, HIGH in yellow, MEDIUM in blue, LOW in green, INFO in white, UNDEFINED in white, also include the average age of the scan/push in days
            $summary = $workbook.Worksheets.Add()
            $summary.Name = "ECR-Summary-$SIMPLE_NAME-$PULL_DATE"
            $summary.Cells.Item(1, 1) = 'Severity'
            $summary.Cells.Item(1, 2) = 'Count'
            $summary.Cells.Item(2, 1) = 'CRITICAL'
            # Add the total number of CRITICAL findings in red
            $summary.Cells.Item(2, 2).Font.ColorIndex = 3
            $summary.Cells.Item(2, 2) = "$CRIT_COUNT"
            $summary.Cells.Item(3, 1) = 'HIGH'
            $summary.Cells.Item(3, 2).Font.ColorIndex = 6
            $summary.Cells.Item(3, 2) = "$HIGH_COUNT"
            $summary.Cells.Item(4, 1) = 'MEDIUM'
            $summary.Cells.Item(4, 2).Font.ColorIndex = 5
            $summary.Cells.Item(4, 2) = "$MEDIUM_COUNT"
            $summary.Cells.Item(5, 1) = 'LOW'
            $summary.Cells.Item(5, 2).Font.ColorIndex = 4
            $summary.Cells.Item(5, 2) = "$LOW_COUNT"
            $summary.Cells.Item(6, 1) = 'INFO'
            $summary.Cells.Item(6, 2).Font.ColorIndex = 2
            $summary.Cells.Item(6, 2) = "$INFO_COUNT"
            $summary.Cells.Item(7, 1) = 'UNDEFINED'
            $summary.Cells.Item(7, 2).Font.ColorIndex = 1
            $summary.Cells.Item(7, 2) = "$UNDEFINED_COUNT"
            # Add oldest scan/push age in days of all repos
            $summary.Cells.Item(8, 1) = 'Oldest Image (Days)'
            $summary.Cells.Item(8, 2) = "$OLDEST_IMAGE_AGE"
            # Add total number of images
            $summary.Cells.Item(9, 1) = 'Total Images'
            $summary.Cells.Item(9, 2) = "$IMAGE_COUNT"
            # function to create Criticals and Highs sheets, add headers and write data, headers are: Name, Description, Severity, URL, Other Refs, Type, Repo:Tag
            function Add-Data {
                param(
                    [Parameter(Mandatory = $true)]
                    [string]$LIST_SHEET_NAME,
                    [Parameter(Mandatory = $true)]
                    [array]$DATA
                )
                # Ensure sheet is add after the last sheet
                $workbook.Worksheets.Item($workbook.Worksheets.Count)
                $sheet = $workbook.Worksheets.Add($workbook.Worksheets.Item($workbook.Worksheets.Count))
                $sheet.Name = $LIST_SHEET_NAME
                $sheet.Cells.Item(1, 1) = 'Name'
                $sheet.Cells.Item(1, 2) = 'Description'
                $sheet.Cells.Item(1, 3) = 'Severity'
                $sheet.Cells.Item(1, 4) = 'URL'
                $sheet.Cells.Item(1, 5) = 'Other Refs'
                $sheet.Cells.Item(1, 6) = 'Type'
                $sheet.Cells.Item(1, 7) = 'Repo:Tag'
                $ROW_NUM = 2
                foreach ($FINDING in $DATA) {
                    $col = 1
                    foreach ($item in $FINDING) {
                        $sheet.Cells.Item($ROW_NUM, $col) = $item
                        $col++
                    }
                    $ROW_NUM++
                }
            }
            Add-Data -LIST_SHEET_NAME 'CRITICALS' -DATA $CRIT_LIST
            Add-Data -LIST_SHEET_NAME 'HIGHS' -DATA $HIGH_LIST
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
                # autofit the columns only for the first 2 worksheets
                if ($sheet.Name -eq $SheetName -or $sheet.Name -eq "ECR-Summary-$SIMPLE_NAME-$PULL_DATE") {
                    $sheet.Columns.AutoFit() | Out-Null
                }
                else {
                    #wrap text for all other sheets
                    $sheet.Cells.WrapText = $false
                    # set width of columns to 50, 50, 15, 40, 15, 15
                    $sheet.Columns.Item(1).ColumnWidth = 50
                    $sheet.Columns.Item(2).ColumnWidth = 50
                    $sheet.Columns.Item(3).ColumnWidth = 15
                    $sheet.Columns.Item(4).ColumnWidth = 40
                    $sheet.Columns.Item(5).ColumnWidth = 15
                    $sheet.Columns.Item(6).ColumnWidth = 15
                }
            }
            # Save and close
            $workbook.SaveAs($REPORT_PATH)
            $workbook.Close()
            $excel.Quit()
        }
    }
    $PULL_DATE = Get-Date -Format 'yyyyMMdd'
    if (!$env:AWSPowerKitAWSProfile) {
        $env:AWSPowerKitAWSProfile = Select-AWSProfile
    }
    $REPORT_PATH = "${PWD}\ECR-Scan-Report-${SIMPLE_NAME}-${PULL_DATE}.xlsx"
    if (!$AWS_ACCOUNT_LIST) {
        $AWS_ACCOUNT_LIST = Get-AccountList $env:AWSPowerKitAWSProfile
    }
    Write-Debug "REPORT_PATH: $REPORT_PATH, SheetName: ${SIMPLE_NAME}-${PULL_DATE}, AWS_PROFILE: $AWS_PROFILE, AWS_ACCOUNT_LIST: $AWS_ACCOUNT_LIST"
    # Export-ExcelFile -REPORT_PATH "$REPORT_PATH" -SheetName "${SIMPLE_NAME}-${PULL_DATE}" -ECR_REPO_LIST ${ECR_REPO_LIST} -AWS_PROFILE $AWS_PROFILE -SIMPLE_NAME $SIMPLE_NAME
    Write-Debug "Report saved to $REPORT_PATH"
}
