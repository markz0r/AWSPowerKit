# Set exit on error
$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Select-AWSProfile {
    # Show list of AWS profiles configured on the system with numbered selection then prompt for selection and set the $AWS_PROFILE variable to the corresponding profile name, not the profile number
    $AWS_PROFILES = aws configure list-profiles
    $AWS_PROFILES | ForEach-Object -Begin { $i = 1 } -Process { Write-Host "$i. $_"; $i++ }
    $AWS_PROFILE = Read-Host 'Select AWS Profile Number'
    $AWS_PROFILE = $AWS_PROFILES[$AWS_PROFILE - 1]
    Write-Debug "You selected profile $AWS_PROFILE"
    $AWS_PROFILE
}

function Get-Repolist {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AWS_PROFILE
    )
    $REPO_LIST = aws ecr describe-repositories --profile $AWS_PROFILE
    Write-Debug "REPO_LIST: $REPO_LIST"
    $REPO_LIST_OBJECT = $REPO_LIST | ConvertFrom-Json
    $REPO_LIST_OBJECT
}

function Add-Row {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ROW_ENTRY,
        [Parameter(Mandatory = $true)]
        [ValidateCount(15, 15)]
        [array]$ROW_DATA,
        [Parameter(Mandatory = $true)]
        [Object]$ExcelSheet
    )
    $col = 1
    Write-Debug "Adding row $ROW_ENTRY with for $($ROW_DATA[0]) to $($ExcelSheet.Name)" 
    #Write-Debug $($ROW_DATA | Format-Table | Out-String)
    #Write-Debug ' ... being iteration'
    $ROW_DATA | ForEach-Object {
        #Write-Debug "Adding item $_ to row $ROW_ENTRY column $col"
        $ExcelSheet.Cells.Item($ROW_ENTRY, $col) = [string]$_
        $col++
    }
    return $ROW_ENTRY + 1
}

function Add-Data {
    param(
        [Parameter(Mandatory = $true)]
        [Object]$ExcelSheet,
        [Parameter(Mandatory = $true)]
        [array]$DATA
    )
    $ExcelSheet.Cells.Item(1, 1) = 'Name'
    $ExcelSheet.Cells.Item(1, 2) = 'Description'
    $ExcelSheet.Cells.Item(1, 3) = 'Severity'
    $ExcelSheet.Cells.Item(1, 4) = 'URL'
    $ExcelSheet.Cells.Item(1, 5) = 'Other Refs'
    $ExcelSheet.Cells.Item(1, 6) = 'Type'
    $ExcelSheet.Cells.Item(1, 7) = 'Repo:Tag'
    $ROW_NUM = 2
    foreach ($FINDING in $DATA) {
        $col = 1
        foreach ($item in $FINDING) {
            $ExcelSheet.Cells.Item($ROW_NUM, $col) = $item
            $col++
        }
        $ROW_NUM += 1
    }
}

function Add-Summary {
    param(
        [Parameter(Mandatory = $true)]
        [Object]$ExcelSheet,
        [Parameter(Mandatory = $true)]
        [Object]$SUMMARY_MAP
    )
    # Add a summary sheet that counts the total number of findings by severity with CRITICAL in red, HIGH in yellow, MEDIUM in blue, LOW in green, INFO in white, UNDEFINED in white, also include the average age of the scan/push in days
    Write-Debug "SUMMARY_MAP is: $($SUMMARY_MAP.GetType())"

    $ExcelSheet.Cells.Item(1, 1) = 'Severity'
    $ExcelSheet.Cells.Item(1, 2) = 'Count'
    $ExcelSheet.Cells.Item(2, 1) = 'CRITICAL'
    # Add the total number of CRITICAL findings in red
    $ExcelSheet.Cells.Item(2, 2).Font.ColorIndex = 3
    $ExcelSheet.Cells.Item(2, 2) = $SUMMARY_MAP.CRIT_COUNT
    $ExcelSheet.Cells.Item(3, 1) = 'HIGH'
    $ExcelSheet.Cells.Item(3, 2).Font.ColorIndex = 6
    $ExcelSheet.Cells.Item(3, 2) = $SUMMARY_MAP.HIGH_COUNT
    $ExcelSheet.Cells.Item(4, 1) = 'MEDIUM'
    $ExcelSheet.Cells.Item(4, 2).Font.ColorIndex = 5
    $ExcelSheet.Cells.Item(4, 2) = $SUMMARY_MAP.MEDIUM_COUNT
    $ExcelSheet.Cells.Item(5, 1) = 'LOW'
    $ExcelSheet.Cells.Item(5, 2).Font.ColorIndex = 4
    $ExcelSheet.Cells.Item(5, 2) = $SUMMARY_MAP.LOW_COUNT
    $ExcelSheet.Cells.Item(6, 1) = 'INFO'
    $ExcelSheet.Cells.Item(6, 2).Font.ColorIndex = 2
    $ExcelSheet.Cells.Item(6, 2) = $SUMMARY_MAP.INFO_COUNT
    $ExcelSheet.Cells.Item(7, 1) = 'UNDEFINED'
    $ExcelSheet.Cells.Item(7, 2).Font.ColorIndex = 1
    $ExcelSheet.Cells.Item(7, 2) = $SUMMARY_MAP.UNDEFINED_COUNT
    # Add oldest scan/push age in days of all repos
    $ExcelSheet.Cells.Item(8, 1) = 'Oldest Image (Days)'
    $ExcelSheet.Cells.Item(8, 2) = "$OLDEST_IMAGE_AGE"
    # Add total number of images
    $ExcelSheet.Cells.Item(9, 1) = 'Total Images'
    $ExcelSheet.Cells.Item(9, 2) = "$IMAGE_COUNT"
}
function Format-ExcelSheets {
    param(
        [Parameter(Mandatory = $true)]
        [Object]$workbook
    )
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
}

Function New-ImageScanReportSheet {
    param(
        [Parameter(Mandatory = $true)]
        [System.Object]$ECR_REPO_LIST,
        [Parameter(Mandatory = $true)]
        [string]$AWS_PROFILE,
        [Parameter(Mandatory = $true)]
        [string]$SIMPLE_NAME,
        [Parameter(Mandatory = $true)]
        [Object]$ExcelSheet

    )
    $CRIT_COUNT = 0; $HIGH_COUNT = 0; $MEDIUM_COUNT = 0; $LOW_COUNT = 0; $INFO_COUNT = 0; $UNDEFINED_COUNT = 0; $OLDEST_IMAGE_AGE = 0; $IMAGE_COUNT = 0
    $CRIT_LIST = @(); $HIGH_LIST = @()
    $ExcelSheet.Cells.Item(1, 1) = 'Repo Name'
    $ExcelSheet.Cells.Item(1, 2) = 'Latest Image'
    $ExcelSheet.Cells.Item(1, 3) = 'Image Pushed'
    $ExcelSheet.Cells.Item(1, 4) = 'Scan Status'
    $ExcelSheet.Cells.Item(1, 5) = 'Image Scanned'
    $ExcelSheet.Cells.Item(1, 6) = 'Latest Age (Days)'
    $ExcelSheet.Cells.Item(1, 7) = 'Last Scan (Days)'
    $ExcelSheet.Cells.Item(1, 8) = 'CRITICAL'
    $ExcelSheet.Cells.Item(1, 9) = 'HIGH'
    $ExcelSheet.Cells.Item(1, 10) = 'MEDIUM'
    $ExcelSheet.Cells.Item(1, 11) = 'LOW'
    $ExcelSheet.Cells.Item(1, 12) = 'INFO'
    $ExcelSheet.Cells.Item(1, 13) = 'UNDEFINED'
    $ExcelSheet.Cells.Item(1, 14) = 'Image count'
    $ExcelSheet.Cells.Item(1, 15) = 'Description'
    $IMAGE_ROW_NUM = 2
    foreach ($REPO in $ECR_REPO_LIST.repositories) {
        $image_list_json = aws ecr describe-images --repository-name $($REPO.repositoryName) --query 'sort_by(imageDetails,& imagePushedAt)[*]' --profile $AWS_PROFILE --output json | Out-String
        try {
            $image_list = $image_list_json | ConvertFrom-Json -Depth 10
        }
        catch {
            Write-Error "Failed to parse image list for repo $($REPO.repositoryName)"
        }

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
            $LATEST_IMAGE_SCAN_FINDINGS_JSON = aws ecr describe-image-scan-findings --repository-name $($REPO.repositoryName) --image-id imageDigest=$($LATEST_IMAGE.imageDigest) --profile $AWS_PROFILE --output json | Out-String
            try {
                $LATEST_IMAGE_SCAN_FINDINGS_JSON = $LATEST_IMAGE_SCAN_FINDINGS_JSON -replace '\uff1a', ':'
                # If latest image scan findings contains non-ascii characters, write warning and replace with 'unparseable'
                if ($LATEST_IMAGE_SCAN_FINDINGS_JSON -match '[^\x00-\x7F]') {
                    Write-Warning "Non-ASCII characters detected in image scan findings for repo $($REPO.repositoryName)"
                    Write-Debug 'Replacing with unparseable'
                    $LATEST_IMAGE_SCAN_FINDINGS_JSON -replace '[^\x00-\x7F]', 'unparseable'
                }
                $LATEST_IMAGE_SCAN_FINDINGS = $LATEST_IMAGE_SCAN_FINDINGS_JSON | ConvertFrom-Json -Depth 10
            }
            catch {
                # Write the failed JSON to an Error file
                $PULL_DATE = Get-Date -Format 'yyyyMMdd'
                $LATEST_IMAGE_SCAN_FINDINGS_JSON | Out-File -FilePath "ECR-Scan-Report-${SIMPLE_NAME}-${PULL_DATE}-Error.json" -Append                
                #Write-Debug $LATEST_IMAGE_SCAN_FINDINGS_JSON
                Write-Warning "Failed to parse image scan findings for repo $($REPO.repositoryName), check ECR-Scan-Report-${SIMPLE_NAME}-${PULL_DATE}-Error.json for details"
                continue
            }
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
                Write-Error 'Array count is not 15, exiting script'
                exit
            }
            # replace empty and null values with N/A
            $rdata = $rdata | ForEach-Object { if ([string]::IsNullOrEmpty($_)) { 'N/A' } else { $_ } }

            #Write-Debug "Calling Add-Row [$IMAGE_ROW_NUM] with rdata: $rdata"
            $IMAGE_ROW_NUM = Add-Row -ROW_ENTRY $IMAGE_ROW_NUM -ROW_DATA $rdata -ExcelSheet $ExcelSheet
            #Write-Debug "Added row, row number is now: $IMAGE_ROW_NUM"
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
    }
    $SUMMARY_MAP = @{
        CRIT_COUNT       = $CRIT_COUNT
        HIGH_COUNT       = $HIGH_COUNT
        MEDIUM_COUNT     = $MEDIUM_COUNT
        LOW_COUNT        = $LOW_COUNT
        INFO_COUNT       = $INFO_COUNT
        UNDEFINED_COUNT  = $UNDEFINED_COUNT
        OLDEST_IMAGE_AGE = $OLDEST_IMAGE_AGE
        IMAGE_COUNT      = $IMAGE_COUNT
        CRIT_LIST        = $CRIT_LIST
        HIGH_LIST        = $HIGH_LIST
    }
    return $SUMMARY_MAP
}

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
    $ECR_REPO_JSON = aws ecr describe-repositories --profile $AWS_PROFILE --output json | Out-String
    try {
        $ECR_REPO_LIST = $ECR_REPO_JSON | ConvertFrom-Json
    }
    catch {
        Write-Error 'Failed to parse ECR Repo list'
    }

    $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false; $workbook = $excel.Workbooks.Add() 
    $sheet = $workbook.Worksheets.Item(1); $sheet.Name = $SheetName

    # Insert Data
    # function to add row with data as array parameter
    $SUMMARY_MAP = New-ImageScanReportSheet -ECR_REPO_LIST $ECR_REPO_LIST -AWS_PROFILE $AWS_PROFILE -SIMPLE_NAME $SIMPLE_NAME -ExcelSheet $sheet
    Write-Debug "SUMMARY_MAP Object type: $($SUMMARY_MAP.GetType())"
    Write-Debug "Summary Map Data Contents: $($SUMMARY_MAP | Format-List | Out-String)"
    $summary_sheet = $workbook.Worksheets.Add()
    $summary_sheet.Name = "ECR-Summary-$SIMPLE_NAME-$PULL_DATE"
    Add-Summary -ExcelSheet $summary_sheet -SUMMARY_MAP $SUMMARY_MAP

    $crit_image_data_sheet = $workbook.Worksheets.Add()
    $crit_image_data_sheet.Name = "Crital-ECRList-$SIMPLE_NAME-$PULL_DATE"
    Add-Data -ExcelSheet $crit_image_data_sheet -DATA $SUMMARY_MAP.CRIT_LIST

    $high_image_data_sheet = $workbook.Worksheets.Add()
    $high_image_data_sheet.Name = "High-ECRList-$SIMPLE_NAME-$PULL_DATE"
    Add-Data -ExcelSheet $high_image_data_sheet -DATA $SUMMARY_MAP.HIGH_LIST

    Format-ExcelSheets -workbook $workbook

    # Save and close
    $workbook.SaveAs($REPORT_PATH)
    $workbook.Close()
    $excel.Quit()
}

function Get-ECRReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$AWS_PROFILE,
        [Parameter(Mandatory = $false)]
        [System.Object]$ECR_REPO_LIST,
        [Parameter(Mandatory = $true)]
        [string]$SIMPLE_NAME
    )
    $PULL_DATE = Get-Date -Format 'yyyyMMdd'
    if (!$AWS_PROFILE) {
        $AWS_PROFILE = Select-AWSProfile
    }
    $REPORT_PATH = "${PWD}\ECR-Scan-Report-${SIMPLE_NAME}-${PULL_DATE}.xlsx"
    if (!$ECR_REPO_LIST) {
        $ECR_REPO_LIST = Get-Repolist -AWS_PROFILE $AWS_PROFILE
    }
    Write-Debug "REPORT_PATH: $REPORT_PATH, SheetName: ${SIMPLE_NAME}-${PULL_DATE}, AWS_PROFILE: $AWS_PROFILE, ECR_REPO_LIST: $ECR_REPO_LIST"
    Export-ExcelFile -REPORT_PATH "$REPORT_PATH" -SheetName "${SIMPLE_NAME}-${PULL_DATE}" -ECR_REPO_LIST ${ECR_REPO_LIST} -AWS_PROFILE $AWS_PROFILE -SIMPLE_NAME $SIMPLE_NAME
    Write-Debug "Report saved to $REPORT_PATH"
}
