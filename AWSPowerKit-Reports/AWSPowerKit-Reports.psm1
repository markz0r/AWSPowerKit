# Set exit on error
$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'
$script:SELECT_PROFILE = $false

function Select-AWSProfile {
    param(
        [Parameter(Mandatory = $false)]
        [string]$AWS_PROFILE
    )
    if ($null -eq $AWS_PROFILE) {
        # Show list of AWS profiles configured on the system with numbered selection then prompt for selection and set the $AWS_PROFILE variable to the corresponding profile name, not the profile number
        $AWS_PROFILES = aws configure list-profiles
        $AWS_PROFILES | ForEach-Object -Begin { $i = 1 } -Process { Write-Host "$i. $_"; $i++ }
        $AWS_PROFILE = Read-Host 'Select AWS Profile Number'
        $AWS_PROFILE = $AWS_PROFILES[$AWS_PROFILE - 1]
    }
    Write-Debug "You selected profile $AWS_PROFILE"
    $script:SELECT_PROFILE = $false
}

function Get-ECRReport {
    # Function that creates a new excel file and adds a worksheet
    function Export-ExcelFile {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,
            [Parameter(Mandatory = $true)]
            [string]$SheetName,
            [Parameter(Mandatory = $true)]
            [System.Object]$ECR_REPO_LIST
        )
        $PULL_DATE = Get-Date -Format 'yyyyMMdd'
        $ECR_REPO_JSON = aws ecr describe-repositories --profile $AWS_PROFILE 
        $ECR_REPO_LIST = $ECR_REPO_JSON | ConvertFrom-Json
        $CRIT_COUNT = 0
        $HIGH_COUNT = 0
        $MEDIUM_COUNT = 0
        $LOW_COUNT = 0
        $INFO_COUNT = 0
        $UNDEFINED_COUNT = 0
        $OLDEST_IMAGE_AGE = 0
        $IMAGE_COUNT = 0

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = $SheetName

        $Data = $workbook.Worksheets.Item(1)
        $Data.Name = $SheetName
        $Data.Cells.Item(1, 1) = 'Repo Name'
        $Data.Cells.Item(1, 2) = 'Scan on Push'
        $Data.Cells.Item(1, 3) = 'Latest Image'
        $Data.Cells.Item(1, 4) = 'Image Pushed'
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

        # Insert Data
        # function to add row with data as array parameter
        function Add-Row {
            param(
                [Parameter(Mandatory = $true)]
                [int]$ROW_NUM,
                # add mandadataory parameter for array of data to add to row validating it is not null and not empty and contains 12 items
                [Parameter(Mandatory = $true)]
                [ValidateCount(14, 14)]
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
            $image_list_json = aws ecr describe-images --repository-name $($REPO.repositoryName) --query 'sort_by(imageDetails,& imagePushedAt)[*]' --profile $AWS_PROFILE
            $image_list = $image_list_json | ConvertFrom-Json
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
                    $($REPO.repositoryName), "$($REPO.imageScanningConfiguration.scanOnPush)", 'N/A', 
                    [datetime]::Now, [datetime]::Now, 0, 0, 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', $image_list.count
                )
            }
            # IF IMAGE SCAN STATUS IS NOT COMPLETE
            elseif ($LATEST_IMAGE.imageScanStatus.Status -ne 'COMPLETE') {
                $rdata = @(
                    $($REPO.repositoryName),
                    $($REPO.imageScanningConfiguration.scanOnPush),
                    "$($REPO.repositoryUri):$($LATEST_IMAGE.imageTags[0])",
                    $LATEST_IMAGE.imagePushedAt,
                    $LATEST_IMAGE.imageScanStatus.Status,
                    [math]::Round(([datetime]::Now - [datetime]$LATEST_IMAGE.imagePushedAt).TotalDays),
                    'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', $image_list.count
                )
            } 
            else {
                $rdata = @(
                    $($REPO.repositoryName),
                    $($REPO.imageScanningConfiguration.scanOnPush),
                    "$($REPO.repositoryUri):$($LATEST_IMAGE.imageTags[0])",
                    $LATEST_IMAGE.imagePushedAt,
                    $LATEST_IMAGE.imageScanFindingsSummary.imageScanCompletedAt,
                    # Get days since image was pushed             
                    [math]::Round(([datetime]::Now - [datetime]$LATEST_IMAGE.imagePushedAt).TotalDays),
                    # Get days since image was scanned
                    [math]::Round(([datetime]::Now - [datetime]$LATEST_IMAGE.imageScanFindingsSummary.imageScanCompletedAt).TotalDays),
                    $LATEST_IMAGE.imageScanFindingsSummary.findingSeverityCounts.CRITICAL,
                    $LATEST_IMAGE.imageScanFindingsSummary.findingSeverityCounts.HIGH,
                    $LATEST_IMAGE.imageScanFindingsSummary.findingSeverityCounts.MEDIUM,
                    $LATEST_IMAGE.imageScanFindingsSummary.findingSeverityCounts.LOW,
                    $LATEST_IMAGE.imageScanFindingsSummary.findingSeverityCounts.INFORMATIONAL,
                    $LATEST_IMAGE.imageScanFindingsSummary.findingSeverityCounts.UNDEFINED,
                    $image_list.count
                )
            }
            if ($rdata.count -ne 14) {
                Write-Host 'Array count is not 14, exiting script'
                exit
            }
            # replace empty and null values with N/A
            $rdata = $rdata | ForEach-Object { if ($_ -eq $null -or $_ -eq '') { 0 } else { $_ } }
        
            # Write rdata array to console
            #$rdata | ForEach-Object { Write-Host $_ }
            Add-Row -ROW_NUM $ROW_NUM -ROW_DATA $rdata
            $ROW_NUM++
            $CRIT_COUNT += $LATEST_IMAGE.imageScanFindingsSummary.findingSeverityCounts.CRITICAL
            $HIGH_COUNT += $LATEST_IMAGE.imageScanFindingsSummary.findingSeverityCounts.HIGH
            $MEDIUM_COUNT += $LATEST_IMAGE.imageScanFindingsSummary.findingSeverityCounts.MEDIUM
            $LOW_COUNT += $LATEST_IMAGE.imageScanFindingsSummary.findingSeverityCounts.LOW
            $INFO_COUNT += $LATEST_IMAGE.imageScanFindingsSummary.findingSeverityCounts.INFORMATIONAL
            $UNDEFINED_COUNT += $LATEST_IMAGE.imageScanFindingsSummary.findingSeverityCounts.UNDEFINED
            $IMAGE_COUNT += $image_list.count

            if ($image_list.count -ne 0) {
                $OLDEST_IMAGE_AGE = [math]::Max($OLDEST_IMAGE_AGE, $rdata[5])
            }
            Clear-Host
            Write-Progress -Activity "Processing $($REPO.repositoryName)" -Status "Completed $($ROW_NUM - 2) of $($ECR_REPO_LIST.repositories.count)"  -PercentComplete (($($ROW_NUM - 2) / $($ECR_REPO_LIST.repositories.count)) * 100)
            # Below the percentage complete, show the running count of CRITICAL, HIGH, MEDIUM, LOW, INFO, UNDEFINED findings to console window with CRITICAL in red, HIGH in yellow, MEDIUM in blue, LOW in green, INFO in white, UNDEFINED in white
            Write-Host "`n`n`n`n`n`n`n`n"
            Write-Host "CRITICAL: $CRIT_COUNT" -ForegroundColor Magenta
            Write-Host "HIGH: $HIGH_COUNT" -ForegroundColor Red
            Write-Host "MEDIUM: $MEDIUM_COUNT" -ForegroundColor DarkYellow
            Write-Host "LOW: $LOW_COUNT" -ForegroundColor DarkGreen
            Write-Host "INFO: $INFO_COUNT" -ForegroundColor White
            Write-Host "UNDEFINED: $UNDEFINED_COUNT" -ForegroundColor DarkRed
            Write-Host "OLDEST IMAGE: $OLDEST_IMAGE_AGE" -ForegroundColor DarkMagenta
            Write-Host "TOTAL IMAGES: $IMAGE_COUNT" -ForegroundColor DarkCyan
        }

        # Add a summary sheet that counts the total number of findings by severity with CRITICAL in red, HIGH in yellow, MEDIUM in blue, LOW in green, INFO in white, UNDEFINED in white, also include the average age of the scan/push in days
        $summary = $workbook.Worksheets.Add()
        $summary.Name = "ECR-Summary-$AWS_PROFILE-$PULL_DATE"
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
        $workbook.SaveAs($Path)
        $workbook.Close()
        $excel.Quit()
    }
    Export-ExcelFile -Path ".\ECR-Scan-Report-${AWS_PROFILE}-${PULL_DATE}.xlsx" -SheetName "${AWS_PROFILE}-${PULL_DATE}" -ECR_REPO_LIST ${ECR_REPO_LIST}
}
