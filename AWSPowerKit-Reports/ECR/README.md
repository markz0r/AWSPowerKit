# AWSPowerKit-ReportsECR

# This is a simple example of how to use the AWSPowerKit to generate a report of the ECR repositories in your account.

```powershell
git clone https://github.com/markz0r/AWSPowerKit.git
cd .\AWSPowerKit\AWSPowerKit-Reports\ECR
git pull
# Import the AWSPowerKit module
Import-Module AWSPowerKit -Force # -Force is optional, but it ensures that the latest version of the module is loaded if I forget to increment the version number.
# Run a report
Get-ECRReport -AWS_PROFILE "AdministratorAccess-99999999999" -SIMPLE "marks-webapp-aws-acc"

# if you dont provide arguments, you will be prompted for them interactively, with a list of available profiles.
Get-ECRReport
```
