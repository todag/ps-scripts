#MIT License
#Copyright (c) 2017 https://github.com/todag

#Permission is hereby granted, free of charge, to any person obtaining a copy
#of this software and associated documentation files (the "Software"), to deal
#in the Software without restriction, including without limitation the rights
#to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#copies of the Software, and to permit persons to whom the Software is
#furnished to do so, subject to the following conditions:
#The above copyright notice and this permission notice shall be included in all
#copies or substantial portions of the Software.

#THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
#SOFTWARE.

<#
.SYNOPSIS
  This scripts collects computer hardware information and name of last logged on user.
  The data is written to the specified active directory attributes on the computer account.
  The data in the attributes will only be updated if the values have changed.

    .PARAMETER <Sleep>
      Script will sleep this long in seconds before starting collection operations

    .PARAMETER <UserInfoAppendString>
        This string will be appended to $userInfoAttribute

    .PARAMETER <UserInfoPrependString>
        This string will be prepended to $userInfoAttribute

.NOTES

  Version:        1.0

  Author:         <https://github.com/todag>

  Creation Date:  <2017-11-11>

  Purpose/Change: Initial script development

.EXAMPLE
  ./ad-collect -Sleep 10 -UserInfoPrependstring "VPN-Logon: "
#>

Param
(
    #Setting this will cause script to sleep for x seconds.
    [Parameter(Mandatory=$false)]
    [int]$Sleep = 0,

    #Setting this will cause string to be appended to $userInfoAttribute
    #For example, if script is called after loggin on through VPN, it can append (' (VPN-Logon)')
    [Parameter(Mandatory=$false)]
    [string]$UserInfoAppendString,

    #Setting this will cause string to be prepended to $userInfoAttribute
    #For example, if script is called after loggin on through VPN, it can prepend ('VPN-Logon: ')
    [Parameter(Mandatory=$false)]
    [string]$UserInfoPrependString
)
Set-StrictMode -Version 2.0
$VerbosePreference="Continue"

#------------------------------------------------------------ Script settings ------------------------------------------------------------

# Will write computer info (manufacturer, model and serial number) to this attribute
$computerInfoAttribute = "employeeType"

# Will write full name and samAccountName of logged on user to this attribute
$userInfoAttribute = "info"

# Set to $true to log to eventlog
$logToEventLog = $true

#Set name of eventlog source
$eventLogSource = "AD COLLECT"

# If this is set to true, it will timestamp the username entry in info attribute. This will cause it to be updated on every run of the script
# If this is set to false, it will not timestamp the username entry in info attribute and only update the $userInfoAttribute attribute if a different user has logged on
$TimeStampUserInfo = $false

# Set datetime format. Only used if $TimeStampUserInfo = $true.
# If set to a short format, eg. "yyyy-mm-dd" attribute will only get updated once a day (unless a new user logs on)
$DateTimeFormat = "yyyy-MM-dd HH:mm"

#-----------------------------------------------------------------------------------------------------------------------------------------

#Used for event logging, whether to log Information or Error
$scriptFailed = $false

#Check if event log source exists, if not, create it. Script might need elevated permissions to be able to add event log source.
#If logging to eventlog is enabled and this fails, script will terminate.
if($logToEventLog)
{
    try
    {
        Write-Verbose("Checking if event log source '" + $eventLogSource + "' exists")
        if ([System.Diagnostics.EventLog]::SourceExists($eventLogSource) -eq $False)
        {
            Write-Verbose("Source does not exist, creating new event log source named " + $eventLogSource)
            New-EventLog –LogName Application –Source $eventLogSource
        }
        else
        {
            Write-Verbose("Event log source '" + $eventLogSource + "' exists")
        }
    }
    catch
    {
        Write-Warning $_.Exception.Message
        Write-Warning ("Event logging enabled, but something went wrong exiting script")
        Exit
    }
}

#This function logs event to the $script:log string variable.
#If $logToEventLog is set to $true, this string will be logged to eventlog.
$script:log = ""
Function Write-Log ($logString)
{
    Write-Verbose $logString
    $script:log = $script:log + "`nLog: " + (Get-Date).ToString("HH:mm:ss") + " " + $logString
}

if($Sleep -gt 0)
{
    Write-Log ("Sleeping " + $Sleep + " seconds...")
    Start-Sleep -s $Sleep
}

#To calculate script runtime
$scriptStartTime = (Get-Date)

#Gathers data to be written to the $computerInfoAttribute
Function Get-ComputerInfo
{
    try
    {
        [string]$result = $null
        Write-Log "Computer scan: Getting computer info..."
        #Get WMI data
        $computerManufacturer = (Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue).Manufacturer
        $computerModel = (Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue).Model
        $computerSerialNumber = (Get-WMIObject Win32_BIOS -ErrorAction SilentlyContinue).SerialNumber
        Write-Log ("Computer scan: Manufacturer: " + $computerManufacturer)
        Write-Log ("Computer scan: Model: " + $computerModel)
        Write-Log ("Computer scan: Serial#: " + $computerSerialNumber)

        #Some manufacturers have a not so descriptive Model ID (like Lenovo).
        #This will translate the model id to a more readable format.
        Write-Log ("Computer scan: Checking if model ID needs conversion...")
        $ModelTranslations=@{
            "20EV000TMS" = "Thinkpad E560";
            "20J4000LMX" = "Thinkpad L470";
            "20J4003YMX" = "Thinkpad L470";
            "7033W5U" = "ThinkCentre M91p";
            "20A7008LMS" = "ThinkPad Carbon X1 G2";
            "20BS003HMS" = "ThinkPad Carbon X1 G3";
            "20FB002WMX" = "ThinkPad Carbon X1 G4";
            "20HR0021MX" = "ThinkPad Carbon X1 G5";
            "20AL007YMS" = "ThinkPad X240";
            "2356FN5" = "ThinkPad T430s";
            "4480B2G" = "ThinkCentre M91p";
            "5536W6B" = "ThinkCentre M90p";
            "10FD001LMX" = "ThinkCentre M900";
            "80EW" = "B50-80";
            "80H8" = "M30-70";
            "80KX" = "E31-70";
            "80LT" = "B50-80";
            "80MR" = "B70-80";
            "80TL" = "V110";
        }

        if($ModelTranslations.ContainsKey($computerModel))
        {
            Write-Log ("Computer scan: Converting '" + $computerModel + "' to '" + $ModelTranslations.Item($computerModel) + "'")
            $computerModel = $ModelTranslations.Item($computerModel)
        }
        else
        {
            Write-Log "Computer scan: No model ID conversion needed"
        }

        #Structure up the value
        $result = $computerManufacturer + " " + $computerModel + " Serial# " + $computerSerialNumber
        Write-Log ("Computer scan: Returning '" + $result + "'")
        return $result.ToString()
    }
    catch
    {
        Write-Log ("Computer scan: Unable to get computer info, exception: " + $_.Exception.Message)
        return $null
    }

}

# Gathers data to be written to the $userInfoAttribute Currently cannot retrieve user info for users logged on through RDP. Should fix that...
Function Get-UserInfo
{
    try
    {
        Write-Log ("User scan: Getting logged on user...")
        #Need to use wmi here since if script is running as SYSTEM, we cannot simply use $env:username
        [string]$userName =  (Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue).Username
        if([string]::IsNullOrEmpty($userName))
        {
            Write-Log ("User scan: Could not find any user logged on locally")
            return $null
        }

        Write-Log ("User scan: Got logged on user with wmi: " + $userName)
        #The Get-WmiObject method to retrieve username returns it in domain\username format. This will strip the domain\ part from it...
        $userName = $userName.Substring($userName.IndexOf("\") + 1)
        Write-Log ("User scan: Username from wmi stripped to: " + $userName)

        #Do an AD search and retrieve the full name of the user. Is there really no way to get full name without hitting AD?
        Write-Log ("User scan: Getting full name of user from AD")
        $adUserSearcher = New-Object System.DirectoryServices.DirectorySearcher
        $adUserSearcher.Filter = ("(&(objectCategory=User)(samAccountName=" + $userName + "))")
        $adUserSearcher.PropertiesToLoad.Add("displayName") | Out-Null
        $adUserEntry = $adUserSearcher.FindOne().GetDirectoryEntry()
        $fullName = $adUserEntry.displayName.ToString()
        $adUserEntry.Close()

        Write-Log ("User scan: Got full name from AD: " + $fullName)

        if($TimeStampUserInfo -eq $true)
        {
            [string]$result = $fullName + " (" + $userName + ") " + (Get-Date).ToString($DateTimeFormat)
        }
        else
        {
            [string]$result = $fullName + " (" + $userName + ")"
        }

        Write-Log ("User scan: Returning '" + $result + "'")
        return $result
    }
    catch
    {
        Write-Log ("User scan: Unable to get username, exception: " + $_.Exception.Message)
        if($adUserEntry -ne $null)
        {
            $adUserEntry.Close()
        }
        return $null
    }
}

# This is where the action begins...
# This will get existing values from AD. It will then compare existing values with
# values retrieved through Get-ComputerInfo and Get-UserInfo. If values don't
# match, it will update the attributes with the new values.
try
{
    #Get existing data in the two attributes.
    Write-Log ("Opening Directory Entry for " + $env:computername)
    $adComputerSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $adComputerSearcher.Filter = "(&(objectCategory=Computer)(SamAccountname=$($env:computerName)`$))"
    $adComputerSearcher.PropertiesToLoad.Add($computerInfoAttribute) | Out-Null
    $adComputerSearcher.PropertiesToLoad.Add($userInfoAttribute) | Out-Null
    $adComputerEntry = $adComputerSearcher.FindOne().GetDirectoryEntry()

    [string]$currentComputerInfoValue = $adComputerEntry.Properties[$computerInfoAttribute].ToString()
    [string]$currentUserInfoValue = $adComputerEntry.Properties[$userInfoAttribute].ToString()

    #Scan computer for new values
    [string]$newComputerInfoValue = Get-ComputerInfo
    [string]$newUserInfoValue = Get-UserInfo

    #Add prepend string (if any)

    if([string]::IsNullOrEmpty($UserInfoPrependString) -and [string]::IsNullOrEmpty($UserInfoAppendString))
    {
        Write-Log "Prepend/Append: Noting to append or prepend to userInfo"
    }
    else
    {
        if(![string]::IsNullOrEmpty($UserInfoPrependString))
        {
            Write-Log ("Prepend/Append: Prepending '" + $UserInfoPrependString + "' to userInfo")
            $newUserInfoValue = ($UserInfoPrependString + $newUserInfoValue)
        }

        #Add append string (if any)
        if(![string]::IsNullOrEmpty($UserInfoAppendString))
        {
            Write-Log ("Prepend/Append: Appending '" + $UserInfoAppendString + "' to userInfo")
            $newUserInfoValue =  $newUserInfoValue + $UserInfoAppendString
        }
        Write-Log ("Prepend/Append: userInfo is now '" + $newUserInfoValue + "'")
    }

    $commitChanges = $false
    [int]$changeCount = 0
    #Check if computer model has changed. This will likely not happen more than once per computer (since values retrieved from Get-ComputerInfo is quite static)
    if([string]::IsNullOrEmpty($newComputerInfoValue))
    {
        Write-Log ("Commit: No changes to attribute '" + $computerInfoAttribute + "', scanned value '" + $newComputerInfoValue + "' is null or empty, existing value is '" + $currentComputerInfoValue + "'")
    }
    elseif($newComputerInfoValue -eq $currentComputerInfoValue)
    {
        Write-Log ("Commit: No changes to attribute '" + $computerInfoAttribute + ", scanned value '" + $newComputerInfoValue + "' matches existing value '" + $currentComputerInfoValue + "'")
    }
    else
    {
        Write-Log ("Commit: Changing data on attribute '" + $computerInfoAttribute + "', with existing value '" + $currentComputerInfoValue + "' to scanned value '" + $newComputerInfoValue + "'")
        $adComputerEntry.Properties[$computerInfoAttribute].Value = $newComputerInfoValue.ToString()
        $commitChanges = $true
        $changeCount++
        Write-Log ($computerInfoAttribute + "-attribute changes will be commited")
    }

    #Check if currently logged in user is same as previously logged in user. If not update the value.
    #If $TimeStampUserInfo is set to $true, it will update the attribute regardless.
    if([string]::IsNullOrEmpty($newUserInfoValue))
    {
        Write-Log ("Commit: No changes to attribute '" + $userInfoAttribute + "', scanned value is null or empty, existing value is '" + $currentUserInfoValue + "'")
    }
    elseif($newUserInfoValue -eq $currentUserInfoValue)
    {
        Write-Log ("Commit: No changes to attribute '" + $userInfoAttribute + "', scanned value '" + $newUserInfoValue + "' matches existing value '" + $currentUserInfoValue + "'")
    }
    else
    {
        Write-Log ("Commit: Changing data on attribute '" + $userInfoAttribute + "', with existing value '" + $currentUserInfoValue + "' to scanned value '" + $newUserInfoValue + "'")
        $adComputerEntry.Properties[$userInfoAttribute].Value = $newUserInfoValue.ToString()
        $commitChanges = $true
        $changeCount++
        Write-Log ($userInfoAttribute + "-attribute changes will be commited")
    }

    #Commit changes to AD
    if($commitChanges -eq $true)
    {
        Write-Log ("Commit: ** Commiting " + $changeCount.ToString() + " changes to AD **")
        $adComputerEntry.CommitChanges()
    }
    else
    {
        Write-Log "Commit: ** No changes commited to AD **"
    }
    $adComputerEntry.Close()
}
catch
{
    Write-Log ("Script failed with exeption: " + $_.Exception.Message)
    $scriptFailed = $true
}
finally
{
    if($adComputerEntry -ne $null)
    {
        $adComputerEntry.Close()
    }

    #Calculate and log script runtime
    $scriptEndTime = (Get-Date)
    $scriptRunTimeMS = ($scriptEndTime - $scriptStartTime).TotalMilliseconds
    $scriptRunTimeS = ($scriptEndTime - $scriptStartTime).Seconds
    Write-Log ("Total time elapsed: " +  + $scriptRunTimeS + " seconds, or " + $scriptRunTimeMS + " milliseconds (Excluding `$Sleep time)")

    if($scriptFailed)
    {
        if($logToEventLog)
        {
            Write-EventLog –LogName Application –Source $eventLogSource –EntryType Error –EventID 1 –Message $script:log
            Write-Verbose("Script finished with errors, check event log for details")
        }
        else
        {
            Write-Verbose("Script finished with errors")
        }
    }
    else
    {
        if($logToEventLog)
        {
            Write-EventLog –LogName Application –Source $eventLogSource –EntryType Information –EventID 1 –Message $script:log
            Write-Verbose("Script finished, check event log for details")
        }
        else
        {
            Write-Verbose("Script finished")
        }
    }
}