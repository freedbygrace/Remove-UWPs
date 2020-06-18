#Requires -Version 3

<#
    .SYNOPSIS
    Removes an app package (.appx) from a Windows image using a regular expression that can be modified from the command line without changing the script code.
          
    .DESCRIPTION
    Slightly more detailed description of what your function does
          
    .PARAMETER ImagePath
    The directory path to the mounted windows image or the drive letter of a windows image that has been expanded onto a disk while in WindowsPE.

    .PARAMETER PackagesToRemove
    A valid regular expression that matches the DisplayName(s) of the universal windows applications you wish to deprovision.

    .PARAMETER LogDir
    A valid folder path. If the folder does not exist, it will be created. This parameter can also be specified by the alias "LogPath".

    .PARAMETER ContinueOnError
    Ignore failures.
          
    .EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -NoProfile -NoLogo -File "%FolderPathContainingScript%\%ScriptName%.ps1" -ImagePath "%OSDisk%\\" -PackagesToRemove '.*Xbox.*|.*Alarms.*|.*Bing.*|.*Skype.*|.*Communi.*'

    .NOTES
    Any useful tidbits
          
    .LINK
    https://docs.microsoft.com/en-us/powershell/module/dism/remove-appxprovisionedpackage?view=win10-ps#:~:text=The%20Remove%2DAppxProvisionedPackage%20cmdlet%20removes,removed%20from%20existing%20user%20accounts.
#>

[CmdletBinding()]
    Param
        (        	     
            [Parameter(Mandatory=$False)]
            [ValidateNotNullOrEmpty()]
            [ValidateScript({($_ -imatch '^[a-zA-Z][\:]\\{1,}$') -or ($_ -imatch '^[a-zA-Z][\:]\\.*?[^\\]$')})]
            [System.IO.DirectoryInfo]$ImagePath,
            
            [Parameter(Mandatory=$False)]
            [ValidateNotNullOrEmpty()]
            [Regex]$PackagesToRemove = '.*Xbox.*|.*Alarms.*|.*Bing.*|.*Skype.*|.*Communi.*|.*Zune.*|.*Solit.*|.*Games.*|.*3D.*|.*Camera.*|.*Phone.*|.*Soundrec.*|.*Xboxapp.*|.*Feedback.*|.*Message.*|.*Messaging.*|.*Gethelp.*|.*Mobile.*|.*Mixed.*|.*Microsoft\.MicrosoftOfficeHub.*|.*OneNote.*|.*Microsoft\.People.*|.*Windows.*Defender.*|.*Movies.*|.*Microsoft\.ZuneMusic.*|.*Microsoft\.Windows\.HolographicFirstRun.*|.*Microsoft\.Oneconnect.*|.*Microsoft\.WindowsStore.*|.*Microsoft\.Getstarted.*|.*Microsoft\.MSPaint.*|.*Microsoft\.ZuneVideo.*|.*Microsoft\.PPIProjection.*|.*Microsoft\.WindowsMaps.*',
            
            [Parameter(Mandatory=$False)]
            [ValidateNotNullOrEmpty()]
            [ValidateScript({($_ -imatch '^[a-zA-Z][\:]\\.*?[^\\]$') -or ($_ -imatch "^\\(?:\\[^<>:`"/\\|?*]+)+$")})]
            [Alias('LogPath')]
            [System.IO.DirectoryInfo]$LogDir,
            
            [Parameter(Mandatory=$False)]
            [Switch]$ContinueOnError
        )

#Define Default Action Preferences
    $Script:DebugPreference = 'SilentlyContinue'
    $Script:ErrorActionPreference = 'Stop'
    $Script:VerbosePreference = 'SilentlyContinue'
    $Script:WarningPreference = 'Continue'
    $Script:ConfirmPreference = 'None'
    
#Load WMI Classes
  $Baseboard = Get-WmiObject -Namespace "root\CIMv2" -Class "Win32_Baseboard" -Property * | Select-Object -Property *
  $Bios = Get-WmiObject -Namespace "root\CIMv2" -Class "Win32_Bios" -Property * | Select-Object -Property *
  $ComputerSystem = Get-WmiObject -Namespace "root\CIMv2" -Class "Win32_ComputerSystem" -Property * | Select-Object -Property *
  $OperatingSystem = Get-WmiObject -Namespace "root\CIMv2" -Class "Win32_OperatingSystem" -Property * | Select-Object -Property *

#Retrieve property values
  $OSArchitecture = $($OperatingSystem.OSArchitecture).Replace("-bit", "").Replace("32", "86").Insert(0,"x").ToUpper()

#Define variable(s)
  $DateTimeLogFormat = 'dddd, MMMM dd, yyyy hh:mm:ss tt'  ###Monday, January 01, 2019 10:15:34 AM###
  [ScriptBlock]$GetCurrentDateTimeLogFormat = {(Get-Date).ToString($DateTimeLogFormat)}
  $DateTimeFileFormat = 'yyyyMMdd_hhmmsstt'  ###20190403_115354AM###
  [ScriptBlock]$GetCurrentDateTimeFileFormat = {(Get-Date).ToString($DateTimeFileFormat)}
  [System.IO.FileInfo]$ScriptPath = "$($MyInvocation.MyCommand.Definition)"
  [System.IO.DirectoryInfo]$ScriptDirectory = "$($ScriptPath.Directory.FullName)"
  [System.IO.DirectoryInfo]$FunctionsDirectory = "$($ScriptDirectory.FullName)\Functions"
  [System.IO.DirectoryInfo]$ModulesDirectory = "$($ScriptDirectory.FullName)\Modules"
  [System.IO.DirectoryInfo]$ToolsDirectory = "$($ScriptDirectory.FullName)\Tools"
  [System.IO.DirectoryInfo]$ToolsDirectory_OSAll = "$($ToolsDirectory.FullName)\All"
  [System.IO.DirectoryInfo]$ToolsDirectory_OSArchSpecific = "$($ToolsDirectory.FullName)\$($OSArchitecture)"
  $IsWindowsPE = Test-Path -Path 'HKLM:\SYSTEM\ControlSet001\Control\MiniNT' -ErrorAction SilentlyContinue
	
#Log task sequence variables if debug mode is enabled within the task sequence
  Try
    {
        [System.__ComObject]$TSEnvironment = New-Object -ComObject "Microsoft.SMS.TSEnvironment"
              
        If ($TSEnvironment -ine $Null)
          {
              $IsRunningTaskSequence = $True
          }
    }
  Catch
    {
        $IsRunningTaskSequence = $False
    }

#Determine the default logging path if the parameter is not specified and is not assigned a default value
  If (($PSBoundParameters.ContainsKey('LogDir') -eq $False) -and ($LogDir -ieq $Null))
    {
        If ($IsRunningTaskSequence -eq $True)
          {
              [String]$_SMSTSLogPath = "$($TSEnvironment.Value('_SMSTSLogPath'))"
                    
              If ([String]::IsNullOrEmpty($_SMSTSLogPath) -eq $False)
                {
                    [System.IO.DirectoryInfo]$TSLogDirectory = "$($_SMSTSLogPath)"
                }
              Else
                {
                    [System.IO.DirectoryInfo]$TSLogDirectory = "$($Env:Windir)\Temp\SMSTSLog"
                }
                     
              [System.IO.DirectoryInfo]$LogDir = "$($TSLogDirectory.FullName)\$($ScriptPath.BaseName)"
          }
        ElseIf ($IsRunningTaskSequence -eq $False)
          {
              [System.IO.DirectoryInfo]$LogDir = "$($Env:Windir)\Logs\Software\$($ScriptPath.BaseName)"
          }
    }

#Start transcripting (Logging)
  Try
    {
        [System.IO.FileInfo]$ScriptLogPath = "$($LogDir.FullName)\$($ScriptPath.BaseName)_$($GetCurrentDateTimeFileFormat.Invoke()).log"
        If ($ScriptLogPath.Directory.Exists -eq $False) {[Void][System.IO.Directory]::CreateDirectory($ScriptLogPath.Directory.FullName)}
        Start-Transcript -Path "$($ScriptLogPath.FullName)" -IncludeInvocationHeader -Force -Verbose
    }
  Catch
    {
        If ([String]::IsNullOrEmpty($_.Exception.Message)) {$ExceptionMessage = "$($_.Exception.Errors.Message)"} Else {$ExceptionMessage = "$($_.Exception.Message)"}
          
        $ErrorMessage = "[Error Message: $($ExceptionMessage)][ScriptName: $($_.InvocationInfo.ScriptName)][Line Number: $($_.InvocationInfo.ScriptLineNumber)][Line Position: $($_.InvocationInfo.OffsetInLine)][Code: $($_.InvocationInfo.Line.Trim())]"
        Write-Error -Message "$($ErrorMessage)"
    }

#Log any useful information
  $LogMessage = "IsWindowsPE = $($IsWindowsPE.ToString())"
  Write-Verbose -Message "$($LogMessage)" -Verbose

  $LogMessage = "Script Path = $($ScriptPath.FullName)"
  Write-Verbose -Message "$($LogMessage)" -Verbose

  $LogMessage = "Packages To Remove = $($PackagesToRemove.ToString())"
  Write-Verbose -Message "$($LogMessage)" -Verbose

  $DirectoryVariables = Get-Variable | Where-Object {($_.Value -ine $Null) -and (($_.Value -is [System.IO.DirectoryInfo]) -or ($_.Value -is [System.IO.DirectoryInfo[]]))}
  
  ForEach ($DirectoryVariable In $DirectoryVariables)
    {
        $LogMessage = "$($DirectoryVariable.Name) = $($DirectoryVariable.Value.FullName -Join ', ')"
        Write-Verbose -Message "$($LogMessage)" -Verbose
    }

#region Import Dependency Modules
$Modules = Get-Module -Name "$($ModulesDirectory.FullName)\*" -ListAvailable -ErrorAction Stop 

$ModuleGroups = $Modules | Group-Object -Property @('Name')

ForEach ($ModuleGroup In $ModuleGroups)
  {
      $LatestModuleVersion = $ModuleGroup.Group | Sort-Object -Property @('Version') -Descending | Select-Object -First 1
      
      If ($LatestModuleVersion -ine $Null)
        {
            $LogMessage = "Attempting to import dependency powershell module `"$($LatestModuleVersion.Name) [Version: $($LatestModuleVersion.Version.ToString())]`". Please Wait..."
            Write-Verbose -Message "$($LogMessage)" -Verbose
            Import-Module -Name "$($LatestModuleVersion.Path)" -Global -DisableNameChecking -Force -ErrorAction Stop
        }
  }
#endregion

#region Dot Source Dependency Scripts
#Dot source any additional script(s) from the functions directory. This will provide flexibility to add additional functions without adding complexity to the main script and to maintain function consistency.
  Try
    {
        If ($FunctionsDirectory.Exists -eq $True)
          {
              [String[]]$AdditionalFunctionsFilter = "*.ps1"
        
              $AdditionalFunctionsToImport = Get-ChildItem -Path "$($FunctionsDirectory.FullName)" -Include ($AdditionalFunctionsFilter) -Recurse -Force | Where-Object {($_ -is [System.IO.FileInfo])}
        
              $AdditionalFunctionsToImportCount = $AdditionalFunctionsToImport | Measure-Object | Select-Object -ExpandProperty Count
        
              If ($AdditionalFunctionsToImportCount -gt 0)
                {                    
                    ForEach ($AdditionalFunctionToImport In $AdditionalFunctionsToImport)
                      {
                          Try
                            {
                                $LogMessage = "Attempting to dot source dependency script `"$($AdditionalFunctionToImport.Name)`". Please Wait...`r`n`r`nScript Path: `"$($AdditionalFunctionToImport.FullName)`""
                                Write-Verbose -Message "$($LogMessage)" -Verbose
                          
                                . "$($AdditionalFunctionToImport.FullName)"
                            }
                          Catch
                            {
                                $ErrorMessage = "[Error Message: $($_.Exception.Message)]`r`n`r`n[ScriptName: $($_.InvocationInfo.ScriptName)]`r`n[Line Number: $($_.InvocationInfo.ScriptLineNumber)]`r`n[Line Position: $($_.InvocationInfo.OffsetInLine)]`r`n[Code: $($_.InvocationInfo.Line.Trim())]"
                                Write-Error -Message "$($ErrorMessage)" -Verbose
                            }
                      }
                }
          }
    }
  Catch
    {
        $ErrorMessage = "[Error Message: $($_.Exception.Message)]`r`n`r`n[ScriptName: $($_.InvocationInfo.ScriptName)]`r`n[Line Number: $($_.InvocationInfo.ScriptLineNumber)]`r`n[Line Position: $($_.InvocationInfo.OffsetInLine)]`r`n[Code: $($_.InvocationInfo.Line.Trim())]"
        Write-Error -Message "$($ErrorMessage)" -Verbose            
    }
#endregion

#Perform script action(s)
  Try
    {                          
        #Tasks defined within this block will only execute if a task sequence is running
          If (($IsRunningTaskSequence -eq $True))
            {
                If (($PSBoundParameters.ContainsKey('ImagePath') -eq $False) -and ($ImagePath -eq $Null))
                  {
                      [System.IO.DirectoryInfo]$ImagePath = "$($TSEnvironment.Value('OSDisk'))\"
                  }
            }

        #Tasks defined here will execute whether only if a task sequence is not running
          If ($IsRunningTaskSequence -eq $False)
            {
                $WarningMessage = "There is no task sequence running.`r`n"
                Write-Warning -Message "$($WarningMessage)" -Verbose
            }
        #Remove Metro Application(s)
          If ($ImagePath -ine $Null)
            {             
                  $LogMessage = "ImagePath = $($ImagePath.FullName)"
                  Write-Verbose -Message "$($LogMessage)" -Verbose
              
                  $AppXProvisionedPackages = (Get-AppxProvisionedPackage -Path "$($ImagePath.FullName)" | Select-Object * | Sort-Object DisplayName)
                    
                  $AppXProvisionedPackagesToRemove = $AppXProvisionedPackages | Where-Object {($_.DisplayName -imatch $PackagesToRemove.ToString())}
                    
                  $AppXProvisionedPackagesToRemoveCount = $AppXProvisionedPackagesToRemove | Measure-Object | Select-Object -ExpandProperty Count

                  $LogMessage = "Found $($AppXProvisionedPackagesToRemoveCount.ToString()) Appx provisioned package(s) to remove."
                  Write-Verbose -Message "$($LogMessage)" -Verbose
    
                  $Counter = 1
                    
                  ForEach ($AppXProvisionedPackageToRemove In $AppXProvisionedPackagesToRemove)
                      {
                          Try
                            {
                                [Int]$ProgressID = 1
                                [String]$ActivityMessage = "Remove-AppxProvisionedPackage $($AppXProvisionedPackageToRemove.DisplayName) [Version: $($AppXProvisionedPackageToRemove.Version)]"
                                [String]$StatusMessage = "Attempting to remove $($AppXProvisionedPackageToRemove.DisplayName) $($AppXProvisionedPackageToRemove.Version) ($($Counter.ToString()) of $($AppXProvisionedPackagesToRemoveCount.ToString()))"
                                [Int]$PercentComplete = (($Counter / $AppXProvisionedPackagesToRemoveCount) * 100)

                                $LogMessage = "$($StatusMessage). Please Wait..."
                                Write-Verbose -Message "$($LogMessage)" -Verbose
                              
                                Write-Progress -ID ($ProgressID) -Activity ($ActivityMessage) -Status ($StatusMessage) -PercentComplete ($PercentComplete)

                                [System.IO.FileInfo]$AppXProvisionedPackageToRemoveLogPath = "$($LogDir.FullName)\Remove-AppxProvisionedPackage\$($AppXProvisionedPackageToRemove.DisplayName).log"

                                If ($AppXProvisionedPackageToRemoveLogPath.Directory.Exists -eq $False) {$Null = [System.IO.Directory]::CreateDirectory($AppXProvisionedPackageToRemoveLogPath.Directory.FullName)}

                                $LogMessage = "Log Path = `"$($AppXProvisionedPackageToRemoveLogPath.FullName)`""
                                Write-Verbose -Message "$($LogMessage)" -Verbose
                              
                                $RemoveAppXProvisionedPackage = Remove-AppxProvisionedPackage -PackageName "$($AppXProvisionedPackageToRemove.PackageName)" -Path "$($ImagePath.FullName)" -LogPath "$($AppXProvisionedPackageToRemoveLogPath.FullName)" -LogLevel 'Errors' -Verbose

                                If ($? -eq $True)
                                  {
                                      $LogMessage = "Removal of the `"$($AppXProvisionedPackageToRemove.DisplayName)`" AppX provisioned package was successful!"
                                      Write-Verbose -Message "$($LogMessage)" -Verbose
                                  }
                                ElseIf ($? -eq $False)
                                  {
                                      $LogMessage = "Removal of the `"$($AppXProvisionedPackageToRemove.DisplayName)`" AppX provisioned package was unsuccessful!"
                                      Write-Error -Message "$($LogMessage)" -Verbose
                                  }
                            }
                          Catch
                            {
                                If ([String]::IsNullOrEmpty($_.Exception.Message)) {$ExceptionMessage = "$($_.Exception.Errors.Message)"} Else {$ExceptionMessage = "$($_.Exception.Message)"}
          
                                $ErrorMessage = "[Error Message: $($ExceptionMessage)][ScriptName: $($_.InvocationInfo.ScriptName)][Line Number: $($_.InvocationInfo.ScriptLineNumber)][Line Position: $($_.InvocationInfo.OffsetInLine)][Code: $($_.InvocationInfo.Line.Trim())]"
                                Write-Error -Message "$($ErrorMessage)" -Verbose 
                            }
                              
                          $Counter++
                      }
            }
          Else
            {
                $ErrorMessage = "[Error Message: A valid image path was either not specified or contains an invalid value.]"
                Throw "$($ErrorMessage)"
            }
    
          #Stop transcripting (Logging)
            Try
              {
                  Stop-Transcript -Verbose
              }
            Catch
              {
                  If ([String]::IsNullOrEmpty($_.Exception.Message)) {$ExceptionMessage = "$($_.Exception.Errors.Message)"} Else {$ExceptionMessage = "$($_.Exception.Message)"}
          
                  $ErrorMessage = "[Error Message: $($ExceptionMessage)][ScriptName: $($_.InvocationInfo.ScriptName)][Line Number: $($_.InvocationInfo.ScriptLineNumber)][Line Position: $($_.InvocationInfo.OffsetInLine)][Code: $($_.InvocationInfo.Line.Trim())]"
                  Write-Error -Message "$($ErrorMessage)"
              }
    }
  Catch
    {
        If ([String]::IsNullOrEmpty($_.Exception.Message)) {$ExceptionMessage = "$($_.Exception.Errors.Message -Join "`r`n`r`n")"} Else {$ExceptionMessage = "$($_.Exception.Message)"}
          
        $ErrorMessage = "[Error Message: $($ExceptionMessage)]`r`n`r`n[ScriptName: $($_.InvocationInfo.ScriptName)]`r`n[Line Number: $($_.InvocationInfo.ScriptLineNumber)]`r`n[Line Position: $($_.InvocationInfo.OffsetInLine)]`r`n[Code: $($_.InvocationInfo.Line.Trim())]`r`n"
        Write-Error -Message "$($ErrorMessage)"
        
        Stop-Transcript -Verbose
        
        If ($ContinueOnError.IsPresent -eq $False)
          {
              [System.Environment]::Exit(50)
          }
    }