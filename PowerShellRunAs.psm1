Function New-LocalUserAccount {
<#
.SYNOPSIS  
		Creates a local user account
    Credit: http://blogs.technet.com/b/heyscriptingguy/archive/2010/11/23/use-powershell-to-create-local-user-accounts.aspx

.DESCRIPTION  
		Description goes here

.LINK  
    http://link.com
                
.NOTES  
	Version:
			0.1
  
	Author/Copyright:	 
			Copyright Tom Arbuthnot - All Rights Reserved
    
	Email/Blog/Twitter:	
			tom@tomarbuthnot.com tomtalks.uk @tomarbuthnot
    
	Disclaimer:   	
			THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
			OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
			While these scripts are tested and working in my environment, it is recommended 
			that you test these scripts in a test environment before using in your production 
			environment. Tom Arbuthnot further disclaims all implied warranties including, 
			without limitation, any implied warranties of merchantability or of fitness for 
			a particular purpose. The entire risk arising out of the use or performance of 
			this script and documentation remains with you. In no event shall Tom Arbuthnot, 
			its authors, or anyone else involved in the creation, production, or delivery of 
			this script/tool be liable for any damages whatsoever (including, without limitation, 
			damages for loss of business profits, business interruption, loss of business 
			information, or other pecuniary loss) arising out of the use of or inability to use 
			the sample scripts or documentation, even if Tom Arbuthnot has been advised of 
			the possibility of such damages.
    
     
	Acknowledgements: 	
    
	Assumptions:	      
			ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
    
	Limitations:		  										
    		
	Ideas/Wish list:
			Detects loops based on number of log lines, this can vary, should detect based on timestamp of line 
    
	Rights Required:	

	Known issues:	


.EXAMPLE
		Function-Template
 
		Description
		-----------
		Returns Objects

.EXAMPLE
		Function-Template -Param1
 
		Description
		-----------
		Actions Param1

# Parameters

.PARAMETER Param1
		Param1 description

.PARAMETER Param2
		Param2 Description
		

#>
  
  
  #############################################################
  # Param Block
  #############################################################
  
  # Sets that -Whatif and -Confirm should be allowed
  [cmdletbinding(SupportsShouldProcess=$true)]
  
  Param 	(
    [Parameter(Mandatory=$false,
    HelpMessage='Computer')]
    $Computer = 'localhost',
    
    
    [Parameter(Mandatory=$false,
    HelpMessage='UserName')]
    $UserName,
    
    [Parameter(Mandatory=$false,
    HelpMessage='Password')]
    $Password,
    
    [Parameter(Mandatory=$false,
    HelpMessage='Error Log location, default C:\<Command Name>_ErrorLog.txt')]
    [string]$ErrorLog = "c:\$($myinvocation.mycommand)_ErrorLog.txt",
    [switch]$LogErrors
    
  ) #Close Parameters
  
  
  #############################################################
  # Begin Block
  #############################################################
  
  Begin 	{
    Write-Verbose "Starting $($myinvocation.mycommand)"
    Write-Verbose "Error log will be $ErrorLog"
    
    # Script Level Variable to Stop Execution if there is an issue with any stage of the script
    $script:EverythingOK = $true
    
    
  } #Close Function Begin Block
  
  #############################################################
  # Process Block
  #############################################################
  
  Process {
    
    # First Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
        
        
        if(!$UserName -or !$Password)
        
        {
          
          $(Throw 'A value for $user and $password is required.')
          
          
        }
        
        
        
        $objOu = [ADSI]"WinNT://$computer"
        
        $objUser = $objOU.Create('User', $UserName)
        
        $objUser.setpassword($Password)
        
        $objUser.SetInfo()
        
        $objUser.description = 'User Created with PowerShell'
        
        $objUser.SetInfo()
        
        
      } # Close Try Block
      
      Catch 	{ErrorCatch-Action} # Close Catch Block
      
      
    } # Close If EverthingOK Block 1
    
    #############################################################
    # Next Script Action or Try,Catch Block
    #############################################################
    
    # Second Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
        # Code Goes here
        
        
      } # Close Try Block
      
      Catch 	{ErrorCatch-Action} # Close Catch Block
      
      
    } # Close If EverthingOK Block 2
    
    
  } #Close Function Process Block 
  
  #############################################################
  # End Block
  #############################################################
  
  End 	{
    Write-Verbose "Ending $($myinvocation.mycommand)"
  } #Close Function End Block
  
} #End Function

Function Add-LocalUserToLocalGroup {
<#
.SYNOPSIS  
		Creates a local user account
    Credit: http://blogs.technet.com/b/heyscriptingguy/archive/2010/11/25/use-powershell-to-add-local-users-to-local-groups.aspx

.DESCRIPTION  
		Description goes here

.LINK  
    http://link.com
                
.NOTES  
	Version:
			0.1
  
	Author/Copyright:	 
			Copyright Tom Arbuthnot - All Rights Reserved
    
	Email/Blog/Twitter:	
			tom@tomarbuthnot.com tomtalks.uk @tomarbuthnot
    
	Disclaimer:   	
			THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
			OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
			While these scripts are tested and working in my environment, it is recommended 
			that you test these scripts in a test environment before using in your production 
			environment. Tom Arbuthnot further disclaims all implied warranties including, 
			without limitation, any implied warranties of merchantability or of fitness for 
			a particular purpose. The entire risk arising out of the use or performance of 
			this script and documentation remains with you. In no event shall Tom Arbuthnot, 
			its authors, or anyone else involved in the creation, production, or delivery of 
			this script/tool be liable for any damages whatsoever (including, without limitation, 
			damages for loss of business profits, business interruption, loss of business 
			information, or other pecuniary loss) arising out of the use of or inability to use 
			the sample scripts or documentation, even if Tom Arbuthnot has been advised of 
			the possibility of such damages.
    
     
	Acknowledgements: 	
    
	Assumptions:	      
			ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
    
	Limitations:		  										
    		
	Ideas/Wish list:
			Detects loops based on number of log lines, this can vary, should detect based on timestamp of line 
    
	Rights Required:	

	Known issues:	


.EXAMPLE
		Function-Template
 
		Description
		-----------
		Returns Objects

.EXAMPLE
		Function-Template -Param1
 
		Description
		-----------
		Actions Param1

# Parameters

.PARAMETER Param1
		Param1 description

.PARAMETER Param2
		Param2 Description
		

#>
  
  
  #############################################################
  # Param Block
  #############################################################
  
  # Sets that -Whatif and -Confirm should be allowed
  [cmdletbinding(SupportsShouldProcess=$true)]
  
  Param 	(
    [Parameter(Mandatory=$true,
    HelpMessage='Computer')]
    $Computer = 'localhost',
    
    
    [Parameter(Mandatory=$true,
    HelpMessage='UserName')]
    $UserName,
    
    [Parameter(Mandatory=$true,
    HelpMessage='Group')]
    $Group,
    
    [Parameter(Mandatory=$false,
    HelpMessage='Error Log location, default C:\<Command Name>_ErrorLog.txt')]
    [string]$ErrorLog = "c:\$($myinvocation.mycommand)_ErrorLog.txt",
    [switch]$LogErrors
    
  ) #Close Parameters
  
  
  #############################################################
  # Begin Block
  #############################################################
  
  Begin 	{
    Write-Verbose "Starting $($myinvocation.mycommand)"
    Write-Verbose "Error log will be $ErrorLog"
    
    # Script Level Variable to Stop Execution if there is an issue with any stage of the script
    $script:EverythingOK = $true
    
    
  } #Close Function Begin Block
  
  #############################################################
  # Process Block
  #############################################################
  
  Process {
    
    # First Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
        if(!$UserName -or !$group)
        
        {
          
          $(Throw 'A value for $user and $group is required.')
          
        }
        
        
        # Get the local group
        $group = [ADSI]("WinNT://$computer/$Group,group")
        
        # add the user    
        $group.add("WinNT://$username,user")
        
        
        
        
      } # Close Try Block
      
      Catch 	{ErrorCatch-Action} # Close Catch Block
      
      
    } # Close If EverthingOK Block 1
    
    #############################################################
    # Next Script Action or Try,Catch Block
    #############################################################
    
    # Second Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
        # Code Goes here
        
        
      } # Close Try Block
      
      Catch 	{ErrorCatch-Action} # Close Catch Block
      
      
    } # Close If EverthingOK Block 2
    
    
  } #Close Function Process Block 
  
  #############################################################
  # End Block
  #############################################################
  
  End 	{
    Write-Verbose "Ending $($myinvocation.mycommand)"
  } #Close Function End Block
  
} #End Function

Function PowerShell-RunAs {
<#
.SYNOPSIS  
		Synopsis goes here

.DESCRIPTION  
		Description goes here

.LINK  
    http://link.com
                
.NOTES  
	Version:
			0.8
  
	Author/Copyright:	 
			Copyright Tom Arbuthnot - All Rights Reserved
    
	Email/Blog/Twitter:	
			tom@tomarbuthnot.com tomtalks.uk @tomarbuthnot
    
	Disclaimer:   	
			THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
			OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
			While these scripts are tested and working in my environment, it is recommended 
			that you test these scripts in a test environment before using in your production 
			environment. Tom Arbuthnot further disclaims all implied warranties including, 
			without limitation, any implied warranties of merchantability or of fitness for 
			a particular purpose. The entire risk arising out of the use or performance of 
			this script and documentation remains with you. In no event shall Tom Arbuthnot, 
			its authors, or anyone else involved in the creation, production, or delivery of 
			this script/tool be liable for any damages whatsoever (including, without limitation, 
			damages for loss of business profits, business interruption, loss of business 
			information, or other pecuniary loss) arising out of the use of or inability to use 
			the sample scripts or documentation, even if Tom Arbuthnot has been advised of 
			the possibility of such damages.
    
     
	Acknowledgements: 	
    
	Assumptions:	      
			ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
    
	Limitations:		  										
    		
	Ideas/Wish list:
			Detects loops based on number of log lines, this can vary, should detect based on timestamp of line 
    
	Rights Required:	

	Known issues:	


.EXAMPLE
		Function-Template
 
		Description
		-----------
		Returns Objects

.EXAMPLE
		Function-Template -Param1
 
		Description
		-----------
		Actions Param1

# Parameters

.PARAMETER Param1
		Param1 description

.PARAMETER Param2
		Param2 Description
		

#>
  
  
  #############################################################
  # Param Block
  #############################################################
  
  # Sets that -Whatif and -Confirm should be allowed
  [cmdletbinding(SupportsShouldProcess=$true)]
  
  Param 	(
    [Parameter(Mandatory=$true,
    HelpMessage='Username')]
    $UserName,
    
    [Parameter(Mandatory=$true,
    HelpMessage='Password in Plain Text')]
    $Password,
    
    [Parameter(Mandatory=$true,
    HelpMessage='Program with full path to exe')]
    $Program,
    
    [Parameter(Mandatory=$false,
    HelpMessage='Error Log location, default C:\<Command Name>_ErrorLog.txt')]
    [string]$ErrorLog = "c:\$($myinvocation.mycommand)_ErrorLog.txt",
    [switch]$LogErrors
    
  ) #Close Parameters
  
  
  #############################################################
  # Begin Block
  #############################################################
  
  Begin 	{
    Write-Verbose "Starting $($myinvocation.mycommand)"
    Write-Verbose "Error log will be $ErrorLog"
    
    # Script Level Variable to Stop Execution if there is an issue with any stage of the script
    $script:EverythingOK = $true
    
    
  } #Close Function Begin Block
  
  #############################################################
  # Process Block
  #############################################################
  
  Process {
    
    # First Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
        $creds = New-Object System.Management.Automation.PSCredential -ArgumentList @($userName,(ConvertTo-SecureString -String $password -AsPlainText -Force))
        
        Start-Process "$Program" -Credential ($creds) -WorkingDirectory C:\
        
        
      } # Close Try Block
      
      Catch 	{ErrorCatch-Action} # Close Catch Block
      
      
    } # Close If EverthingOK Block 1
    
    #############################################################
    # Next Script Action or Try,Catch Block
    #############################################################
    
    # Second Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
        # Code Goes here
        
        
      } # Close Try Block
      
      Catch 	{ErrorCatch-Action} # Close Catch Block
      
      
    } # Close If EverthingOK Block 2
    
    
  } #Close Function Process Block 
  
  #############################################################
  # End Block
  #############################################################
  
  End 	{
    Write-Verbose "Ending $($myinvocation.mycommand)"
  } #Close Function End Block
  
} #End Function

Function New-RunASDesktopIEShortcut {
<#
.SYNOPSIS  
		Synopsis goes here

.DESCRIPTION  
		Description goes here

.LINK  
    http://link.com
                
.NOTES  
	Version:
			0.8
  
	Author/Copyright:	 
			Copyright Tom Arbuthnot - All Rights Reserved
    
	Email/Blog/Twitter:	
			tom@tomarbuthnot.com tomtalks.uk @tomarbuthnot
    
	Disclaimer:   	
			THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
			OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
			While these scripts are tested and working in my environment, it is recommended 
			that you test these scripts in a test environment before using in your production 
			environment. Tom Arbuthnot further disclaims all implied warranties including, 
			without limitation, any implied warranties of merchantability or of fitness for 
			a particular purpose. The entire risk arising out of the use or performance of 
			this script and documentation remains with you. In no event shall Tom Arbuthnot, 
			its authors, or anyone else involved in the creation, production, or delivery of 
			this script/tool be liable for any damages whatsoever (including, without limitation, 
			damages for loss of business profits, business interruption, loss of business 
			information, or other pecuniary loss) arising out of the use of or inability to use 
			the sample scripts or documentation, even if Tom Arbuthnot has been advised of 
			the possibility of such damages.
    
     
	Acknowledgements: 	
    
	Assumptions:	      
			ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
    
	Limitations:		  										
    		
	Ideas/Wish list:
			Detects loops based on number of log lines, this can vary, should detect based on timestamp of line 
    
	Rights Required:	

	Known issues:	


.EXAMPLE
		Function-Template
 
		Description
		-----------
		Returns Objects

.EXAMPLE
		Function-Template -Param1
 
		Description
		-----------
		Actions Param1

# Parameters

.PARAMETER Param1
		Param1 description

.PARAMETER Param2
		Param2 Description
		

#>
  
  
  #############################################################
  # Param Block
  #############################################################
  
  # Sets that -Whatif and -Confirm should be allowed
  [cmdletbinding(SupportsShouldProcess=$true)]
  
  Param 	(
    [Parameter(Mandatory=$true,
    HelpMessage='Param1 Description')]
    $ModulePath,
    
    [Parameter(Mandatory=$true,
    HelpMessage='Param1 Description')]
    $UserName,
    
    [Parameter(Mandatory=$true,
    HelpMessage='Param1 Description')]
    $Password,
    
    [Parameter(Mandatory=$false,
    HelpMessage='Param2 Description')]
    $Program = 'C:\Program Files\Internet Explorer\iexplore.exe',
    
    [Parameter(Mandatory=$true,
    HelpMessage='Param2 Description')]
    $ShortCutName,
    
    
    [Parameter(Mandatory=$false,
    HelpMessage='Error Log location, default C:\<Command Name>_ErrorLog.txt')]
    [string]$ErrorLog = "c:\$($myinvocation.mycommand)_ErrorLog.txt",
    [switch]$LogErrors
    
  ) #Close Parameters
  
  
  #############################################################
  # Begin Block
  #############################################################
  
  Begin 	{
    Write-Verbose "Starting $($myinvocation.mycommand)"
    Write-Verbose "Error log will be $ErrorLog"
    
    # Script Level Variable to Stop Execution if there is an issue with any stage of the script
    $script:EverythingOK = $true
    
    
  } #Close Function Begin Block
  
  #############################################################
  # Process Block
  #############################################################
  
  Process {
    
    # First Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
        # Build the Target String, has to be build in sections or the pareser gets confused
        $TargetPathSring = $null
        # if you want to see whats going on, add -noexit
        # [string]$TargetPathSring = ' -noexit -command "'
        [string]$TargetPathSring = ' -command "'
        $TargetPathSring = $TargetPathSring + "& Import-Module '$ModulePath'"
        $TargetPathSring = $TargetPathSring + " ; PowerShell-RunAs -UserName $UserName -Password $Password -Program '$Program' -Verbose"
        $TargetPathSring = $TargetPathSring + '"'
        
        $ws = New-Object -com WScript.Shell
        $Dt = $ws.SpecialFolders.Item("Desktop")
        $Scp = Join-Path -Path $Dt -ChildPath "$ShortCutName.lnk"
        $Sc = $ws.CreateShortcut($Scp)
        $Sc.IconLocation = $Program
        $Sc.TargetPath = "powershell.exe"
        $sc.Arguments = "$TargetPathSring"
        $Sc.Description = "$ShortCutName"
        $Sc.Save()
        
        
      } # Close Try Block
      
      Catch 	{ErrorCatch-Action} # Close Catch Block
      
      
    } # Close If EverthingOK Block 1
    
    #############################################################
    # Next Script Action or Try,Catch Block
    #############################################################
    
    # Second Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
        # Code Goes here
        
        
      } # Close Try Block
      
      Catch 	{ErrorCatch-Action} # Close Catch Block
      
      
    } # Close If EverthingOK Block 2
    
    
  } #Close Function Process Block 
  
  #############################################################
  # End Block
  #############################################################
  
  End 	{
    Write-Verbose "Ending $($myinvocation.mycommand)"
  } #Close Function End Block
  
} #End Function

Function New-LocalAccountAndDesktopShortcut {
<#
.SYNOPSIS  
		Synopsis goes here

.DESCRIPTION  
		Description goes here

.LINK  
    http://link.com
                
.NOTES  
	Version:
			0.8
  
	Author/Copyright:	 
			Copyright Tom Arbuthnot - All Rights Reserved
    
	Email/Blog/Twitter:	
			tom@tomarbuthnot.com tomtalks.uk @tomarbuthnot
    
	Disclaimer:   	
			THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
			OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
			While these scripts are tested and working in my environment, it is recommended 
			that you test these scripts in a test environment before using in your production 
			environment. Tom Arbuthnot further disclaims all implied warranties including, 
			without limitation, any implied warranties of merchantability or of fitness for 
			a particular purpose. The entire risk arising out of the use or performance of 
			this script and documentation remains with you. In no event shall Tom Arbuthnot, 
			its authors, or anyone else involved in the creation, production, or delivery of 
			this script/tool be liable for any damages whatsoever (including, without limitation, 
			damages for loss of business profits, business interruption, loss of business 
			information, or other pecuniary loss) arising out of the use of or inability to use 
			the sample scripts or documentation, even if Tom Arbuthnot has been advised of 
			the possibility of such damages.
    
     
	Acknowledgements: 	
    
	Assumptions:	      
			ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
    
	Limitations:		  										
    		
	Ideas/Wish list:
			Detects loops based on number of log lines, this can vary, should detect based on timestamp of line 
    
	Rights Required:	

	Known issues:	


.EXAMPLE
		Function-Template
 
		Description
		-----------
		Returns Objects

.EXAMPLE
		Function-Template -Param1
 
		Description
		-----------
		Actions Param1

# Parameters

.PARAMETER Param1
		Param1 description

.PARAMETER Param2
		Param2 Description
		

#>
  
  
  #############################################################
  # Param Block
  #############################################################
  
  # Sets that -Whatif and -Confirm should be allowed
  [cmdletbinding(SupportsShouldProcess=$true)]
  
  Param 	(
    [Parameter(Mandatory=$true,
    HelpMessage='Param1 Description')]
    $ModulePath,

    [Parameter(Mandatory=$true,
    HelpMessage='Param1 Description')]
    $Computer = 'localhost',
    
    [Parameter(Mandatory=$true,
    HelpMessage='Param1 Description')]
    $UserName,
    
    [Parameter(Mandatory=$true,
    HelpMessage='Param1 Description')]
    $Password,
    
    [Parameter(Mandatory=$false,
    HelpMessage='Param2 Description')]
    $Program = 'C:\Program Files\Internet Explorer\iexplore.exe',
    
    [Parameter(Mandatory=$true,
    HelpMessage='Param2 Description')]
    $ShortCutName,
    
    
    [Parameter(Mandatory=$false,
    HelpMessage='Error Log location, default C:\<Command Name>_ErrorLog.txt')]
    [string]$ErrorLog = "c:\$($myinvocation.mycommand)_ErrorLog.txt",
    [switch]$LogErrors
    
  ) #Close Parameters
  
  
  #############################################################
  # Begin Block
  #############################################################
  
  Begin 	{
    Write-Verbose "Starting $($myinvocation.mycommand)"
    Write-Verbose "Error log will be $ErrorLog"
    
    # Script Level Variable to Stop Execution if there is an issue with any stage of the script
    $script:EverythingOK = $true
    
    
  } #Close Function Begin Block
  
  #############################################################
  # Process Block
  #############################################################
  
  Process {
    
    # First Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
       
        New-LocalUserAccount -Computer $Computer -UserName $UserName -Password $Password -Verbose
        Start-Sleep -Seconds 1

        Add-LocalUserToLocalGroup -Computer $Computer -UserName $UserName -Group Administrators -Verbose
        Start-Sleep -Seconds 1

        New-RunASDesktopIEShortcut -ModulePath "$ModulePath" -UserName $UserName -Password $Password -Program "$Program" -ShortCutName "$ShortCutName" -Verbose
          
        
      } # Close Try Block
      
      Catch 	{ErrorCatch-Action} # Close Catch Block
      
      
    } # Close If EverthingOK Block 1
    
    #############################################################
    # Next Script Action or Try,Catch Block
    #############################################################
    
    # Second Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
       
        
        
      } # Close Try Block
      
      Catch 	{ErrorCatch-Action} # Close Catch Block
      
      
    } # Close If EverthingOK Block 2
    
    
  } #Close Function Process Block 
  
  #############################################################
  # End Block
  #############################################################
  
  End 	{
    Write-Verbose "Ending $($myinvocation.mycommand)"
  } #Close Function End Block
  
} #End Function

# Helper Functions below this line ##########################################

Function ErrorCatch-Action 
{
  Param 	(
    [Parameter(Mandatory=$false,
    HelpMessage='Switch to Allow Errors to be Caught without setting EverythingOK to False, stopping other aspects of the script running')]
    # By default any errors caught will set $EverythingOK to false causing other parts of the script to be skipped
    [switch]$SetEverythingOKVariabletoTrue
  ) # Close Parameters
  
  # Set EverythingOK to false to avoid running dependant actions
  If ($SetEverythingOKVariabletoTrue) {$script:EverythingOK = $true}
  else {$script:EverythingOK = $false}
  Write-Verbose "EverythingOK set to $script:EverythingOK"
  
  # Write Errors to Screen
  Write-Verbose 'Error Hit'
  Resolve-Error ($Global:Error)[0]
  # If Error Logging is runnning write to Error Log
  
  if ($LogErrors) {
    # Add Date to Error Log File
    Get-Date -format 'dd/MM/yyyy HH:mm' | Out-File $ErrorLog -Append
    $Error | Out-File $ErrorLog -Append
    '## LINE BREAK BETWEEN ERRORS ##' | Out-File $ErrorLog -Append
    Write-Warning "Errors Logged to $ErrorLog"
    # Clear Error Log Variable
    $Error.Clear()
  } #Close If
} # Close Error-CatchActons Function

Function Write-Log{
  
  [cmdletbinding(SupportsShouldProcess=$true)]
  
  param (
    [Parameter(Mandatory=$true,Position=1,HelpMessage='Log String')]
    [string]$LogString,
    
    [string]$LogLevel,
    [string]$LogFilePath
  )
  
  
  #-------------------------------------------------
  # Write-Log
  #-------------------------------------------------
  # Usage:	Writes logging information to file
  # **Parameters are not for interactive execution.**
  # Write-Log -LogString <String> -LogLevel <Error | INFO | Othery> -LogFilePath <LogFilePAth> 
  #
  #-------------------------------------------------
  
  $date = get-date
  
  If($Verbose){write-host $LogString -foregroundcolor green}
  
  $FinalLogString = '['+$date+ ']:'+'['+$LogLevel+']'+':'+$LogString
  
  Add-content $LogFilePath $FinalLogString
  
} # Close Write-Log Function


### helper functions

function New-Underline

{
  
<#

.Synopsis

 Creates an underline the length of the input string

.Example

 New-Underline -strIN "Hello world"

.Example

 New-Underline -strIn "Morgen welt" -char "-" -sColor "blue" -uColor "yellow"

.Example

 "this is a string" | New-Underline

.Notes

 NAME:

 AUTHOR: Ed Wilson

 LASTEDIT: 5/20/2009

 KEYWORDS:

.Link

 Http://www.ScriptingGuys.com

#>
  
  [CmdletBinding()]
  
  param(
    
    [Parameter(Mandatory = $true,Position = 0,valueFromPipeline=$true)]
    
    [string]
    
    $strIN,
    
    [string]
    
    $char = '=',
    
    [string]
    
    $sColor = 'Green',
    
    [string]
    
    $uColor = 'darkGreen',
    
    [switch]
    
    $pipe
    
  ) #end param
  
  $strLine= $char * $strIn.length
  
  if(-not $pipe)
  
  {
    
    Write-Host -ForegroundColor $sColor $strIN
    
    Write-Host -ForegroundColor $uColor $strLine
    
  }
  
  Else
  
  {
    
    $strIn
    
    $strLine
    
  }
  
} #end New-Underline function

 
 function Test-IsAdministrator

{
  
    <#

    .Synopsis

        Tests if the user is an administrator

    .Description

        Returns true if a user is an administrator, false if the user is not an administrator        

    .Example

        Test-IsAdministrator

    #>  
  
  param()
  
  $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
  
  (New-Object Security.Principal.WindowsPrincipal $currentUser).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
  
} #end function Test-IsAdministrator


Function Resolve-Error
{
  
  #############################################################################
  ##
  ## Resolve-Error
  ##
  ## From Windows PowerShell Cookbook (O'Reilly)
  ## by Lee Holmes (http://www.leeholmes.com/guide)
  ##
  ##############################################################################
  
<#

.SYNOPSIS

Displays detailed information about an error and its context.

#>
  
  param(
    ## The error to resolve
    $ErrorRecord = ($error[0])
  )
  
  Set-StrictMode -Off
  
  ''
  'If this is an error in a script you wrote, use the Set-PsBreakpoint cmdlet'
  'to diagnose it.'
  ''
  
  'Error details ($error[0] | Format-List * -Force)'
  '-'*80
  $errorRecord | Format-List * -Force
  
  'Information about the command that caused this error ' +
  '($error[0].InvocationInfo | Format-List *)'
  '-'*80
  $errorRecord.InvocationInfo | Format-List *
  
  'Information about the error''s target ' +
  '($error[0].TargetObject | Format-List *)'
  '-'*80
  $errorRecord.TargetObject | Format-List *
  
  'Exception details ($error[0].Exception | Format-List * -Force)'
  '-'*80
  
  $exception = $errorRecord.Exception
  
  for ($i = 0; $exception; $i++, ($exception = $exception.InnerException))
  {
    "$i" * 80
    $exception | Format-List * -Force
  }
  
} #end function Resolve-Error