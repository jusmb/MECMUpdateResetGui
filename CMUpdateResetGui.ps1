<#

Copyright (c) Microsoft Corporation.
By Justin T. Mnatsakanyan-Barbalace, Sr. CSA-E, Microsoft
MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

Add-Type -AssemblyName PresentationCore, PresentationFramework

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="800" Height="400" Background="#00ced1">
    <Grid>
        <Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="Site Code" Margin="16.5,13,0,0"/>
        <TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="23" Width="120" TextWrapping="Wrap" Margin="85,12,0,0" Name="SiteCodeTxt"/>
        <Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="SMS Provider Server" Margin="240.5,13,0,0"/>
        <TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="30" Width="282" TextWrapping="Wrap" Margin="374,12,0,0" Name="SMSProviderServerName"/>

        <DataGrid HorizontalAlignment="Left" IsReadOnly="True" VerticalAlignment="Top" Background="#ffffff" Width="712" Height="122" Margin="44,137,0,0" Name="UpdateListData">
            <DataGrid.Columns>
            </DataGrid.Columns>
        </DataGrid>
        <Button Content="Refresh Updates" HorizontalAlignment="Left" VerticalAlignment="Top" Width="120" Margin="76,295,0,0" Name="RefreshBtn" Height="19"/>
        <Button Content="Reset Selected" HorizontalAlignment="Left" VerticalAlignment="Top" Width="120" Margin="266,295,0,0" Name="Remove"/>
        <Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="SQL Server" Margin="10,54,0,0"/>
        <TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="28" Width="210" TextWrapping="Wrap" Margin="82,53,0,0" Name="SqlServerTxt"/>
        <Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="Sql Instance" Margin="295,54,0,0"/>
        <TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="23" Width="120" TextWrapping="Wrap" Margin="370,53,0,0" Name="SqlInstanceTxt"/>
        <Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="SQL Database" Margin="502,54,0,0"/>
        <TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="23" Width="120" TextWrapping="Wrap" Margin="600,53,0,0" Name="SQLDBTxt"/>
    </Grid>
</Window>
"@


$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }

##################
# Events
####################
#$LastUsedMP=(Get-ItemProperty HKLM:\SOFTWARE\Microsoft\CCM\LocationServices ).EventLastUsedMP
$Provider=$null
$Instance=$null
$SiteServer=(Get-ItemProperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\ConfigMgr10\AdminUI\Connection).server
$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $SiteServer)
$RegKey= $Reg.OpenSubKey("SOFTWARE\\Microsoft\\SMS\\Setup")
$Provider = $RegKey.GetValue("Provider Location")
if($Provider -ne $null){
    $SMSProviderServerName.Text=$Provider
}
$RegKey= $Reg.OpenSubKey("SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account")
$SQLServer=$RegKey.GetValue("Server")
$SqlServerTxt.Text=$SQLServer
$SQLDB=$RegKey.GetValue("Database Name")
$SQLDBTxt.Text=$SQLDB
$Instance=$RegKey.GetValue("Instanace")
$SqlInstanceTxt.Text=$instance


$RegKey= $Reg.OpenSubKey("SOFTWARE\\Microsoft\\SMS\\DP")
$SiteCodeTxt.Text=$RegKey.GetValue("SiteCode")

$StateCodes=(262145,262146,327679,65537,65538,131071,131073,131074,131075,196607)
$RefreshBtn.Add_Click({
    Fresh
})

$Remove.Add_Click({
    if($SqlInstanceTxt.Text -ne $null -and $SqlInstanceTxt.Text -ne ""){
        $instance = "-i $($SqlInstanceTxt.Text) "
    }
    $SiteServer=(get-CMSite).ServerName
    $Roles=Get-CMSiteRole -SiteCode $($SiteCodeTxt.Text)
    #$DBServer=($Roles | Where-Object {$_.RoleName -eq "SMS SQL Server"} | Where-Object {$_.RoleName -eq "SMS SQL Server"} ).NetworkOSPath.replace("\","")
    foreach($SelectedUpdate in $UpdateListData.SelectedItems){
        $PackageGuid=$SelectedUpdate.PackageGuid
        Start-Process "\\$($SiteServer)\SMS_$($SiteCodeTxt.Text)\cd.latest\SMSSETUP\TOOLS\CMUpdateReset\CMUpdateReset.exe" -ArgumentList "-S $($SqlServerTxt.Text) -D $($SQLDBTxt.Text) $instance -P $PackageGuid"
    }
    Fresh
})

Function Fresh{
    Write-Host $SiteCodeTxt.Text
    # Site configuration
    $SiteCode = $SiteCodeTxt.Text # Site code 
    #$ProviderMachineName = $SMSProviderServerName.Text # SMS Provider machine name

    # Customizations
    $initParams = @{}
    #$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
    #$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

    # Do not change anything below this line

    # Import the ConfigurationManager.psd1 module 
    if($Provider -eq $null -and $SMSProviderServerName.Text -ne $null){
        $Provider=$SMSProviderServerName.Text
        if($Provider -like "*,*"){$Provider=$Provider.Split(",")}
    }

    if((Get-Module ConfigurationManager) -eq $null) {
        if(Test-Path "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"){
            Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
        }else{ 
            [System.Windows.MessageBox]::Show('Iam unable to locate the "Configuration Manager" PowerShell Module. The MECM console may not be installed but is required.')
            $Console=$false
        }
    }        



    # Connect to the site's drive if it is not already present
    if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        if($Provider.count -gt 1){
        try{
           New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $Provider[0] @initParams
        }catch{
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $Provider[1] @initParams
        }
        }else{
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $Provider @initParams
        }
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams
    Write-Host $($SiteCode)
    $CMUpdates=Get-CMSiteUpdate -Fast  | Select Name, PackageGuid, State, MoreInfoLink | Where-Object {$_.State -in $StateCodes}
    $UpdateSet=New-Object System.Collections.ArrayList
    Foreach($Update in $CMUpdates){
         $Row = "" | Select Name, PackageGuid, State, MoreInfoLink
         $Row.name=$Update.Name
         $Row.PackageGuid=$Update.PackageGuid
         switch ($Update.State){
            262145 {$Row.State="DOWNLOAD_IN_PROGRESS"}
            262146 {$Row.State="DOWNLOAD_SUCCESS"}
            327679 {$Row.State="CONTENT_REPLICATING"}
            65538 {$Row.State="CONTENT_REPLICATION_SUCCESS"}
            131071 {$Row.State="CONTENT_REPLICATION_FAILED"}
            131073 {$Row.State="PREREQ_IN_PROGRESS"}
            131074 {$Row.State="PREREQ_WARNING"}
            196607 {$Row.State="PREREQ_ERROR"}
        }
         
         $Row.MoreInfoLink=$Update.MoreInfoLink
         $UpdateSet.add($Row)
    } 
    $UpdateListData.ItemsSource=$UpdateSet
}

Fresh
###########################



$Global:SyncHash = [HashTable]::Synchronized(@{})
$SyncHash.Window = $Window
$Jobs = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
$initialSessionState = [initialsessionstate]::CreateDefault()

Function Start-RunspaceTask
{
    [CmdletBinding()]
    Param([Parameter(Mandatory=$True,Position=0)][ScriptBlock]$ScriptBlock,
          [Parameter(Mandatory=$True,Position=1)][PSObject[]]$ProxyVars)
            
    $Runspace = [RunspaceFactory]::CreateRunspace($InitialSessionState)
    $Runspace.ApartmentState = 'STA'
    $Runspace.ThreadOptions  = 'ReuseThread'
    $Runspace.Open()
    ForEach($Var in $ProxyVars){$Runspace.SessionStateProxy.SetVariable($Var.Name, $Var.Variable)}
    $Thread = [PowerShell]::Create('NewRunspace')
    $Thread.AddScript($ScriptBlock) | Out-Null
    $Thread.Runspace = $Runspace
    [Void]$Jobs.Add([PSObject]@{ PowerShell = $Thread ; Runspace = $Thread.BeginInvoke() })
}

$JobCleanupScript = {
    Do
    {    
        ForEach($Job in $Jobs)
        {            
            If($Job.Runspace.IsCompleted)
            {
                [Void]$Job.Powershell.EndInvoke($Job.Runspace)
                $Job.PowerShell.Runspace.Close()
                $Job.PowerShell.Runspace.Dispose()
                $Job.Powershell.Dispose()
                
                $Jobs.Remove($Job)
            }
        }

        Start-Sleep -Seconds 1
    }
    While ($SyncHash.CleanupJobs)
}

Get-ChildItem Function: | Where-Object {$_.name -notlike "*:*"} |  select name -ExpandProperty name |
ForEach-Object {       
    $Definition = Get-Content "function:$_" -ErrorAction Stop
    $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "$_", $Definition
    $InitialSessionState.Commands.Add($SessionStateFunction)
}


$Window.Add_Closed({
    Write-Verbose 'Halt runspace cleanup job processing'
    $SyncHash.CleanupJobs = $False
})

$SyncHash.CleanupJobs = $True
function Async($scriptBlock){ Start-RunspaceTask $scriptBlock @([PSObject]@{ Name='DataContext' ; Variable=$DataContext},[PSObject]@{Name="State"; Variable=$State},[PSObject]@{Name = "SyncHash";Variable = $SyncHash})}

Start-RunspaceTask $JobCleanupScript @([PSObject]@{ Name='Jobs' ; Variable=$Jobs })



$Window.ShowDialog()


<#
State Codes we care about:
DOWNLOAD_IN_PROGRESS = 262145
DOWNLOAD_SUCCESS = 262146
DOWNLOAD_FAILED = 327679
CONTENT_REPLICATING = 65537
CONTENT_REPLICATION_SUCCESS = 65538
CONTENT_REPLICATION_FAILED = 131071
PREREQ_IN_PROGRESS = 131073
PREREQ_SUCCESS = 131074
PREREQ_WARNING = 131075
PREREQ_ERROR = 196607

Full List of Update State Codes
DOWNLOAD_IN_PROGRESS = 262145
DOWNLOAD_SUCCESS = 262146
DOWNLOAD_FAILED = 327679
APPLICABILITY_CHECKING = 327681
APPLICABILITY_SUCCESS = 327682
APPLICABILITY_HIDE = 393213
APPLICABILITY_NA = 393214
APPLICABILITY_FAILED = 393215
CONTENT_REPLICATING = 65537
CONTENT_REPLICATION_SUCCESS = 65538
CONTENT_REPLICATION_FAILED = 131071
PREREQ_IN_PROGRESS = 131073
PREREQ_SUCCESS = 131074
PREREQ_WARNING = 131075
PREREQ_ERROR = 196607
INSTALL_IN_PROGRESS = 196609
INSTALL_WAITING_SERVICE_WINDOW = 196610
INSTALL_WAITING_PARENT = 196611
INSTALL_SUCCESS = 196612
INSTALL_PENDING_REBOOT = 196613
INSTALL_FAILED = 262143
INSTALL_CMU_VALIDATING = 196614
INSTALL_CMU_STOPPED = 196615
INSTALL_CMU_INSTALLFILES = 196616
INSTALL_CMU_STARTED = 196617
INSTALL_CMU_SUCCESS = 196618
INSTALL_WAITING_CMU = 196619
INSTALL_CMU_FAILED = 262142
INSTALL_INSTALLFILES = 196620
INSTALL_UPGRADESITECTRLIMAGE = 196621
INSTALL_CONFIGURESERVICEBROKER = 196622
INSTALL_INSTALLSYSTEM = 196623
INSTALL_CONSOLE = 196624
INSTALL_INSTALLBASESERVICES = 196625
INSTALL_UPDATE_SITES = 196626
INSTALL_SSB_ACTIVATION_ON = 196627
INSTALL_UPGRADEDATABASE = 196628
INSTALL_UPDATEADMINCONSOLE = 196629
#>