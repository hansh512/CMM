################################################################################
# Code Written by Hans Halbmayr
# Created On: 28.05.2021
# Last change on: 29.06.2021
#
# Module: CMM
#
# Version 0.90
#
# Purpos: Helper function for module CMM
################################################################################

class getModuleData
{        
    getModuleData () {
        $this.PSObject.Properties.Add(
            (New-Object PSScriptProperty 'vaultName', {$this._vaultName})
        ) # end vaultName
        
        $this.PSObject.Properties.Add(
            (New-Object PSScriptProperty 'cfgNamePrefix', {$this._cfgNamePrefix})
        ) # end cfgNamePrefix
        
        $this.PSObject.Properties.Add(
            (New-Object PSScriptProperty 'modulePrfx', {$this._modulePrefix})
        ) # end modulePrefix
        
        $this.PSObject.Properties.Add(
            (New-Object PSScriptProperty 'verStr', {$this._verStr})
        ) # end verStr  
        $this.PSObject.Properties.Add(
            (New-Object PSScriptProperty 'entrySep', {$this._entrySep})
        ) # end verStr        
    } # end getModuleData

    [hashtable]GetConfig ([string]$ModuleName,
                      [version]$Version,
                      [string]$Name,
                      [bool]$SkipTemplateVersionFiltering
                     )
    {        
        return returnConfigData -ModuleName $ModuleName -Version $Version -Name $Name -SkipStrictVersionFiltering:$SkipTemplateVersionFiltering;
           
    } # end method getConfig

    [hashtable]GetConfig ([string]$ModuleName,
                      [version]$Version,
                      [string]$Name
                     )
    {        
        return returnConfigData -ModuleName $ModuleName -Version $Version -Name $Name;         
    } # end method getConfig
    [array]GetModuleList ()
    {        
        $errMsg='Failed to load module list.';
        $ModuleList='';      
        try {
            $tmp=(Get-SecretInfo -Vault $Script:vaultName -Name ($Script:ModuleCfgName) -ErrorAction Stop).Metadata;
            $ModuleList=($tmp.RegisteredModuleList).Split($Script:sepVal);
        } # end try
        catch {
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        }; # end catch
        return [array]([System.Linq.Enumerable]::Distinct([string[]]$ModuleList));        
    } # end method getModuleList

    [array]GetCfgOrTemplateList ([string]$ModuleName,
                                 [Boolean]$IsTemplate)
    {        
        $ModuleName=$ModuleName.ToUpper();
        if ($IsTemplate -eq $true)
        {
            $filterStr=$Script:TemplatePrefix;
        } #end if
        else {
            $filterStr=$Script:ConfigPrefix;            
        }; # end else
        #$replaceString=$filterStr;
        if (!([System.String]::IsNullOrEmpty($ModuleName)))
        {
             $filterStr+=($ModuleName+$Script:EntrySep);
        }; # end if
        $tmpList=@(Get-SecretInfo -Name ($filterStr +'*') -Vault $Script:vaultName);     
        if ($tmpList.Count -eq 0)
        {
            $templateList=@('no config available');
        } # end if
        else {            
            $templateList=[System.Collections.ArrayList]::new();
            foreach ($entry in $tmpList.Name)
            {
                [void]$templateList.Add($entry.Replace($filterStr,''));
            }; # end if
        }; # end else  
                  
        return [array]($templateList);        
    } # end method getTemplateList

    [array]GetModuleAndVerList ()
    {        
        $tmpList=@(Get-SecretInfo -Name ($Script:RegisteredModulePrefix + 'Template_*') -Vault $Script:vaultName);     
        if ($tmpList.Count -eq 0)
        {
            $moduleList=@('no config available');
        } # end if
        else {            
            [array]$moduleList=@();
            foreach ($entry in $tmpList.Name)
            {
                $tmpArr=$entry.Replace($Script:RegisteredModulePrefix,'').split($Script:EntrySep);
                $moduleList+=($tmpArr[1] + ($Script:EntrySep+'Ver:') + ($tmpArr[2]));
            }; # end if
        }; # end else            
        return [array]$moduleList;        
    } # end method getModuleList

    [array]GetConfigForTemplate ([string]$Template
                          )
    { 
        $Template=$Template.ToUpper();
        try {
            $tmpList=@(Get-SecretInfo -Name ($Script:ConfigPrefix+$Template+'*') -Vault $Script:vaultName); 
            if ($tmpList.Count -gt 0)
            {
                $rvList=[System.Collections.ArrayList]::new();
                for ($i=0;$i -lt $tmpList.Count;$i++)
                {
                    [void]$rvList.Add(($tmpList[$i].name).Replace(($Script:ConfigPrefix+$Template+$Script:EntrySep),''));
                }; # end for
            } # end if
            else {
                $rvList=@()
            }; # end else
        } # end try
        catch {
            $rvList=@();
            $errMsg=('Failed to list configurations for Template ' + $Template);
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        }; # end catch
        return [array]$rvList;
    } # end method GetConfigForTemplate


    [array]GetConfigerationList ([string]$ModuleName,
                           [version]$Version
                          )
    {        
        return [array](returnConfigList -ModuleName $ModuleName -Version $Version);
    } # end method GetConfigerationList

    [array]GetConfigerationList ([string]$ModuleName,
                           [version]$Version,
                           [bool]$SkipTemplateVersionFiltering
                          )
    {        
        return [array](returnConfigList -ModuleName $ModuleName -Version $Version -SkipStrictVersionFiltering:$SkipTemplateVersionFiltering);
        
    } # end method GetConfigerationList
    
    [string]GetDefaultConfig ([string]$ModuleName,
                              [version]$Version
                             )
    {
        $errMsg=('Failed to load template version ' + $Version.ToString() + ' for module ' + $ModuleName);
        try {
            $templateName=getClosestVersion -SearchString ($script:TemplatePrefix + $ModuleName + '_*') -Version $Version;
            if ([System.String]::IsNullOrEmpty($templateName))
            {
                $rv='';
            } # end if
            else {
                $mData=loadMetadata -MetaData ((Get-SecretInfo -Vault $script:vaultName -Name $templateName -ErrorAction Stop).Metadata);
                $rv=$mData.($script:MetadataDefaultCfg);
            }; # end else                        
        } # end try
        catch {
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
            $rv='';
        }; # end catch        
        return [string]$rv;
    } # end method GetDefaultConfig
    # list of constants needed in Argument Completer and default values
    hidden [string]$_vaultName=$Script:vaultName;
    hidden [string]$_cfgNamePrefix=$Script:cfgNamePrefix;
    hidden [string]$_modulePrefix=$Script:RegisteredModulePrefix;  
    hidden [string]$_verStr=$Script:verStr;
    hidden [string]$_entrySep=$Script:EntrySep;
    
}; # end class getModuleData

function getClosestVersion
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][string]$SearchString,
      [Parameter(Mandatory = $true, Position = 1)][Version]$Version
     )

    try {
        $tmpList=@(Get-SecretInfo -Vault $script:vaultName -Name $SearchString -ErrorAction Stop);
        if ($tmpList.Count -eq 0)
        {
            $cfName=@('');
            [version]$lastVer='0.0.0.0';
        } # end if
        else {            
            [version]$lastVer='0.0.0.0';
            $cfName='';
            foreach ($entry in $tmpList.Name)
            {
                [version]$mVer=($entry.split($Script:EntrySep))[3];
                if (($mVer -ge $lastVer) -and ($mVer -le $Version))
                {
                    [version]$lastVer=$mVer;
                    $cfName=$entry;
                }; # end if
            }; # end if
        }; # end else 
    } # end try
    catch {
        $errMsg='Failed to read the template/config info';
        writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
    }; # end catch
    return $cfName
}; # end function getClosestVersion
function returnConfigList
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][string]$ModuleName,
      [Parameter(Mandatory = $true, Position = 1)][Version]$Version,
      [Parameter(Mandatory = $false, Position = 3)][switch]$SkipStrictVersionFiltering=$false
     )
     
    If (!([System.String]::IsNullOrEmpty($ModuleName)))
    {
        #$ModuleName+='_';
        $ModuleName+=$Script:EntrySep;
    }; #end if
    
    if ($SkipStrictVersionFiltering)
    {
        $cfgPrefixStr=($Script:ConfigPrefix + $ModuleName);
         
    } # end if
    else {        
        $cfgSearchStr=@(getClosestVersion -SearchString ($Script:TemplatePrefix + $ModuleName+'*') -Version $Version);
        $cfgPrefixStr=($cfgSearchStr.Replace($Script:TemplatePrefix,($Script:ConfigPrefix)))+$Script:EntrySep;
    }; # end else
   
    $tmpList=@(Get-SecretInfo -Name ($cfgPrefixStr + '*') -Vault $Script:vaultName);

    
    $rvList=[System.Collections.ArrayList]::New(); # init array    
    if ($tmpList.Count -eq 0)
    {
        $cfgList=@('no data available');
    } # end if
    else {                        
        $cfgList=[System.Collections.ArrayList]::new(); # init array
        if ($SkipStrictVersionFiltering)
        {
            for ($i=0;$i -lt $tmpList.Count; $i++)
            {
                $tmp=$tmpList[$i].name.split($Script:EntrySep);                
                if ($Version -ge [version]$tmp[3]) # check if version is ok
                {
                    [void]$cfgList.Add($tmp[4]); # end to list
                }; # end if
            }; # end for
        } # end if
        else {
            for ($i=0;$i -lt $tmpList.Count; $i++)
            {
                $tmp=$tmpList[$i].name.split($Script:EntrySep);  
                [void]$cfgList.Add($tmp[4]); # end to list              
                
            }; # end for
        }; # end else
        
        $cfgList.sort(); # sort list
        $lastVal=''; # init helper val            
        for ($i=0;$i -lt $cfgList.count;$i++)
        {                
            if ($cfgList[$i] -ne $lastVal) # check if duplicate (case ingnore case!!!)
            {
                [void]$rvList.Add($cfgList[$i]); # if not add to list
            }; # end if
            $lastVal=$cfgList[$i];
        }; # end for            
    }; # end else
    return [array]$rvList; 
}; # end function returnConfigList
function returnConfigData
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][string]$ModuleName,
      [Parameter(Mandatory = $true, Position = 1)][Version]$Version,
      [Parameter(Mandatory = $true, Position = 2)][string]$Name,
      [Parameter(Mandatory = $false, Position = 3)][switch]$SkipStrictVersionFiltering=$false
     )
     
    $ModuleName=$ModuleName.ToUpper();
    writeLogOutput -LogString 'Searching configuration data';
    $errMsg=('Failed to read configuration data ' + ($Script:RegisteredModulePrefix + $ModuleName + '_*_' + $Name));
    
    if ($SkipStrictVersionFiltering)
    {
        $tmpList=@(Get-SecretInfo -Name ($Script:ConfigPrefix + $ModuleName + '_*_' + $Name) -Vault $Script:vaultName);                
    } # end if
    else {
        
        $tmpList=@(Get-SecretInfo -Name ($Script:TemplatePrefix + $ModuleName + '_*') -Vault $Script:vaultName); 
    }; # end else
            
    if ($tmpList.Count -eq 0)
    {
        $cfName=@('no config available');
        [version]$lastVer='0.0.0.0';
    } # end if
    else {            
        [version]$lastVer='0.0.0.0';
        $cfName='';
        foreach ($entry in $tmpList.Name)
        {
            [version]$mVer=($entry.split($Script:EntrySep))[3];
            if (($mVer -ge $lastVer) -and ($mVer -le $Version))
            {
                [version]$lastVer=$mVer;
                $cfName=$entry;
            }; # end if
        }; # end if
    }; # end else 
    if ([System.String]::IsNullOrEmpty($cfName))
    {
        writeLogOutput -LogString ('No data available') -LogType Warning;
        return $null;
    }; # end if
    $errMsg=('Failed to read the configuraton ' + $cfName); 
    try {                        
        writeLogOutput -LogString ('Reading credential data from ' +  $cfName);
        $configName=($script:ConfigPrefix+$ModuleName+$Script:EntrySep+$lastVer.tostring()+$Script:EntrySep + $Name) ;
        $cfgTmp=(Get-SecretInfo -Vault Local -Name $configName -ErrorAction Stop);
        if ($null -eq $cfgTmp)
        {
            return $null;
        }; # end if
        $tmpCfgMetaData=loadMetadata -MetaData ($cfgTmp.Metadata);
        $templateName=$tmpCfgMetaData.$Script:BaseTemplate
        writeLogOutput -LogString ('Reading configuration from ' +  $cfName);
        $templateMetadata=loadMetadata -MetaData ((Get-SecretInfo -Vault $Script:vaultName -Name $templateName -ErrorAction Stop).Metadata);
        $credVarName=($tmpCfgMetaData.($script:CredVarName));
        if ($templateMetadata.Count -gt 0)
        {
            $errMsg='Failed to read credential information'
            $cred=(Get-Secret -Vault $Script:vaultName -Name $configName -ErrorAction Stop);
            $errMsg='Failed to read the template information'
            $rv=@{
                ConfigName=$Name;
                ConfigVersion=$lastVer;
                Data=@{
                    $credVarName=$cred;
                }; # end data
                InconsistentAttributes=@();
            }; # end rv
            $attribList=($templateMetadata.($script:MetadataPList)).split($Script:pLstSepVal); # get list of attributes
            $attribList=([System.Linq.Enumerable]::Except([string[]]$attribList,[string[]]$credVarName)); # exclude attribute for cred 
            
            [Func[string,bool]] $p={ param($e); return (-not $e.StartsWith($script:cfgNamePrefix))}; # prepare linq query
            $rParamList=[System.Linq.Enumerable]::Where(([string[]]($templateMetadata.Keys)), $p);
            foreach ($attrib in $attribList) # iterate through attrib list
            {
                try {
                    [Func[string,bool]] $p={ param($e); return ($e.EndsWith(($script:paramPrefixSep)+$attrib))}; # prepare linq query
                    $rParam=[System.Linq.Enumerable]::Where(([string[]]($rParamList)), $p); # get data type info
                    $rv.data.add($attrib,($tmpCfgMetaData.$attrib -as [System.Type]$($templateMetadata.$($rParam).Split($Script:sepVal)[0]))); # add entry with the appropriate data type
                } # end try
                catch {
                    writeLogOutput ('Failed to read the attribute ' + $attrib + '. Template and config are in an inconsistant state.') -LogType Warning;
                    $rv.InconsistentAttributes+=$attrib;
                }; # end catch
                
            }; # end forach
        } # end if
        else {
            #$rv=@{};
            writeLogOutput -LogString ('Template ' + $templateName + ' not found. Cannot verify data structure integrity.') -LogType Warning;
            return $null;
        }; # end else
    } # end try
    catch {
        writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        #$rv=@{};
        return $null;
    }; # end catch
    return $rv;     
}; # end function returnConfigData


function initEnumerator
{

try {
Add-Type -TypeDefinition @'
public enum OutFormat {
    Table,
    List,
    PassValue
}
'@
}
catch {
    # type exist
}; # end catch

}; # end initEnumerator

function createNewTable
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][string]$TableName,
      [Parameter(Mandatory = $true, Position = 1)][array]$FieldList
     )
    
    $tmpTable = New-Object System.Data.DataTable $TableName; # init table   
    $fc=$FieldList.Count;
        
    for ($i=0;$i -lt $fc;$i++)
    {
        if ((!($null -eq $FieldList[$i][1])) -and ($FieldList[$i][1].GetType().name -eq 'runtimetype'))
        {
            [void]($tmpTable.Columns.Add(( New-Object System.Data.DataColumn($FieldList[$i][0],$FieldList[$i][1])))); # add columns to table
        } # end if
        else
        {
            [void]($tmpTable.Columns.Add(( New-Object System.Data.DataColumn($FieldList[$i][0],[System.String])))); # add columns to table
        }; #end else
    }; #end for
    
    return ,$tmpTable;
}; # end createNewTable

<#function testPort
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)][string]$ComputerName,
      [Parameter(Mandatory = $true, Position = 1)][int]$Port, 
      [Parameter(Mandatory = $false, Position = 2)][int]$TcpTimeout=100    
     )

    begin {        
    }; #end begin
    
    process {
        writeTolog -LogString ('Testing port ' + $Port.ToString() + ' on computer ' + $computerName);
        $TcpClient = New-Object System.Net.Sockets.TcpClient
        $Connect = $TcpClient.BeginConnect($ComputerName, $Port, $null, $null)
        $Wait = $Connect.AsyncWaitHandle.WaitOne($TcpTimeout, $false)
        if (!$Wait) 
        {
	        writeTolog -LogString ('Server ' + $computerName + ' failed to answer on port ' + $Port.ToString()) -LogType Warning;
            return $false;
        } # end if
        else 
        {	        
	        return $true;
        }; # end else        
    } # end process

    end {        
        $TcpClient.Close();
        $TcpClient.Dispose();
    } # end END

}; # end function testPort
#>
function verifyModuleName
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][string]$ModuleName,
      [Parameter(Mandatory = $true, Position = 1)][string]$TextString
     )    

    $checkVal=0;
    $controlVal=0;
    $tmpVal=1;
    $notAllowedChars=@(
        ' ',
        $Script:EntrySep
    ); #end notAllowedChars

    for ($i=0;$i -lt $notAllowedChars.Count;$i++)
    {        
        if ($ModuleName.Contains($notAllowedChars[$i]))
        {
            Write-Warning ('The character ''' + $notAllowedChars[$i] + ''' is not allowed in the ' + $TextString);
        } # end if
        else {
            
            $checkVal = ($checkVal -bor ($tmpVal)); 
        }; # end if
        
        $controlVal=($controlVal -bor ($tmpVal));
        $tmpVal = ($tmpVal -shl 1);  
    }; # end for
    return ($checkVal -eq $controlVal);
}; #end funciton verifyModuleName


function writeLogOutput
{
[CmdLetBinding()]
param([Parameter(Mandatory = $true, Position = 0)]
      [string]$LogString,
      [Parameter(Mandatory = $false, Position = 2)]
      [ValidateSet('Info','Warning','Error')] 
      [string]$LogType="Info",
      
      [Parameter(Mandatory = $false, Position = 10)]
      [switch]$ShowInfo=$false      
     )

    if (($LogType -eq 'Info') -and ($ShowInfo.IsPresent -eq $false))
    {
        Write-Verbose -Message $LogString;
    }; # end if

    switch ($LogType)
    {
        {$_ -eq 'Info' -and $ShowInfo}  {
            Write-Host $LogString;
            break;
        }; # end info and ShowInfo
        'Warning'                       {
            Write-Warning $LogString;
            break;
        }; # end warning
        'Error'                         {
            Write-Host $LogString -ForegroundColor Red -BackgroundColor Black;
            break;
        }; # end Error
    }; # end switch
}; # end function writeLogOutput

function writeLogError
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][string]$ErrorMessage,
      [Parameter(Mandatory = $true, Position = 1)][string]$PSErrMessage,
      [Parameter(Mandatory = $true, Position = 2)][string]$PSErrStack
     )

    writeLogOutput -LogString $ErrorMessage -LogType Error;
    writeLogOutput -LogString $PSErrMessage -LogType Error;
}; # end function writeLogError

