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


function initModule
{
    
    try
    {
        $vaultList=@(Get-SecretVault -ErrorAction Stop | Where-Object {$_.ModuleName -eq $script:SecureVaultModuleName} );
    } # end try
    catch
    {
        $errMsg='Failed to search for a vault';
        writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        return;
    }; # end catch

    if (!($vaultList.name -contains $Script:vaultName))  # verify if vault exist
    {
        $msg=('Creating vault ' + $Script:vaultName);
        writeLogOutput -LogString $msg;
        try {
            Register-SecretVault -ModuleName $script:SecureVaultModuleName -Name $Script:vaultName -Description 'Created by CMM PS module';
        } # end try
        catch {
            $errMsg=('Failed to register vault ' + $Script:vaultName);
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        }; # end catch
    }; # end if
    
    verifyModuleConfig;
    
}; # end function initModule



function verifyModuleConfig
{
    # verify if module config exist, if not create default config
    if (!($script:ModuleCfg=Get-SecretInfo -Name $Script:ModuleCfgName -Vault $Script:vaultName -ErrorAction SilentlyContinue)) # check if the module config exist
    {
        try { # if module cfg does not exist, create an empty cfg
            $mCfgMetaData= @{
                CfgVersion=$script:ModuleCfgVersion;
                RegisteredTemplateList='';
                RegisteredModuleList='';
                $script:CreateDate=[System.DateTime]::Now;
                $script:ChangeDate=[System.DateTime]::Now;
            }; # end mCfgMetaData
            Set-Secret -Name $Script:ModuleCfgName -Secret ([System.Guid]::NewGuid().guid) -Metadata $mCfgMetaData -Vault $Script:vaultName;
            $Script:ModuleCfg=Get-SecretInfo -Name $Script:ModuleCfgName -Vault $Script:vaultName -ErrorAction Stop;
        } # end try
        catch {
            $errMsg='Failed to initialize the module configuration.';
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        }; # end catch
    };# end if
    
}; # end function getModuleConfig


function addBackLinkToTemplate
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)][string]$ConfigurationName,
      [Parameter(Mandatory = $true, Position = 1)][string]$TemplateName
     )

    try {
        $errMsg=('Failed to read the template ' + $TemplateName);
        writeLogOutput -LogString ('Reading template ' + $TemplateName);
        $templateData=loadMetadata -MetaData((Get-SecretInfo -Vault $Script:vaultName -Name $TemplateName -ErrorAction Stop).Metadata);
        if ([System.String]::IsNullOrEmpty($templateData.$Script:ConfigurationBackLink))
        {
            $backLink=$ConfigurationName;
        } # end if
        else {
           $backLink=(($templateData.$Script:ConfigurationBackLink).Split($Script:blSepVal));
            $backLink+=$ConfigurationName; 
        }; # end else
        $backLink=[System.Linq.Enumerable]::Distinct([string[]]$backLink);
        $templateData.($Script:ConfigurationBackLink)=($backLink -join $Script:blSepVal);
        $errMsg=('Failed to save backlink info for config ' + $ConfigurationName + ' to template ' + $TemplateName);
        writeLogOutput -LogString ('Writing link to configuration ' + $ConfigurationName + ' to template ' + $TemplateName);
        Set-SecretInfo -Vault $Script:vaultName -Name $TemplateName -Metadata $templateData -ErrorAction Stop;
    } # end try
    catch {
        writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
    }; # end catch
}; # end function addBackLinkToTemplate

function loadMetadata
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)]$MetaData
     )
     $tmpData=@{};
     foreach ($key in $MetaData.keys)
     {
        if (($MetaData.$key.Gettype()).Name -ne 'Int64') # make sure it's not Int64 (seams to occur only in PS Core)
        {
            $tmpData.Add($key,$MetaData.$key -as $MetaData.$key.Gettype());
        } # end if
        else {            
            if ($Script:ShowIntConversionWarning)
            {
                writeLogOutput -LogString 'Datatype Int64 dedected, converting to Int32' -LogType Warning;
            }; # end if
            try {
                $tmpData.Add($key,[int32]$MetaData.$key);
            } # end try
            catch {
                $errMsg='Failed to convert data to Int32';
                writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
            }; # end catch             
        }; # end else
     }; # end foreach
     return $tmpData;
}; # end function loadMetadata

function setCfgAsDefault
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)][string]$Template,
      [Parameter(Mandatory = $true, Position = 1)][string]$Configuration,
      [Parameter(Mandatory = $false, Position = 2)][switch]$SetAsDefault=$false
     )

    try {
        $errMsg=('Failed to read the template ' + $Template);
        writeLogOutput -LogString ('Loading template ' + $Template);
        $templateData=Get-SecretInfo -Vault $Script:vaultName -Name $Template -ErrorAction Stop;
        writeLogOutput -LogString ('Setting configurtion ' + $Configuration + ' as default.');
        $mData=(loadMetadata -MetaData $templateData.Metadata);
        if ($SetAsDefault.IsPresent)
        {
            $mData.($Script:MetadataDefaultCfg)=$Configuration;
        } # end if
        else {
            if ($mData.($Script:MetadataDefaultCfg) -eq $Configuration)
            {
                $mData.($Script:MetadataDefaultCfg)='';
            } # end if
            else {
                writeLogOutput -LogString ('Cannot remove configuration ' + $Configuration.Split('_')[3] + ' as default configuration. The configuration is currently not the default configuration.' ) -LogType Warning;
                return;
            }; # end else
        }; # end if
        $errMsg=('Failed to save the template ' + $Template);
        writeLogOutput -LogString ('Saving template ' + $Template);
        $mData.($Script:ChangeDate)=([System.DateTime]::Now);
        Set-SecretInfo -Vault $Script:vaultName -Name $Template -Metadata $mData -ErrorAction Stop;
    } # end try
    catch {
        writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
    }; # end catch
}; # end function setCfgAsDefault

function getListOfCfgParameters
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)][array]$EntryList
     )

    $pExcludeList=@(
        $Script:CreateDate,
        $Script:ChangeDate,
        $Script:MetadataDesc,
        $Script:MetadataPList,
        $Script:ConfigurationBackLink,
        $Script:MetadataDefaultCfg,
        $Script:ObjectName
    ); # end pExcludeList
    [array]$rvList=[System.Linq.Enumerable]::Except([string[]]$EntryList,[string[]]$pExcludeList);    
    return $rvList;
}; # end function getListOfCfgParameters

function testDefaultValue
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)][string]$DefaultValue,
      [Parameter(Mandatory = $true, Position = 1)][string]$DataType,
      [Parameter(Mandatory = $true, Position = 2)][string]$ParameterName
     )    

    if (! ($DefaultValue -as [System.Type]$DataType))
    {
        writeLogOutput -LogString ('The value for the default value <' + $DefaultValue + '> has not the expected data type ' + $DataType) -LogType Error;
        return $false;
    };

    if ($DefaultValue.Contains(($Script:sepVal)) -or $DefaultValue.Contains("'"))
    {
        writeLogOutput -LogString ('The caracters ; and ' + '''' + ' are not allowed in the default value') -LogType Error;
        return $false;
    }; # end if
    verifyDefaultValue -DefaultValue $DefaultValue -ParamName $ParameterName
}; # end function testDefaultValue

function isTemlateRegistered
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)][string]$TemplateName
     )

    try {
        $errMsg='Failed to load module configuration';
        writeLogOutput -LogString ('Loading module configuration')
        $moduleCfg=loadMetadata -MetaData (Get-SecretInfo -Vault $Script:vaultName -Name $script:ModuleCfgName).Metadata;        
        if ([System.Linq.Enumerable]::Contains([string[]](($moduleCfg.RegisteredTemplateList).Split(($Script:sepVal))),$TemplateName))
        {
            return $true;
        } # end if
        else {
            writeLogOutput -LogString ('The template ' + $TemplateName + ' is not registered. Data maybe inconsistent.') -LogType Warning;
            return $false;
        }; # end else
    } # end try
    catch {
        writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
    }; # end catch
}; # end funciton isTemlateRegistered

function writeToModuleConfig
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)][hashtable]$CfgData
     )
    $CfgData.($script:ChangeDate)=[System.DateTime]::Now;
    Set-SecretInfo -Vault $script:vaultName -Name $script:ModuleCfgName -Metadata $CfgData -ErrorAction Stop;
}; # end function writeToModuleConfig

function isConfigDefault
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)][string]$ConfigName,
      [Parameter(Mandatory = $true, Position = 1)][bool]$ForceSwitch
     )  
    
    try {
        $errMsg=('Failed to read the configuration ' + $ConfigName);
        writeLogOutput -LogString ('Reading configuration ' + $ConfigName);
        $cfgMetadata=loadMetadata -MetaData ((Get-SecretInfo -Vault $script:vaultName -Name $ConfigName -ErrorAction Stop).Metadata);
        $errMsg=('Failed to read the the template ' + $cfgMetadata.($script:BaseTemplate));
        writeLogOutput -LogString ('Reading configuration ' + $cfgMetadata.($script:BaseTemplate));
        $templateMetadata=loadMetadata -MetaData ((Get-SecretInfo -Vault $script:vaultName -Name $cfgMetadata.($script:BaseTemplate) -ErrorAction Stop).Metadata);
        if ($configName -eq ($templateMetadata.($Script:MetadataDefaultCfg)))
        {
            if (! $ForceSwitch)
            {
                $cfgArr=$ConfigName.Split($Script:AttribNameSep);
            writeLogOutput -LogString ('The configuration ' + $cfgArr[4] + ' is configured as default configuration for template ' + $cfgArr[2] + $Script:EntrySep + $cfgArr[3] + '. To perform the action use the Force parameter.') -LogType Warning;
            }; # end if
            
            return $true;
        }; # end if
        return $false;
    } # end try
    catch {
        writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        return $true;
    }; # end catch
}; # end function isConfigDefault

function convertVersion
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)][Version]$Version
     )   
     
    try {
        $tmpVer=([math]::Max($Version.Major,0).ToString()+'.'+[math]::Max($Version.Minor,0).ToString()+'.'+[math]::Max($Version.Build,0).ToString()+'.'+[math]::Max($Version.Revision,0).ToString())
        return $tmpVer.ToString()
    } # end try
    catch {
        Throw('Failed to convert verstion');        
    }; # end catch
    
}; # end funciton convertVersion

function verifyHelpMessage
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)][string]$Helpmessage,
      [Parameter(Mandatory = $true, Position = 1)][string]$ParamName
     )  

    if (!($Helpmessage -match $Script:RxHelpMsg))
    {       
        Throw('The parameter ' + $ParamName + ' contains characters not allowed for a help message. Only characters matching the ' +  $Script:RxHelpMsg.ToString() + ' are allowed.');
    }; # end else

}; # end function verifyHelpMessage

function verifyDefaultValue
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)]$DefaultValue,
      [Parameter(Mandatory = $true, Position = 1)][string]$ParamName
     )  

    if (([System.Type]'String' -eq $DefaultValue.GetType()) -and (!($DefaultValue -match $Script:RxDefValue)))
    {       
        Throw('The parameter ' + $ParamName + ' contains characters not allowed for a help message. Only characters matching the ' +  $Script:RxHelpMsg.ToString() + ' are allowed.');
    }; # end else

}; # end function verifyDefaultValue