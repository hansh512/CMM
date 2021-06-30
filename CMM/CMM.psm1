################################################################################
# Code Written by Hans Halbmayr
# Created On: 02.06.2021
# Last change on: 29.06.2021
#
# Module: CMM
#
# Version 0.90
#
# Purpos: Manage secrets and secret metadata
################################################################################
Set-StrictMode -Version 3.0
[CmdLetBinding()]
# build list of files to load
$filesToLoad=@(    
    ($PSScriptroot + '\\Libs\HelperFunctions_1.ps1'),
    ($PSScriptroot + '\\Libs\dynParam.ps1'),
    ($PSScriptroot + '\Libs\HelperFunctions_2.ps1'),
    ($PSScriptroot + '\Libs\Functions.ps1')
); # end filesToLoad

# set const for module name
New-Variable -Name 'moduleName' -Value (([System.IO.Path]::GetFileNameWithoutExtension(($MyInvocation.MyCommand.path))).ToUpper()) -Scope Script -Option Constant;

#region init vars and constants
New-Variable -Name 'Prfx' -Value 'ZZZ' -Scope Script -Option Constant; # prefix for entries in vault
New-Variable -Name 'EntrySep' -Value '_' -Scope Script -Option Constant; # separator for names of entries in vault
New-Variable -Name 'AttribNameSep' -Value '_' -Scope Script -Option Constant; # separator for attribute names in metadata
New-Variable -Name 'cfgNamePrefix' -Value ($Script:Prfx+'5c4170ca9d934808b860c772d87074c1') -Scope Script -Option Constant; # prefix for config entries
New-Variable -Name 'vaultName' -Value ($Script:cfgNamePrefix) -Scope Script -Option Constant; # name of vault used
New-Variable -Name 'SecureVaultModuleName' -Value 'Microsoft.PowerShell.SecretStore' -Scope Script -Option Constant; #  only Microsoft.PowerShell.SecretStore supported
New-Variable -Name 'ModuleCfgName' -Value ($Script:cfgNamePrefix + $Script:EntrySep+'ModuleCfg') -Scope Script -Option Constant; #  name of cfg for module
New-Variable -Name 'ModuleCfgVersion' -Value 1 -Scope Script -Option Constant; #  version of the module configuration
New-Variable -Name 'RegisteredModulePrefix' -Value ($Script:Prfx + $Script:EntrySep) -Option Constant -Scope Script;
New-Variable -Name 'MetadataPList' -Value ($Script:cfgNamePrefix + $Script:EntrySep + 'PList') -Option Constant -Scope Script; # list of parameters, for vality check
New-Variable -Name 'MetadataDesc' -Value ($Script:cfgNamePrefix  + $Script:EntrySep + 'Desc') -Option Constant -Scope Script; # description template
New-Variable -Name 'ConfigDesc' -Value ($Script:cfgNamePrefix  + $Script:EntrySep + 'Description') -Option Constant -Scope Script; # description config
New-Variable -Name 'CredVarName' -Value ($Script:cfgNamePrefix  + $Script:EntrySep + 'CredVarName') -Option Constant -Scope Script;
New-Variable -Name 'MetadataDefaultCfg' -Value ($Script:cfgNamePrefix  + $Script:EntrySep + 'DefCfg') -Option Constant -Scope Script; # name of the default config
New-Variable -Name 'VerStr' -Value 'Ver:' -Scope Script -Option Constant;
New-Variable -Name 'ConfigurationBackLink' -Value ($Script:cfgNamePrefix  + $Script:EntrySep + 'LinkedConfigs') -Option Constant -Scope Script; # attrib for linked cfg in template
New-Variable -Name 'BaseTemplate' -Value ($Script:cfgNamePrefix  + $Script:EntrySep + 'BaseTemplate') -Option Constant -Scope Script;
New-Variable -Name 'TemplatePrefix' -Value ($Script:RegisteredModulePrefix+'Template'  + $Script:EntrySep) -Option Constant -Scope Script;
New-Variable -Name 'ObjectName' -Value ($Script:cfgNamePrefix  + $Script:EntrySep + 'Name') -Option Constant -Scope Script;
New-Variable -Name 'ConfigPrefix' -Value ($Script:RegisteredModulePrefix+'Cfg' + $Script:EntrySep) -Option Constant -Scope Script;
New-Variable -Name 'ReturnConfigName' -Value ($script:ConfigPrefix+'CMM_ConfigName') -Option Constant -Scope Script;  # name of the attribute in hash where the config name is stored (when queried from other PS module)
New-Variable -Name 'CreateDate' -Value ($Script:cfgNamePrefix + $Script:AttribNameSep + 'whenCreated') -Option Constant -Scope Script;
New-Variable -Name 'ChangeDate' -Value ($Script:cfgNamePrefix + $Script:AttribNameSep + 'whenChanged') -Option Constant -Scope Script;
New-Variable -Name 'CfgAllowPush' -Value ($Script:cfgNamePrefix + $Script:AttribNameSep + 'AllowPush') -Option Constant -Scope Script;
New-Variable -Name 'cfgVerString' -Value ($Script:cfgNamePrefix + $Script:AttribNameSep + 'Version') -Option Constant -Scope Script;

New-Variable -Name 'HelpMsgString' -Value ('Helpmessage') -Option Constant -Scope Script;
New-Variable -Name 'DefValString' -Value ('DefaultValue') -Option Constant -Scope Script;
New-Variable -Name 'MandatoryString' -Value ('IsMandatory') -Option Constant -Scope Script;
New-Variable -Name 'sepVal' -Value ([char]166) -Option Constant -Scope Script;      # separator for config entries
New-Variable -Name 'blSepVal' -Value ([char]166) -Option Constant -Scope Script;    # separator for backlinks (config links)
New-Variable -Name 'pLstSepVal' -Value ([char]166) -Option Constant -Scope Script;    # separator for parameter list
New-Variable -Name 'pPropSepVal' -Value ([char]166) -Option Constant -Scope Script;    # separator for parameter properties (type,madatory,helpmsg,...)

New-Variable -Name ModuleCfg -Scope Script;
New-Variable -Name 'mCfgVer' -Value 1 -Scope Script -Option Constant; # version for module configuration
New-Variable -Name 'tCfgVer' -Value 1 -Scope Script -Option Constant; # version template configuration
New-Variable -Name 'paramPrefixSep' -Value ([char]172) -Scope Script -Option Constant # parameter prefix seperator (used in metadata)
New-Variable -Name 'paramSepDynVar' -Value '_' -Scope Script -Option Constant # parameter prefix seperator (used in metadata)
New-Variable -Name 'credParamPrefix' -Value ('000'+$script:paramPrefixSep) -Scope Script -Option Constant; # prefix for cred var (used in template config)
New-Variable -Name 'hostParamPrefix' -Value ('001'+$script:paramPrefixSep) -Scope Script -Option Constant; # prefix for host var (used in template config)
#endregion init vars and constants

#region regex constans
$VarNameRxStr='^(?i)(?!('+$script:Prfx+'))[a-zA-Z0-9_]+$'; # regex for variable/parameter name
New-Variable -Name 'RxPSCredVarName' -Value ([regex]::new($VarNameRxStr)) -Scope Script -Option Constant;   # PSCred var
New-Variable -Name 'RxHostVarName' -Value ([regex]::new($VarNameRxStr)) -Scope Script -Option Constant;     # host var
New-Variable -Name 'RxVarName' -Value ([regex]::new($VarNameRxStr)) -Scope Script -Option Constant;         # div vars (Add-CMMModuleTemplateVariable)

$HelpMsgRxStr='^[a-zA-Z0-9 ,.=\d-<>();?@]+$'; # regex for help message
New-Variable -Name 'RxHelpMsg' -Value ([regex]::new($HelpMsgRxStr)) -Scope Script -Option Constant;           # help message (for all vars/params)

$cfgRxStr='^[a-zA-Z0-9'']+$'; # regex for configuration name
New-Variable -Name 'RxCfgName' -Value ([regex]::new($cfgRxStr)) -Scope Script -Option Constant;               # config name (New-CMMConfiguration)

$defVAl='^[a-zA-Z0-9 ,.=\d-<>()@]+$';
New-Variable -Name 'RxDefValue' -Value ([regex]::new($defVAl)) -Scope Script -Option Constant;               # default value
#endregion regex constants

#region debug var
New-Variable -Name 'ShowIntConversionWarning' -Value $false -Scope Script;
#endregion debug var

foreach ($file in $filesToLoad)
{
    try
    {
    . $file;
    } # end try
    catch
    {
        Write-Warning ('Failed to load the file ' + $file);
        Write-Host ($_.Exception.Message);
        #return;
    }; # end catch
}; # end foreach

# search for vault

initModule;
initEnumerator;
New-Variable -Name 'ExportVarPrfx' -Value '__' -Scope Script -Option Constant;
New-Variable -Name ($script:ExportVarPrfx + $script:ModuleName +'_ModuleData') -Value ([GetModuleData]::new());  # init class
Export-ModuleMember -Variable ('__' + $script:ModuleName +'_ModuleData'); # export object with methods and properties

$cmdList=@(
    'Get-Configuration',
    'Get-ModuleTemplate',
    'New-Configuration',
    'Set-Configuration',
    'Remove-Configuration',
    'Unregister-ModuleTemplate'
    'Register-ModuleTemplate',
    'Remove-ModuleTemplateVariable'
    'Add-ModuleTemplateVariable',
    'Set-ModuleTemplate'
); # end cmdList
Export-ModuleMember $cmdList;
