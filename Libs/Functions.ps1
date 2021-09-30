################################################################################
# Code Written by Hans Halbmayr
# Created On: 28.05.2021
# Last change on: 01.07.2021
#
# Module: CMM
#
# Version 0.91
#
# Purpos: Function for module CMM
################################################################################


#region template
function Register-ModuleTemplate
{
<#
.SYNOPSIS 
Registers a new template for a PowerShell module or script.
.DESCRIPTION
The command is used to register a template which is used to create appropriate configurations for PowerShell module or scripts. The template determine which variables are available. For most of the variables the following properties are available:
-	Name
-	Data type
-	Default value (not for the variable defined with the parameter PSCredentialVarName)
-	Help message
.PARAMETER ModuleName
The parameter is mandatory. The data type is string.
The name of the PowerShell module or script.
.PARAMETER MinModuleVersion
The parameter is optional. Data type version.
The minimum version of the module which should use configurations created from this template. If the parameter is omitted, a version of 0.0.0.0 is assumed. If you have multiple versions of PowerShell modules, which need different configurations, you can create, the appropriate templates (module name and minimum version). 

.PARAMETER TemplateDescription
The parameter is optional. Data type string.
Enter a description for the template. If the parameter is omitted, the description wil. be blank.

.PARAMETER PSCredentialVarName
The parameter is mandatory. Data type PSCredential
The name for the variable which stores the credential in a module or script. When a configuration is created, the parameter for the credential will be named according the value of this parameter.
When a module or script queries a configuration, the module will provide a hashtable with the element Data. The name defined with this parameter, will be an element of Data (data type PSCredential). 
The PSCredentialVarName must not start with ZZZ or zzz. Allowed characters:
-	a-z
-	A-Z
-	0-9
-	_

.PARAMETER PSCredentialVarHelpMessage
The parameter is optional. Data type string.
The parameter expects a help message for the parameter created with PSCredentialVarName. The following characters are allowed: 
-	a-z
-	A-Z
-	0-9
-	,
-	.
-	-
-	<
-	>
-	(
-	)
.PARAMETER HostVarName
The parameter is mandatory. Data type string.
The name for the variable which stores the host name or URL in a module or script. When a configuration is created, the parameter for the credential will be named according the value of this parameter.
When a module or script queries a configuration, the module will provide a hashtable with the element Data. The name defined with this parameter, will be an element of Data (data type string). 
The HostVarName must not start with ZZZ or zzz. Allowed characters:
-	a-z
-	A-Z
-	0-9
-	_


.PARAMETER HostVarDefaultValue
The parameter is optional. Data type string.
The parameter expects a default value for the parameter created with HostVarName. A default value makes sense if most of the configuration created with the template connects to the same host. The following characters are allowed: 
-	a-z
-	A-Z
-	0-9
-	,
-	.
-	-
-	<
-	>
-	(
-	)

.PARAMETER HostVarHelpMessage
The parameter is optional. Data type string.
The parameter expects a help message for the parameter created with HostVarName. The following characters are allowed: 
-	a-z
-	A-Z
-	0-9
-	,
-	.
-	-
-	<
-	>
-	(
-	)

.PARAMETER SkipModuleValidation
The parameter is optional. Data type switch (Boolean).
The command tries to find the module, provided with the parameter ModuleName (Get-Module <module name> - ListAvailable). If the module can be found, the command will fail. To register a template for module which cannot be found or a script, this parameter can be used. 
.EXAMPLE
Register-CMMModuleTemplate -ModuleName PSTst -MinModuleVersion 0.0.0.0 -TemplateDescription 'Template for module test' -PSCredentialVarName PSTCred -PSCredentialVarHelpMessage 'Enter your credential (mail address and password)' -HostVarName PSTHost -HostVarDefaultValue 'vhost.domain.name' -HostVarHelpMessage 'Enter the FQDN for the host'
A template for the module PSTst will be created. Configurations created from this template will use the default value <vhost.domain.name> for the host name. 
.EXAMPLE
Register-CMMModuleTemplate -ModuleName PSTst -MinModuleVersion 2.0.0.0 -TemplateDescription 'Template for module test' -PSCredentialVarName PSTCred -PSCredentialVarHelpMessage 'Enter your credential (mail address and password)' -HostVarName PSTHost -HostVarDefaultValue 'vhost.domain.name' -HostVarHelpMessage 'Enter the FQDN for the host'
A template for the module PSTst will be created. Configurations created from this template will be used with module versions of 2.0.0.0 or higher. Configurations created from this template will use the default value <vhost.domain.name> for the host name.
#>
[cmdletbinding()]    
param([Parameter(Mandatory = $true, Position = 0)][string]$ModuleName,
      [Parameter(Mandatory = $false, Position = 1)][Version]$MinModuleVersion='0.0.0.0',
      [Parameter(Mandatory = $false, Position = 2)][string]$TemplateDescription='',
      [Parameter(Mandatory = $true, Position = 3)][ValidateScript({$_ -match $Script:RxPSCredVarName})][string]$PSCredentialVarName,
      [Parameter(Mandatory = $false, Position = 4)][ValidateScript({$_ -match $Script:RxHelpMsg})][string]$PSCredentialVarHelpMessage='',
      [Parameter(Mandatory = $true, Position = 5)][ValidateScript({$_ -match $Script:RxHostVarName})][string]$HostVarName,
      [Parameter(Mandatory = $false, Position = 6)][string]$HostVarDefaultValue='',
      [Parameter(Mandatory = $false, Position = 7)][ValidateScript({$_ -match $Script:RxHelpMsg})][string]$HostVarHelpMessage='',
      [Parameter(Mandatory = $false, Position = 8)][switch]$SkipModuleValidation=$false
     )

    begin {
        $errMsg='Unhandled exeption';
        if (! (verifyModuleName -ModuleName $ModuleName -TextString 'module name'))
        {
            throw('The name of the module contains characters which are not supported by the module ' + $Script:moduleName);
        }; # end if
        if ($PSBoundParameters.ContainsKey('HostVarDefaultValue')  -and (! (testDefaultValue -DefaultValue $HostVarDefaultValue -DataType 'String'  -ParameterName 'HostVarDefaultValue')))
        {
            return;
        }; # end if
    }; # end begin

    process {        
        try {        
            $msg='Testing if module already registerd';
            $errMsg=('Failed to get data for PowerShell module ' + $ModuleName)
            $MinModuleVersion=(convertVersion -Version $MinModuleVersion);
            $ModuleName=($ModuleName.ToUpper());
            $errMsg='Failed to verify it the template is already registerd.';
            writeLogOutput -LogString $msg;
            if (Get-SecretInfo -Vault $Script:vaultName -Name ($script:RegisteredModulePrefix + '*' + $Script:EntrySep + $ModuleName + $Script:EntrySep + $MinModuleVersion))
            {
                writeLogOutput -LogString ('Failed to register template. A template for the module ' + $ModuleName + ' and the version ' + $MinModuleVersion + ' is already registered.') -LogType Warning;
                return;
            }; # end if
            if (!($SkipModuleValidation.IsPresent))
            {
                $errMsg=('Failed to get data for PowerShell module ' + $ModuleName);
                writeLogOutput -LogString ('Looking up for module ' + $ModuleName);
                $mData=(Get-Module -Name $ModuleName -ListAvailable);                                
                if (!($mData))
                {
                    writeLogOutput -LogString ('Module ' + $ModuleGUID + ' not found. Cannot register module data.') -LogType Warning;
                    return;
                }; # end if                
            }; # end if
            #>        
            $regString=($script:RegisteredModulePrefix + 'Template_'+$ModuleName+$Script:EntrySep+$MinModuleVersion);           
            $msg='Building metadata'
            writeLogOutput -LogString $msg;
            $d=([System.DateTime]::Now);
            [int]$madatoryHostVal=$true; #([System.String]::IsNullOrEmpty($HostVarDefaultValue)); 
            # add 11 empty entries for future use to cred and hostname
            $spareEntries=$Script:pPropSepVal; #+$Script:pPropSepVal+$Script:pPropSepVal+$Script:pPropSepVal+$Script:pPropSepVal+$Script:pPropSepVal;
            for ($i=0;$i -lt 10;$i++)
            {
                $spareEntries+=$Script:pPropSepVal;
            }; # end for
            $metaData=@{
                (($Script:credParamPrefix)+$PSCredentialVarName)=('PSCredential'+$Script:pPropSepVal+'1'+$Script:pPropSepVal+$PSCredentialVarHelpMessage+($Script:pPropSepVal)+$spareEntries);
                (($Script:hostParamPrefix)+$HostVarName)=('String'+$Script:pPropSepVal+$madatoryHostVal+($Script:pPropSepVal)+$HostVarHelpMessage+($Script:pPropSepVal)+$HostVarDefaultValue+$spareEntries);
                ($script:ObjectName)=($ModuleName+$Script:EntrySep+$MinModuleVersion); # name of the template
                (($script:metadataPList))=($PSCredentialVarName+(($Script:pLstSepVal))+$HostVarName); #write list of vars used to attribute
                (($script:metadataDesc))=($TemplateDescription);
                (($script:metadataDefaultCfg))='';
                ($script:ConfigurationBackLink)='';  # stores list of configs using the template
                ($Script:CreateDate)=[dateTime]$d; # date object created
                ($Script:ChangeDate)=[dateTime]$d; # date object changed
                ($Script:cfgVerString)=$script:tCfgVer;
            }; # end netaData
            
            $errMsg=('Failed to register module ' + $ModuleName);
            writeLogOutput -LogString ('Registering template ' + $regString);            
            Set-Secret -Name $regString -Vault $Script:vaultName -Metadata $metaData -Secret ($Script:cfgNamePrefix) -NoClobber -ErrorAction Stop;
            $errMsg=('Failed to read module configuration from ' +$script:ModuleCfgName);
            writeLogOutput -LogString ('Updating module configruation');
            $moduleCfg=loadMetadata -MetaData (Get-SecretInfo -Vault $Script:vaultName -Name $script:ModuleCfgName).Metadata;
            $tmp=($moduleCfg.RegisteredTemplateList).Split(($Script:sepVal));            
            if ([System.String]::IsNullOrEmpty($tmp))
            {
                $moduleCfg.RegisteredTemplateList=$regString;
                $moduleCfg.RegisteredModuleList=$regString.Replace($script:TemplatePrefix,'').split($Script:EntrySep)[0];
            } # end if
            else {
                $tmp+=$regString;
                $tStr= ([System.Linq.Enumerable]::Distinct([string[]]$tmp) -join ($Script:sepVal));
                $moduleCfg.RegisteredTemplateList = $tStr;
                $tmpArr=[System.Collections.ArrayList]::new();
                [void]$tmpArr.AddRange(($moduleCfg.RegisteredModuleList).split($script:sepVal));
                [void]$tmpArr.Add($regString.Replace($Script:TemplatePrefix,'').Split(($Script:EntrySep))[0]);
                $moduleCfg.RegisteredModuleList = $tmpArr -join ($Script:sepVal);
            }; # end else            
            $errMsg=('Failed to write module configuration to ' + $script:ModuleCfgName)
            writeToModuleConfig -CfgData $moduleCfg;
        } # end try
        catch {            
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack ($_);
        }; # end catch
    }; # end process

    end {
    }; # end END        
}; # end funciton Register-ModuleTemplate

function Add-ModuleTemplateVariable
{
<#
.SYNOPSIS 
Adds a variable to a template.
.DESCRIPTION
The command adds a variable to a template created with the command Register-CMMModuleTemplate. If one or more configurations, based on the template, exists, the command will fail with a warning and a list of configurations. To add the variable even if configurations exist, the Force switch can be used.
.PARAMETER Template
The parameter is mandatory. Data type string.
The name of the template where variable should be added. The parameter supports tab-complete (ArgumentCompleter).

.PARAMETER VariableName
The parameter is optional. Data type string.
Expects the name of the variable which should be added. Allowed characters are:
-	a-z
-	A-Z
-	0-9
-	_
The name of the variable must not start with ZZZ or zzz.

.PARAMETER DataType
The parameter is optional. Data type string.
The data type for the variable which should be added. The following data types are available
-	String
-	Int32
-	Boolean

.PARAMETER IsMandatory
The parameter is optional. Data type switch.
If the parameter is used, the variable will be marked as mandatory. The parameter cannot be used with the parameter DefaultValue.

.PARAMETER DefaultValue
The parameter is optional. Data type depends of the parameter DataType.
The parameter allows to add a default value for the variable. The parameter cannot be used with the parameter IsMandatory.

.PARAMETER HelpMessage
The parameter is optional. Data type string.
The parameter expects a help message for the parameter created. The following characters are allowed: 
-	a-z
-	A-Z
-	0-9
-	,
-	.
-	-
-	<
-	>
-	(
-	)

.PARAMETER Force
The parameter is optional. Data type switch.
If there are already configurations created from the template, and a new variable should be added the whole configuration will be inconsistent. If the Force switch is used, the variable can be added. 
Warning: The configurations already created from this template could be unusable. 

.EXAMPLE
Add-CMMModuleTemplateVariable -Template PSTST_Ver:0.9.0.0 -VariableName PSTPort -DataType Int32 -DefaultValue 9876 -HelpMessage 'Enter the value for the port, default is 9876'
#>
[cmdletbinding(DefaultParametersetName='__AllParameter')]    
param([Parameter(Mandatory = $true, Position = 0)]
        [ArgumentCompleter( {  
            param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )           
            $mlist=($__CMM_ModuleData.GetModuleAndVerList());
            $mlist.Where({ $_ -like "$wordToComplete*" });              
        } )]    
      [string]$Template,
      [Parameter(Mandatory = $true, Position = 1)][ValidateScript({$_ -match $Script:RxVarName})][string]$VariableName,
      [Parameter(Mandatory = $true, Position = 2)][ValidateSet('String','Int32','Boolean')][string]$DataType,
      [Parameter(ParametersetName='IsMan',Mandatory = $false, Position = 3)][switch]$IsMandatory,
      [Parameter(ParametersetName='DefVal',Mandatory = $false, Position = 4)][string]$DefaultValue,
      [Parameter(Mandatory = $false, Position = 5)][ValidateScript({$_ -match $Script:RxHelpMsg})][string]$HelpMessage,
      [Parameter(Mandatory = $false, Position = 6)][switch]$Force=$false
     )

    begin {
        if (($PSBoundParameters.ContainsKey('DefaultValue')) -and (! (testDefaultValue -DefaultValue $DefaultValue -DataType $DataType -ParameterName $VariableName)))
        {
            return;
        }; # end if
        $errMsg='Unhandled exeption';
    }; # end begin

    process {
        try {
            $templateName=($script:TemplatePrefix + ($Template.Replace(($script:verStr),'')));
            $errMsg=('Failed to read the module template ' + $templateName);
            writeLogOutput -LogString ('Reading module template ' + $templateName);
            if ((! (isTemlateRegistered -TemplateName $templateName)) -and ($Force.IsPresent -eq $false))
            {
                writeLogOutput -LogString ('To perform the action, use the Force parameter.') -LogType Warning;
                return;
            }; # end if
            $TemplateData=(Get-SecretInfo -Name $templateName -Vault ($script:vaultName) -ErrorAction Stop);
            $errMsg='Failed to convert metadata to hashtable'
            writeLogOutput -LogString ('Loading metadata for template ' + $templateName);
            $metaData=(loadMetadata -MetaData $TemplateData.Metadata); # convert readonly dictionary to hashtable
            ## check if name for var is already in use (exist in template)
            $entryKeys=($metaData.keys | Where-Object {! ($_.StartsWith($script:cfgNamePrefix))}); # get keys
            $vList=[system.Collections.ArrayList]::new();
            for ($i=0;$i -lt $entryKeys.count;$i++)
            {
                [void]$vList.Add(($entryKeys[$i].Split($script:paramPrefixSep))[1]); # extract var/param name (remove prefix)
            }; # end for
            if ([system.linq.Enumerable]::Contains([string[]]$vList,$PSBoundParameters.VariableName))
            {
                writeLogOutput -LogString ('A variable with the name ' + ($PSBoundParameters.VariableName + ' already exist.')) -LogType Error;
                return;
            }; # end if
            
            if ((!([System.String]::IsNullOrEmpty($metaData.($script:ConfigurationBackLink))) -and (! $Force.IsPresent)))
            {
                writeLogOutput -LogString ('There are one or more configurations based on the template ' + $Template + '. If you add a variable the configurations are in an inconsistance state. To add the variable use the Force parameter.' + "`n" + 'List of configurations:') -LogType Warning;
                $cfgList=($metaData.($script:ConfigurationBackLink)).split($Script:blSepVal);
                foreach ($cfg in $cfgList)
                {
                    #[System.Linq.Enumerable]::Last([string[]]$cfg.split('_')); # extract backlink data
                    [System.Linq.Enumerable]::Last([string[]]$cfg.split($Script:AttribNameSep)); # extract backlink data
                }; # end if
                return;
            }; # end if
            
            $entryKeys=[System.Linq.Enumerable]::OrderBy([string[]]$entryKeys,[Func[string,string]] { $args[0] });
            $lastEntry=[System.Linq.Enumerable]::Last([string[]]$entryKeys);
            $newEntry=([string](([int]($lastEntry.Split($script:paramPrefixSep)[0]))+1) ).PadLeft(3,'0') + ($script:paramPrefixSep) + $VariableName; # create var/param name
            
            
            $madatoryVal=(($IsMandatory.IsPresent) -and (!($PSBoundParameters.ContainsKey('DefaultValue')))); # make sure, that IsMandatory is set to false if default value is set
            $dataString=($DataType + ($Script:pPropSepVal)+([int]$madatoryVal)).ToString()+($Script:pPropSepVal);
            if ($PSBoundParameters.ContainsKey($Script:HelpMsgString))
            {
                $dataString+=$HelpMessage;
            }; # end if
            if ($PSBoundParameters.ContainsKey($Script:DefValString))
            {
                $dataString+=($Script:pPropSepVal)+($DefaultValue);
            } # end if
            else {
                $dataString+=($Script:pPropSepVal);
            }; # end else
            # add 11 empty entries for future use
            $spareEntries=$Script:pPropSepVal; #+$Script:pPropSepVal+$Script:pPropSepVal+$Script:pPropSepVal+$Script:pPropSepVal+$Script:pPropSepVal;
            for ($i=0;$i -lt 10;$i++)
            {
                $spareEntries+=$Script:pPropSepVal;
            }; # end for
            $metaData.Add($newEntry,($dataString+$spareEntries));

            $pList=($metaData.(($script:metadataPList))).Split($Script:pPropSepVal);
            $pList+=$VariableName;             
            $metaData.($script:metadataPList)=($pList -join $Script:pPropSepVal);            
            $metaData.($script:ChangeDate)=[datetime]([System.DateTime]::Now);
            $errMsg='Failed to add variable';            
            writeLogOutput -LogString ('Adding variable ' + $VariableName + ' to module template ' + $templateName);            
            Set-SecretInfo -Name ($TemplateData.name) -Vault ($script:vaultName) -Metadata $metaData -ErrorAction Stop;
        } # end try
        catch {
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        }; # end catch               
    }; # end process

    end {

    }; # end END
}; # end Add-ModuleTemplateVariable

function Remove-ModuleTemplateVariable
{
<#
.SYNOPSIS 
Removes a variable from a template.
.DESCRIPTION
The command removes a variable from a template. Only variables created with the command Add-CMMModuleTemplateVariable can be removed.

.PARAMETER Template
The parameter is mandatory. Data type string.
The name of the template where variable should be removed. The parameter supports tab-complete (ArgumentCompleter).

.PARAMETER VariableName
The parameter is mandatory. Data type string.
Name of the variable to remove. Valid names are variable names added with the command Add-CMMModuleTemplateVariable.

.PARAMETER Force
The parameter is optional. Data type switch
If there exist configurations created from that template, the command will fail with a warning. To remove a variable anyway, the switch Force can be used. If a variable is removed with the switch Force, the configuration is in an inconsistent state.
The parameter is optional.

.EXAMPLE
Remove-CMMModuleTemplateVariable -Template PSTST_Ver:0.9.0.0 -VariableName PSTPort
#>
[cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]    
param([Parameter(Mandatory = $true, Position = 0)]
        [ArgumentCompleter( {  
            param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )           
            $mlist=($__CMM_ModuleData.GetCfgOrTemplateList($null,$true)); # get list of templates
            $mlist.Where({ $_ -like "$wordToComplete*" });              
        } )]    
      [string]$Template,
      [Parameter(Mandatory = $false, Position = 1)][switch]$Force=$false
     )

    DynamicParam
    {                    
        if ($Template) 
        {  
            $mData=loadMetadata -MetaData ((Get-SecretInfo -Vault $script:vaultName -Name ($script:TemplatePrefix +  $Template) -ErrorAction Stop).Metadata);
            $varList=[System.Collections.ArrayList]::new();
            foreach ($entry in $mData.Keys)
            {
                If (!((($entry).StartsWith($script:cfgNamePrefix+$Script:AttribNameSep) -or (($entry).StartsWith((($Script:credParamPrefix)))) -or (($entry).StartsWith(($Script:hostParamPrefix)))))) # getting list of vars
                {
                    [void]$VarList.Add(($entry.Split($script:paramPrefixSep))[1]);
                }; # end if
            }; # end foreadch
            if ($varList.count -gt 0)
            {                
                try {
                    $dynParamList=@{
                    ParamName='VariableName';
                    Position=1;
                    IsMandatory=$True;
                    ValidateSet = $varList;
                } # end dynParamList           
                newDynamicParam @dynParamList              
            } # end try
            catch {
            Write-Host 'Failed to build parameters' -ForegroundColor Red -BackgroundColor Black;
            }; # end catch 
            } ;  
        }; # end if               
    }; # end DynamicParam

    begin {
        $errMsg='Unhandled exeption';
    }; # end begin

    process {
        try {
            if (! ($PSBoundParameters.ContainsKey('VariableName')))
            {
                writeLogOutput -LogString 'There is no variable to remove.' -LogType Error;
                return;
            }; # end if
            # variable name is case sensitive make sure 
            $pos= $varList.ToLower().IndexOf($PSBoundParameters.VariableName.ToLower())
            $PSBoundParameters.VariableName=$varList[$pos];             
            $rawTemplateName=$Template.Replace(($script:verStr),'');
            $errMsg=('Failed to read the module template ' + $rawTemplateName);
            writeLogOutput -LogString ('Reading module template ' + $Template);
            $moduleTemplateData=loadMetadata -MetaData ((Get-SecretInfo -Name ('*' + $rawTemplateName) -Vault ($script:vaultName) -ErrorAction Stop).Metadata);     
            if (!([system.String]::IsNullOrEmpty($moduleTemplateData.($Script:ConfigurationBackLink))) -and ($Force.isPresent -eq $false))
            {                
                writeLogOutput -LogString $('For the template ' + $Template + ' exist one or more configuration.'+ "`n"  + 'Configuraton created from this template will be in an inconsistent state. To remove the variable, use the -Force' + "`n" + 'Configuration(s)') -LogType Warning;
                $moduleTemplateData.($Script:ConfigurationBackLink).split($Script:blSepVal);
                return;
            }; # end if
            
            if($PSCmdlet.ShouldProcess('Performing on ',('Variable ' + $PSBoundParameters.VariableName),'REMOVE')) # ask user for confirmation
            {
                writeLogOutput -LogString ('Removing variable ' + $PsBoundParameters.VariableName);
                $mData=($ModuleTemplateData);
                $errMsg=('Failed to remove the variable ' + $PSBoundParameters.VariableName + ' from metadata');
                $vName=($PSBoundParameters.VariableName);
                #[Func[string,bool]] $p={ param($e); return ($e.EndsWith('_'+$vName) -and -not $e.StartsWith(($script:cfgNamePrefix)+'_'))}; # prepare linq query
                #[Func[string,bool]] $p={ param($e); return ($e.EndsWith($Script:paramSepDynVar+$vName) -and -not $e.StartsWith(($script:paramSepDynVar)+$Script:paramPrefixSep))}; # prepare linq query
                [Func[string,bool]] $p={ param($e); return ($e.EndsWith($Script:paramPrefixSep+$vName) -and -not $e.StartsWith(($script:paramPrefixSep)+$Script:paramPrefixSep))}; # prepare linq query
                $varRawName=[System.Linq.Enumerable]::Where(([string[]]($mData.Keys)), $p); # get raw name of var                
                $errMsg=('Failed to save the configuration for template ' + $rawTemplateName);
                writeLogOutput -LogString ('Removing variable ' + [string]$varRawName);
                $pList=[System.Collections.ArrayList]::new();  # prep array
                $pList.AddRange(@($mData.($script:MetadataPList)).Split($Script:pLstSepVal)); # get list of params from metadata
                [void]$pList.Remove($PSBoundParameters.VariableName); # remove var from list
                $mData.($script:MetadataPList) = ($pList -join $Script:pLstSepVal); # set new pList string
                $mdata.Remove([string]$varRawName[0]); # remove entry (var/param) from hashtable 
                $mdata.($script:ChangeDate)=[datetime]([System.DateTime]::Now);        
                writeLogOutput -LogString ('Saving configuration for template ' + $rawTemplateName);      
                Set-SecretInfo -Vault ($Script:vaultName) -Name ($script:TemplatePrefix + $rawTemplateName) -MetaData $mdata -ErrorAction Stop; # save new config to entry                
            }; # end if            
        } # end try
        catch {
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        }; # end catch               
    }; # end process

    end {

    }; # end END
}; # end Remove-ModuleTemplateVariable

function Set-ModuleTemplate
{
<#
.SYNOPSIS 
Changes attributes of a variable.
.DESCRIPTION
The command updates attributes of a variable of a registered template. Only the parameter template is known. The remaining parameters will be created at runtime of the command.
For the variable, which should store the credential the attribute for the help message can be updated. The parameter will be named as follows:
-	<CredVarName>_Helpmessage
Example: If the variable for the credential is named TstCred, the name of the parameter will be TstCred_Helpmessage. The data type is string.
The variable, which should store the host name the attributes for
-	Help message
-	Default value
can be updated.
Example: If the variable for the host name is named TstHost, the name of the parameter will be 
-	TstHost_Helpmessage (data type string)
-	TstHost_DefaultValue (data type string)
For the variables added with the command Add-CMMModuleTemplateVariable the following attributes can be changed:
-	Help message
-	Default value
-	IsMandatory
Example: If the variable TstPort was added, the name of the parameter will be 
-	TstPort_Helpmessage (data type string)
-	TstPort_DefaultValue (data type depends of the data type configured)
-	TstPort_IsMandatory (data type boolean)

.PARAMETER Template
The parameter is mandatory. Data type string.
The name of the template where variable should be removed. The parameter supports tab-complete (ArgumentCompleter).
.PARAMETER Description
The parameter is optional. Data type string.
Update the description for the template.

.EXAMPLE
Set-CMMModuleTemplate -Template TEST_0.0.0.0 -Description 'Test template'

.EXAMPLE
Set-CMMModuleTemplate -Template TEST_0.0.0.0 -TstHost_DefaultValue myhost.company.com

. EXAMPLE
Set-CMMModuleTemplate -Template OLX_0.0.0.0 -OLXTenantName_Helpmessage 'Enter the name of the tenant' -OLXSendReportsTo_DefaultValue 'tim@domain.com'
#>   
[cmdletbinding()]    
param([Parameter(Mandatory = $true, Position = 0)]
        [ArgumentCompleter( {  
            param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )           
            $mlist=($__CMM_ModuleData.GetCfgOrTemplateList($null,$true)); # get list of templates
            $mlist.Where({ $_ -like "$wordToComplete*" });              
        } )]    
      [string]$Template,      
      [Parameter(Mandatory = $false, Position = 1)][String]$Description
     )

    DynamicParam
    {                    
        if ($Template) 
        {  
            $mData=loadMetadata -MetaData ((Get-SecretInfo -Vault $script:vaultName -Name ($script:TemplatePrefix +  $Template) -ErrorAction Stop).Metadata);
            $dynParamHash=@{};
            $keyList=[System.Collections.ArrayList]::new();
            #$keyList.AddRange([array]($mData.Keys).where({! $_.StartsWith($script:cfgNamePrefix+'_')}));
            $keyList.AddRange([array]($mData.Keys).where({! $_.StartsWith($script:cfgNamePrefix+$Script:AttribNameSep)}));
            for ($i=0;$i -lt $keyList.Count;$i++)
            {
                $dynParamHash.Add($keyList[$i],$mData.($keyList[$i]));
            }; # end for
            
            if ($keyList.count -gt 0)
            {                
                try {                    
                    setTemplateDynamicParamList -ParamHash $dynParamHash  -FirstParamPosition 2;   
                } # end try
                catch {
                    Write-Host 'Failed to build parameters' -ForegroundColor Red -BackgroundColor Black;
                }; # end catch 
            }; # end if 
        }; # end if               
    }; # end DynamicParam

    begin {
        $errMsg='Unhandled exeption';
    }; # end begin

    process {
        try {                    
            $BoundParams = $PSBoundParameters;
            $paramsToExclude=[System.Collections.ArrayList]::new();
            [void]$paramsToExclude.AddRange(@('Template','Configuration','Description'))
            [void]$paramsToExclude.AddRange(@([System.Management.Automation.PSCmdlet]::CommonParameters));
            [void]$paramsToExclude.AddRange(@([System.Management.Automation.PSCmdlet]::OptionalCommonParameters));            
            $paramList=[System.Collections.ArrayList]::new();
            $paramList.AddRange([array][System.Linq.Enumerable]::Except([string[]]$BoundParams.Keys,[string[]]$paramsToExclude));            
            $metaData=loadMetadata -Metadata $mData;
            if ($BoundParams.ContainsKey('Description'))
            {
                $metaData.($Script:MetadataDesc)=$BoundParams.Description;            
            }; # end if

            if ($paramList.count -gt 0)
            {            
                writeLogOutput -LogString ('Prepare configuration update for template ' + $Template);
                $errMsg='Failed to update the template.'                
                for ($i=0;$i -lt $paramList.Count;$i++)
                {                    
                    $plArr=$paramList[$i].Split($script:paramSepDynVar); # split dyn var name
                    $itemType=[System.Linq.Enumerable]::Last($plArr);
                    $pName=($plArr[0..($plArr.Count -2)] -join $script:paramSepDynVar);     
                    $errMsg=('Failed to prepare update for ' + $itemType + ' for parameter ' + $pName);
                    [Func[string,bool]] $p={ param($e); return ($e.EndsWith(($script:paramPrefixSep)+$pName))}; # prepare linq query
                    $param=[System.Linq.Enumerable]::Where(([string[]]($keyList)), $p); # get raw name of var                    
                    #$tmpArr=($metaData.$($param)).Split($Script:sepVal); # split config data from metadata
                    $tmpArr=($metaData.$($param)).Split($Script:pPropSepVal); # split config data from metadata
                    switch ($itemType)
                    {
                        ($Script:HelpMsgString) {
                            $helpMsg=$BoundParams.($param.Split($script:paramPrefixSep)[1]+(($script:paramSepDynVar) + $Script:HelpMsgString)); # get help message
                            if (($null -eq $helpMsg) -or ($helpMsg -eq '') -or ([System.String]::IsNullOrEmpty($helpMsg)))
                            {
                                $tmpArr[2]=$null; # set default value
                            } # end if
                            else {
                                verifyHelpMessage -Helpmessage $helpMsg -ParamName ($param.Split(($script:paramPrefixSep))[1]+(($script:paramPrefixSep) + $Script:HelpMsgString)); # verify help messeage (not allowed chars)                            
                                $tmpArr[2]=$helpMsg; # set default value
                            }; # end else                              
                            break;
                        }; # end HelpMsgString
                        ($Script:DefValString)  {
                            $defVal=$BoundParams.($param.Split(($script:paramPrefixSep))[1]+(($script:paramSepDynVar) + $Script:DefValString)); # get default value
                            if (($null -eq $defVal) -or ($defVal -eq '') -or ([System.String]::IsNullOrEmpty($defVal)))
                            {
                                $tmpArr[3]=$null; # set default value
                            } # end if
                            else {
                                verifyDefaultValue -DefaultValue $defVal -ParamName ($param.Split(($script:paramPrefixSep))[1]+(($script:paramSepDynVar) + $Script:DefValString)); # verify default value (not allowed chars)                            
                                $tmpArr[3]=$defVal.ToString(); # set default value
                            }; # end else                            
                            break;
                        }; # end DefValString
                        ($Script:MandatoryString)  {
                            $isMandVal=$BoundParams.($param.Split(($script:paramPrefixSep))[1]+(($script:paramSepDynVar) + $Script:MandatoryString)); # get default value
                            if (($null -eq $isMandVal) -or ($isMandVal -eq '') -or ([System.String]::IsNullOrEmpty($isMandVal)))
                            {
                                $tmpArr[1]=0; # set default value
                            } # end if
                            else {
                                verifyDefaultValue -DefaultValue $isMandVal -ParamName ($param.Split(($script:paramPrefixSep))[1]+(($script:paramPrefixSep) + $Script:MandatoryString)); # verify default value (not allowed chars)                            
                                $tmpArr[1]=[int]$isMandVal; # set default value
                            }; # end else                            
                            break;
                        }; # end MandatoryString
                    }; # end switch
                    #$metaData.$($param) = $tmpArr -join ($Script:sepVal);
                    $metaData.$($param) = $tmpArr -join ($Script:pPropSepVal);
                }; # end for
                
            }; # end if
            $errMsg=('Failed to save template configuration for ' + ($script:TemplatePrefix +  $Template));
            writeLogOutput -LogString ('Updating configuration for template ' + $Template);
            $metaData.($Script:ChangeDate)=([System.DateTime]::Now);
            Set-SecretInfo -Vault $script:vaultName -Name ($script:TemplatePrefix +  $Template) -Metadata $metaData -ErrorAction Stop;
        } # end try
        catch {
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        }; # end catch
    }; # end process

    end {

    }; # end END
}; # function Set-ModuleTemplate
function Get-ModuleTemplate
{
<#
.SYNOPSIS 
Get a list of templates.
.DESCRIPTION
The command lists templates
-	All available templates
-	Templates for a particular module
-	A particular template

.PARAMETER ModuleName
The parameter is optional. Data type string.
Returns a list of templates for a particular module. The parameter supports tab-complete (ArgumentCompleter).
The command cannot be used with the parameter Template.

.PARAMETER Template
The parameter is optional. Data type string.
Returns the data for a particular template. The parameter supports tab-complete (ArgumentCompleter).
The command cannot be used with the parameter ModuleName.

.PARAMETER Format
The parameter is optional. Data type enum OutFormat
If the parameter is omitted the data is returned in table format. The parameter accepts the following formats:
-	Table (default)
-	List
-	PassValue (returns unformatted data)

.EXAMPLE
Get-CMMModuleTemplate
All templates will be returned.
.EXAMPLE
Get-CMMModuleTemplate -ModuleName OLX
Returns templates for the module OLX.
#>
[cmdletbinding(DefaultParametersetName='__AllParameter')]    
param([Parameter(ParameterSetName='Module',Mandatory = $false, Position = 0)][ArgumentCompleter( {
    param ( $CommandName,
        $ParameterName,
        $WordToComplete,
        $CommandAst,
        $FakeBoundParameters )  
        $cfgList=$__CMM_ModuleData.GetModuleList();
        $tmpList=@();
    foreach ($item in $tmpList)
    {
        if ($item.contains(' '))
        {
            $cfgList+="'"+$item+"'";
        } # end if
        else {
            $cfgList+=$item;
        }; # end else
    };
    $cfgList.Where({ $_ -like "$wordToComplete*" });
    } )][string]$ModuleName,
      [Parameter(ParameterSetName='Template',Mandatory = $false, Position = 0)][ArgumentCompleter( {
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
            $cfgList=$__CMM_ModuleData.GetCfgOrTemplateList($null,$true);
            $tmpList=@();
        foreach ($item in $tmpList)
        {
            if ($item.contains(' '))
            {
                $cfgList+="'"+$item+"'";
            } # end if
            else {
                $cfgList+=$item;
            }; # end else
        };        
        $cfgList.Where({ $_ -like "$wordToComplete*" });
        } )][string]$Template,      
      [Parameter(Mandatory = $false, Position = 2)][OutFormat]$Format='Table'      
     )

    begin {
        $errMsg='Unhandled exeption';
    }; # end begin

    process {
        try {            
            $errMsg='Failed to read templates';            
            switch ($PSCmdlet.ParameterSetName)
            {
                'Module'    {
                    $searchFilter=($script:TemplatePrefix + $ModuleName + '_*') ;
                }; # end Module
                'Template'  {
                    $searchFilter=($script:TemplatePrefix + $Template + '*');
                }; # end Template
                default     {
                    $searchFilter=($script:TemplatePrefix + '*');
                }; # end default
            }; # end switch
            writeLogOutput -LogString ('Reading templates with filter ' + $searchFilter);
            $tData=(Get-SecretInfo -Vault $script:vaultName -Name ($searchFilter) -ErrorAction Stop);
        } # end try
        catch {
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_; 
            return;
        }; # end catch
        
        if ($null -ne $tData) # verify if data was found
        {
            $tblFieldList=@(@('Name',[String]),
                            @('Description',[String]),                            
                            @('Default Cfg',[System.String]),
                            @('# of Cfg',[int]),
                            @('Parameter',[string]),                            
                            @('Type',[string]),
                            @('Mandatory',[System.Boolean]),
                            @('Def. Val',[String]),
                            @('Help Msg.',[String])                     
                        ); # end fieldlist
            $outTbl=createNewTable -TableName 'Output' -FieldList $tblFieldList
            foreach ($entry in $tData.Metadata)
            {
                [array]$tParamList=($entry.Keys).Where({! $_.StartsWith($script:cfgNamePrefix)}); # get list of parameters
                $tParamList=[System.Linq.Enumerable]::OrderBy([string[]]$tParamList,[Func[string,string]] {$args[0]}); # sort list of parameters
                #$tmpArr=($entry.(($tParamList[0])).split(($Script:sepVal)));
                $tmpArr=($entry.(($tParamList[0])).split(($Script:pPropSepVal)));
                if ([System.String]::IsNullOrEmpty($entry.$script:ConfigurationBackLink))
                {
                    $numOfCfg=0;
                } # end if
                else {
                    $numOfCfg=(($entry.$script:ConfigurationBackLink).split($Script:blSepVal)).count;
                }; # end if
                # add first row with name, description and first parameter
                if ([System.String]::IsNullOrEmpty(($entry.($script:MetadataDefaultCfg))))
                {
                    $defcfg=$null;
                } # end if
                else {
                    #$defcfg=(($entry.($script:MetadataDefaultCfg)).Split('_'))[4];
                    $defcfg=(($entry.($script:MetadataDefaultCfg)).Split($Script:AttribNameSep))[4];
                }; # end else
                $tblRow=@(
                    $entry.$script:ObjectName,
                    $entry.$script:MetadataDesc,
                    $defcfg,
                    $numOfCfg,
                    (($tParamList[0]).split(($script:paramPrefixSep)))[1], # split variable/parameter
                    $tmpArr[0],
                    [bool][int]$tmpArr[1],
                    $tmpArr[3],                                      
                    $tmpArr[2]                                      
                ); # end tblRow
                [void]$outTbl.rows.Add($tblRow);
                for ($i=1;$i -lt $tParamList.count;$i++) # add remaining parameters
                {
                    #$tmpArr=($entry.(($tParamList[$i])).split(($Script:sepVal))); # get data type and if param is mandatory
                    $tmpArr=($entry.(($tParamList[$i])).split(($Script:pPropSepVal))); # get data type and if param is mandatory
                    $tblRow=@(
                        $null,
                        $null,
                        $null,
                        $null,
                        (($tParamList[$i]).split(($script:paramPrefixSep)))[1], # split variable/parameter
                        $tmpArr[0],
                        [bool][int]$tmpArr[1],
                        $tmpArr[3],$tmpArr[2]
                    ); # end tblRow
                    [void]$outTbl.rows.Add($tblRow);
                }; # end for
            }; # end foreach 
            switch ($Format)
            {
                'Table' {
                    $outTbl | Format-Table #$fieldList;
                    break;
                }; # format tabel
                'List'  {
                    $outTbl | Format-List #$fieldList;
                    break;
                }; # format list
                'PassValue'  {
                    $outTbl;
                    break;
                }; # format list                                        
            }; # end switch                 
        } # end if
        else {
            writeLogOutput -LogString 'No templates found' -LogType Warning;
        }; # end else                     
    }; # end process

    end {

    }; # end END
    
}; # end funciton Get-ModuleTemplate

function Unregister-ModuleTemplate
{
<#
.SYNOPSIS 
Unregisters a template.
.DESCRIPTION
The command unregisters (deletes) a template. If the template has configurations assigned, the command will fail. A warning message and a list of configurations assigned will be returned. To unregister a template with configurations assigned use the Force switch. If the Force switch is used all assigned configurations will be deleted to.

.PARAMETER ModuleName
The parameter is mandatory. Data type string.
The name of the module. Because every template is registered for a particular module, the parameter makes sure, that only templates for that module are provided for unregistering.

.PARAMETER Template
The parameter is mandatory. Data type string.
The parameter expects the name of the Template to unregister. The template name is expected in the following format:
-	<Module name>_Ver:<Version> (example: OLX_Ver:0.0.0.0)
To simplify the election of the template tab-complete is supported. Simply enter one or more characters of the module and use tab-complete to select the template to unregister.

.PARAMETER Force
The parameter is optional. Data type switch.
If the parameter is used, the template will be unregistered even if configurations are assigned.

.EXAMPLE
Unregister-CMMModuleTemplate -ModuleName OLX -Template OLX_Ver:0.0.0.0
#>
[cmdletbinding(SupportsShouldProcess,ConfirmImpact='High')]    
param([Parameter(Mandatory = $false, Position = 0)][ArgumentCompleter( {
    param ( $CommandName,
        $ParameterName,
        $WordToComplete,
        $CommandAst,
        $FakeBoundParameters )  
        $cfgList=$__CMM_ModuleData.GetModuleList();
        $tmpList=@();
    foreach ($item in $tmpList)
    {
        if ($item.contains(' '))
        {
            $cfgList+="'"+$item+"'";
        } # end if
        else {
            $cfgList+=$item;
        }; # end else
    };
    $cfgList.Where({ $_ -like "$wordToComplete*" });
    } )][string]$ModuleName,
      [Parameter(Mandatory = $false, Position = 3)][switch]$Force
     )
   DynamicParam
    {                    
        if ($ModuleName) 
        {  
            [array]$itemList=$__CMM_ModuleData.GetCfgOrTemplateList($ModuleName,$true); # get list of configs
            for ($i=0;$i -lt $itemList.Count;$i++)
            {
                #$itemList[$i]=($ModuleName + '_' + $Script:VerStr + $itemList[$i]);        # reformat configs
                $itemList[$i]=($ModuleName + $Script:EntrySep + $Script:VerStr + $itemList[$i]);        # reformat configs
            };
            try {
                $dynParamList=@{
                ParamName='Template';
                Position=1;
                IsMandatory=$True;
                ValidateSet = $itemList;
            } # end dynParamList           
            newDynamicParam @dynParamList  
            } # end try
            catch {
            Write-Host 'Failed to build parameters' -ForegroundColor Red -BackgroundColor Black;
            }; # end catch   
        }; # end if               
    }; # end DynamicParam
#>
    begin {
        $errMsg='Unhandled exeption';
    }; # end begin

    process {
        try {
            $cfgString=($PSBoundParameters.Template)  
            $cfgFullName=$Script:TemplatePrefix+($cfgString.Replace($Script:VerStr,''));            
            $errMsg=('Failed to read the template ' + $cfgFullName);
            writeLogOutput -LogString ('Reading template ' + $cfgFullName);
            if ($tmp=(loadMetadata -MetaData ((Get-SecretInfo -Vault $script:vaultName -Name $cfgFullName -ErrorAction Stop).Metadata)))
            {
                if ((! (isTemlateRegistered -TemplateName $cfgFullName)) -and ($Force.IsPresent -eq $false))
                {
                    writeLogOutput -LogString ('To perform the action, use the Force parameter.') -LogType Warning;
                    return;
                }; # end if
                if ((!([System.String]::IsNullOrEmpty($Tmp.$script:ConfigurationBackLink))) -and ($Force.IsPresent -eq $false))
                {
                    writeLogOutput -LogString ('For the template ' + $cfgString + ' are one or more configuration available. To remove the template delete the configurations or use the parameter Force.') -LogType Warning;
                    writeLogOutput -LogString ('Configurations for template ' + $cfgString +':') -ShowInfo;
                    $cfgList=($Tmp.$script:ConfigurationBackLink.split($Script:blSepVal));
                    for ($i=0; $i -lt $cfgList.count;$i++)
                    {
                        #$tItem=$cfgList[$i].split('_');
                        $tItem=$cfgList[$i].split($Script:AttribNameSep);
                        #$tItem[4]+'_'+$Script:VerStr+$tItem[3]
                        $tItem[4]+$Script:AttribNameSep+$Script:VerStr+$tItem[3]
                    };
                    return;
                }; # end if
                
                if($PSCmdlet.ShouldProcess('Performing on ','Module ' + ($script:moduleName + ' template ' + $cfgString),'Unregistering template (deleting configuration if exist)')) # ask user for confirmation
                {
                    $delList=($Tmp.$script:ConfigurationBackLink)  
                    if (! ([System.String]::IsNullOrEmpty(($Tmp.$script:ConfigurationBackLink))))
                    {             
                        $delList=$delList.split($Script:blSepVal);
                        for ($i=0; $i -lt $delList.count;$i++)
                        {
                            $errMsg=('Failed to remove the configuration ' + ($delList[$i]) + ' Please remove the configuration manually.');
                            writeLogOutput -LogString ('Removing configuration ' + ($delList[$i]));
                            try {
                                Remove-Secret -Vault $script:vaultName -Name $delList[$i] -ErrorAction Stop;
                            } # end try
                            catch {
                                writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
                            }; # end try                        
                        }; # end for
                    }; # end if
                    $errMsg=('Failed to remove template ' + $cfgString);
                    writeLogOutput -LogString ('Removing configuration ' + $cfgFullName);
                    Remove-Secret -Vault $script:vaultName -Name $cfgFullName -ErrorAction Stop; 
                    writeLogOutput -LogString ('Updating module configruation');
                    $moduleCfg=loadMetadata -MetaData (Get-SecretInfo -Vault $Script:vaultName -Name $script:ModuleCfgName).Metadata;                    
                    $tmp=($moduleCfg.RegisteredTemplateList).Split(($Script:sepVal)); 
                    $tmp=[array]([System.Linq.Enumerable]::Except([string[]]$tmp,[string[]]$cfgFullName));           
                    if ([System.String]::IsNullOrEmpty($tmp))
                    {
                        $moduleCfg.RegisteredTemplateList=[string]'';
                        $moduleCfg.RegisteredModuleList=[string]'';
                    } # end if
                    else {
                        $moduleCfg.RegisteredTemplateList = [string]($tmp -join ($Script:sepVal));
                        $tmpArr=[System.Collections.ArrayList]::new();
                        $tmpArr.AddRange(@($moduleCfg.RegisteredModuleList.split($Script:sepVal)));
                        #$tmpArr.Remove((($cfgFullName.Replace($script:TemplatePrefix,''))).split('_')[0]);
                        $tmpArr.Remove((($cfgFullName.Replace($script:TemplatePrefix,''))).split($Script:EntrySep)[0]);
                        $moduleCfg.RegisteredModuleList = $tmpArr -join ($Script:sepVal);
                    }; # end else            
                    $errMsg=('Failed to write module configuration to ' + $script:ModuleCfgName)
                    writeToModuleConfig -CfgData $moduleCfg;
                }; # end if
            } # end if
            else {
                writeLogOutput -LogString ('Configuration ' + $cfgFullName + ' not found.') -LogType Error;
            }; # end else
        } # end try
        catch {
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        }; # end catch
    }; # end process

    end {

    }; # end END
}; # end function Unregister-ModuleTemplate

#endregion template


#region config/profile
function New-Configuration
{
<#
.SYNOPSIS 
Creates a new configuration.
.DESCRIPTION
The command creates a new configuration for a template. For a registered template, multiple configurations can be created. A configuration provides the variables and their values to be used in a PowerShell module or script. 
Example: If a PowerShell module can manage multiple O365 tenants, every single tenant requires different credential. Every configuration stores the credential and URL, and other variables when needed, for a particular tenant. A command in a PowerShell module can query a particular configuration to access a particular tenant.
Most of the parameters will be created at runtime of the command. What parameter will be provided and what data type a particular parameter has, is configured in the template selected in the parameter Template.
.PARAMETER Template
The parameter is mandatory. Data type string.
The parameter expects the name of a registered template. The parameter supports tab-complete (ArgumentCompleter).

.PARAMETER Name
The parameter is mandatory. Data type string.
The parameter expects the name of the configuration. Only the following characters are allowed:
-	a-z
-	A-Z
-	0-9

.PARAMETER Description
The parameter is optional. Data type string.
The command expects a description for the configuration.

.PARAMETER AllowPushNotification
The parameter is optional. Data type switch.
If the parameter is used, push notification for a configuration change will be configured. If a configuration is updated, the CMM command will try to find out if the module, for which the template of the configuration was registered for, is active. If the module is active, the module will be notified. 
Disclaimer: PushNotification is an experimental feature and works only in the same PowerShell session. If PushNotification should be used, the module/script must export a variable with the name $__<ModuleName>_* (the * is a placeholder for additional characters). This variable must provide a method with the name UpdateConfig. More information to this topic can be found in the folder Examples under the root folder of the module.

.EXAMPLE
New-CMMConfiguration -Template OLX_Ver:0.0.0.0 -Name Cust1 -Description 'OLX cfg for Cust1' -OLXCred tom@cust1.com -OLXSendReportsTo fred@cust1.com
The names for the parameters
-	OLXCred
-	OLXSendReportsTo
where defined in the template.
#>
[cmdletbinding()]    
param([Parameter(Mandatory = $true, Position = 0)]
      [ArgumentCompleter( {  
        param ( $CommandName,
        $ParameterName,
        $WordToComplete,
        $CommandAst,
        $FakeBoundParameters )           
        $mlist=($__CMM_ModuleData.GetModuleAndVerList());
        $mlist.Where({ $_ -like "$wordToComplete*" });              
      } )][string]$Template,
      [Parameter(Mandatory = $true, Position = 1)][ValidateScript({$_ -match $RxCfgName})][string]$Name,
      [Parameter(Mandatory = $false, Position = 2)][string]$Description='',
      [Parameter(Mandatory = $false, Position = 3)][switch]$AllowPushNotification
     )

DynamicParam
{
    			
    if ($Template -and $Name) 
    {          
        try {
            $mData=loadMetadata -MetaData ((Get-SecretInfo -Vault ($__CMM_ModuleData.vaultName) -Name ($script:TemplatePrefix+$Template.Replace(($__CMM_ModuleData.verStr),''))).Metadata);
            
            $pList=($mData.Keys | Where-Object {! ($_.StartsWith($__CMM_ModuleData.cfgNamePrefix))});                  
            $dynParamList=@{};
            $defValList=@{};
            $dataTypeList=@{};            
            foreach ($p in $pList)
            {
                $dynParamList.Add($p,$mData.$p);
                $tmpDefVal=($mData.$p).split($Script:pPropSepVal)[3];
                #if (!($mData.$p.EndsWith(($Script:sepVal))))
                #if (!([system.string]::IsNullOrEmpty(($mData.$p).Split(($Script:pPropSepVal))[3])))
                if (!([system.string]::IsNullOrEmpty($tmpDefVal)))
                {
                    #$defValList.Add($p.split(($script:paramPrefixSep))[1],(($mData.$p).split(($Script:sepVal)))[3])
                    #$defValList.Add($p.split(($script:paramPrefixSep))[1],(($mData.$p).split(($Script:pPropSepVal)))[3])
                    $defValList.Add($p.split(($script:paramPrefixSep))[1],$tmpDefVal);
                    #$dataTypeList.Add($p.split(($script:paramPrefixSep))[1],(($mData.$p).split(($Script:sepVal)))[0]);
                    $dataTypeList.Add($p.split(($script:paramPrefixSep))[1],(($mData.$p).split(($Script:pPropSepVal)))[0]);
                }; # end if                
            }; # end foreach
            newDynamicParamList -ParamHash $dynParamList -FirstParamPosition 4;
            foreach ($param in $defValList.Keys) # check for default values
            {
                if (! ($dataTypeList.$param -eq 'Boolean'))
                {
                    $PSBoundParameters[$param]=($defValList.$param -as $dataTypeList.$param); # assign default values
                } # end if
                else {
                    $PSBoundParameters[$param]=[System.Convert]::ToBoolean($defValList.$param); # assign default bool value
                }; # end if
               
            }; # end foreach      
        } # end try
        catch {
          Write-Host 'Failed to build parameters' -ForegroundColor Red -BackgroundColor Black;
        }; # end catch   
    }; # end if               
}; # end DynamicParam

    begin {
        $errMsg='Unhandled exeption';
        
    }; # end begin

    process {
        #$cfgName=($script:ConfigPrefix+($Template.Replace($script:VerStr,''))+'_'+$Name); 
        $cfgName=($script:ConfigPrefix+($Template.Replace($script:VerStr,''))+$Script:EntrySep+$Name);         
        try {            
            $errMsg='Failed to verify if configuration exist';
            writeLogOutput -LogString ('Verifing if configuration ' + $cfgName + ' exist');
            if (Get-SecretInfo -Vault $script:vaultName -Name $cfgName)
            {
                writeLogOutput -LogString ('A configuration with the name ' + $cfgName + ' already exist') -LogType Error;
                return;
            }; # end if
            writeLogOutput -LogString 'Creating configuraton to store'
            $paramsToExclude=@('Template','Name','Description','AllowPushNotification');
            $boundParams=$PSBoundParameters.Keys;
            $paramsToExclude+=[System.Management.Automation.PSCmdlet]::CommonParameters;
            $paramsToExclude+=[System.Management.Automation.PSCmdlet]::OptionalCommonParameters;
            $paramList=[System.Linq.Enumerable]::Except([string[]]$boundParams,[string[]]$paramsToExclude);
            $templateName=($script:TemplatePrefix+$Template.Replace(($__CMM_ModuleData.verStr),''));
            $d=([System.DateTime]::Now);
            $metaData=@{
                ($script:ObjectName)=($Name);
                $script:ConfigDesc=$Description;
                ($script:BaseTemplate)=$templateName; # add name of template where the cfg is based on
                ($Script:CreateDate)=$d; # date object created
                ($Script:ChangeDate)=$d; # date object changed
                ($script:CfgAllowPush)=[int]($AllowPushNotification.IsPresent);  
                ($Script:cfgVerString)=$script:tCfgVer;                    
            }; # end metaData
            $cfgParamList=[System.Collections.ArrayList]::new(); # init var
            
            foreach ($param in $paramList)
            {
                if (($PSBoundParameters.$param).GetType().Name -ne 'PSCredential')
                {
                    if (($PSBoundParameters.$param.GetType().Name) -ne 'Boolean')
                    {
                        $metaData.Add($param,($PSBoundParameters.$param));
                    } # end if
                    else {
                        $metaData.Add($param,[Int32][bool]($PSBoundParameters.$param));
                    }; # end else                    
                } # end if
                else {
                    $cred=$PSBoundParameters.$param;
                    $metaData.Add(($script:CredVarName),$param);
                }; # end else  
                [void]$cfgParamList.Add($param);              
            }; # end foreach
            $metaData.Add($script:cfgNamePrefix +'_PList',($cfgParamList -join $Script:pLstSepVal));
            $errMsg=('Failed to create the configuraton ' + $Name + ' for module ' + $Template);
            writeLogOutput -LogString ('Creating configuraton ' + $Name + ' for module ' + $Template);
            Set-Secret -Name $cfgName -Vault $script:vaultName -Secret $cred -NoClobber -Metadata $metaData;
            addBackLinkToTemplate -TemplateName $templateName -ConfigurationName $cfgName;
        } # end try
        catch {
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        }; # end catch
    }; # end process
    
    end {

    }; # end END
}; # end function  New-Configuration

function Remove-Configuration
{
<#
.SYNOPSIS 
Removes a configuration.
.DESCRIPTION
The command removes a particular configuration. If the configuration is configured as a default configuration for the template, removing the configuration will fail. A warning message is displayed with the hint, that the parameter Force must be used to remove the configuration.

.PARAMETER Template
The parameter is mandatory. Data type string.
The name of the template to which the configuration to remove is assigned to. The parameter supports tab-complete (ArgumentCompleter).

.PARAMETER Configuration
The parameter is madatory. Data type string.
The name of the configuration to remove. The parameter supports tab-complete (list of possible configurations will be provided). 

.PARAMETER Force
The parameter is optional. Data type switch.
If the configuration is configured as default configuration for the template, to remove the configuration the parameter Force can be used. Alternatively a different configuration can be configured as default configuration for the template.
.EXAMPLE
Remove-CMMConfiguration -Template OLX_0.0.0.0 -Configuration NoLimit
The configuration NoLimit will be removed.
#>

[cmdletbinding(SupportsShouldProcess,ConfirmImpact='High')]    
param([Parameter(Mandatory = $false, Position = 0)][ArgumentCompleter( {
    param ( $CommandName,
        $ParameterName,
        $WordToComplete,
        $CommandAst,
        $FakeBoundParameters ) 
        $cfgList=$__CMM_ModuleData.GetCfgOrTemplateList($null,$true)
        $tmpList=@();
    foreach ($item in $tmpList)
    {
        if ($item.contains(' '))
        {
            $cfgList+="'"+$item+"'";
        } # end if
        else {
            $cfgList+=$item;
        }; # end else
    };
    $cfgList.Where({ $_ -like "$wordToComplete*" });
    } )][string]$Template,
    [Parameter(Mandatory = $false, Position = 1)][switch]$Force=$false    
     )
    DynamicParam
    {                    
        if ($Template) 
        {  
            [array]$itemList=$__CMM_ModuleData.GetCfgOrTemplateList($Template,$false); # get list of configs            
            try {
                $dynParamList=@{
                ParamName='Configuration';
                Position=2;
                IsMandatory=$True;
                ValidateSet = $itemList;
            } # end dynParamList           
            newDynamicParam @dynParamList  
            } # end try
            catch {
            Write-Host 'Failed to build parameters' -ForegroundColor Red -BackgroundColor Black;
            }; # end catch   
        }; # end if               
    }; # end DynamicParam

    begin {
        $errMsg='Unhandled exeption';
    }; # end begin

    process {
        try {
            $cfgString=($PSBoundParameters.Configuration)
            $errMsg=('Failed to read configuration ' + $cfgString + ' for template ' + $Template);           
            #$tmp=$cfgString.split('_');
            $tmp=$cfgString.split($Script:EntrySep);
            #$cfgFullName=$Script:ConfigPrefix+($Template)+'_'+$cfgString;            
            $cfgFullName=$Script:ConfigPrefix+($Template)+$Script:EntrySep+$cfgString;            
            $errMsg=('Failed to read the configuration ' + $cfgFullName);
            writeLogOutput -LogString ('Reading configuration ' + $cfgFullName);
            if ($tmp=loadMetadata -MetaData ((Get-SecretInfo -Vault $script:vaultName -Name $cfgFullName -ErrorAction Stop).Metadata))
            {
                if ((!($isDefaultCfg=isConfigDefault -ConfigName $cfgFullName -ForceSwitch ($Force.IsPresent))) -or ($Force.IsPresent))
                {
                    if($PSCmdlet.ShouldProcess('Performing on ',($script:moduleName + ' configuration ' + $cfgString),'DELETION')) # ask user for confirmation
                    {
                        $baseTemplate=$tmp.($script:BaseTemplate);
                        $errMsg=('Failed to read the template ' + $baseTemplate);
                        writeLogOutput -LogString ('Reading template ' + $baseTemplate);
                        $baseMetaData=loadMetadata -MetaData ((Get-SecretInfo -Vault $script:vaultName -Name $baseTemplate -ErrorAction Stop).Metadata); # query data from template
                        $backLink=$baseMetaData.($script:ConfigurationBackLink);
                        $backLink=$backLink.split($Script:blSepVal);
                        $errMsg='Failed to format backlink data'
                        writeLogOutput -LogString ('Formating backlink data');
                        [array]$newBackLink=[System.Linq.Enumerable]::Except([string[]]$backLink,[string[]]$cfgFullName);                    
                        $baseMetaData.($script:ConfigurationBackLink)=($newBackLink -join $Script:blSepVal);
                        if($isDefaultCfg)
                        {
                            $baseMetaData.($Script:MetadataDefaultCfg)=''; # set def cfg in template to blank
                        }; # end if
                        $errMsg=('Failed to remove the configuraton ' + $cfgFullName);
                        writeLogOutput -LogString ('Remoing configuration configuration ' + $cfgFullName);
                        Remove-Secret -Vault $script:vaultName -Name $cfgFullName -ErrorAction Stop;
                        $errMsg=('Failed to configure backlink (' + $baseMetaData.($script:ConfigurationBackLink) +')');
                        writeLogOutput -LogString ('Setting backlink data (' + $baseMetaData.($script:ConfigurationBackLink) +')');
                        Set-SecretInfo -Vault $script:vaultName -Name $baseTemplate -Metadata $baseMetaData -ErrorAction Stop;
                    }; # end if
                }; # end if
            } # end if
            else {
                writeLogOutput -LogString ('Configuration ' + $cfgFullName + ' not found.') -LogType Error;
            }; # end else
        } # end try
        catch {
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        }; # end catch
    }; # end process

    end {

    }; # end END
}; # end function Remove-Configuration


function Set-Configuration
{
<#
.SYNOPSIS 
Reconfigures an existing configuration.
.DESCRIPTION
The command allows to reconfigure multiple attributes and parameters of a configuration.
Many of the parameters are created at runtime of the command and cannot be described. This dynamically created parameters are based (name, data type, etc.) on the configuration defined in the template.

.PARAMETER Template
The parameter is mandatory. Data type string.
The name of the template where the configuration to reconfigure is assigned to. The parameter supports tab-complete (ArgumentCompleter).

.PARAMETER Configuration
The parameter is mandatory. Data type string.
The name of the configuration to reconfigure. The parameter supports tab-complete (ArgumentCompleter).

.PARAMETER SetAsDefaultConfiguration
The parameter is optional. Data type switch.
If the parameter is used, the configuration will be configured as default configuration it the template.
The parameter cannot be used together with the parameter RemoveAsDefaultConfiguration.

.PARAMETER RemoveAsDefaultConfiguration
The parameter is optional. Data type switch.
If the parameter is used, the configuration will be removed as default configuration from the template.
The parameter cannot be used together with the parameter SetAsDefaultConfiguration.

.PARAMETER Description
The parameter is optional. Data type string.
Configures the description of the configuration.
.PARAMETER AllowPushNotification
The parameter is optional. Data type Boolean.
The parameter configures if PushNotification, for configuration updates, is allowed.

.EXAMPLE
Set-CMMConfiguration -Template OLX_0.0.0.0 -Configuration Cust1Lab -SetAsDefaultConfiguration
Configures the configuration Cust1Lab1 as default configuration for the template OLX_0.0.0.0.
EXAMPLE
Set-CMMConfiguration -Template OLX_0.0.0.0 -Configuration Cust1Lab -OLXHost host.name
Sets the value for OLXHost in the configuration Cust1Lab1 to host.name.
#>
[cmdletbinding(DefaultParametersetName='__AllParameter')]    
param([Parameter(Mandatory = $true, Position = 0)][ArgumentCompleter( {
    param ( $CommandName,
        $ParameterName,
        $WordToComplete,
        $CommandAst,
        $FakeBoundParameters )  
        $cfgList=$__CMM_ModuleData.GetCfgOrTemplateList($null,$true);
        $tmpList=@();
    foreach ($item in $tmpList)
    {
        if ($item.contains(' '))
        {
            $cfgList+="'"+$item+"'";
        } # end if
        else {
            $cfgList+=$item;
        }; # end else
    };
    $cfgList.Where({ $_ -like "$wordToComplete*" });
    } )][string]$Template,

    [Parameter(Mandatory = $false, Position = 1)][ArgumentCompleter( {        
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters ) 
            $tmpList=@();
         if ($fakeBoundParameters.ContainsKey('Template'))
        {            
            @($tmpList=$__CMM_ModuleData.GetConfigForTemplate($fakeBoundParameters.Template));
        } # end if
        else
        {
            'Error: Missing parameter template'
        }; # end else
        foreach ($item in $tmpList)
        {            
            if ($item.contains(' '))
            {
                $cfgList+="'"+$item+"'";
            } # end if
            else {
                $cfgList+=$item;
            }; # end else
        }; # end foreach
        $tmpList.Where({ $_ -like "$wordToComplete*" }); 
        } )][string]$Configuration,
        [Parameter(ParametersetName='setDef',Mandatory = $false, Position = 2)][switch]$SetAsDefaultConfiguration,
        [Parameter(ParametersetName='clearDef',Mandatory = $false, Position = 2)][switch]$RemoveAsDefaultConfiguration,
        [Parameter(Mandatory = $false, Position = 3)][string]$Description,
        [Parameter(Mandatory = $false, Position = 4)][bool]$AllowPushNotification
     )
    DynamicParam
    {                    
        if ($Configuration -and $Template)        
        {  
            try {
                $mData=loadMetadata -MetaData ((Get-SecretInfo -Vault ($__CMM_ModuleData.vaultName) -Name ($script:TemplatePrefix+$Template)).Metadata);                            
                $pList=($mData.Keys | Where-Object {! ($_.StartsWith($__CMM_ModuleData.cfgNamePrefix))});                
                $dynParamList=@{};
                foreach ($p in $pList)
                {
                    $dynParamList.Add($p,$mData.$p);
                }; # end foreach
                
                newDynamicParamList -ParamHash $dynParamList -FirstParamPosition 5 -SetMandatoryToFalse;            
            } # end try
            catch {
              Write-Host 'Failed to build parameters' -ForegroundColor Red -BackgroundColor Black;
            }; # end catch               
        }; # end if               
    }; # end DynamicParam

    begin {
        $errMsg='Unhandled exeption';
    }; # end begin

    process {
        try {
            $cfgString=($script:ConfigPrefix + $PSBoundParameters.Template + $Script:EntrySep + $PSBoundParameters.Configuration)
            $errMsg=('Failed to read the configuration ' + $cfgString);
            writeLogOutput -LogString ('Reading configuration ' + $cfgString);
            if ($cfgData=Get-SecretInfo -Vault $script:vaultName -Name $cfgString -ErrorAction Stop)
            {
                $excludedParams=[System.Collections.ArrayList]::new();
                $excludedParams.AddRange(@('Template','Configuration','SetAsDefaultConfiguration','RemoveAsDefaultConfiguration','Description','AllowPushNotification'));
                $pList=[System.Collections.ArrayList]::new();
                $pList.AddRange([array]([System.Linq.Enumerable]::Except([string[]]$PsBoundParameters.Keys,[string[]]$excludedParams)));
                $isCred=$false;                
                $i=0;
                while ($i -lt $plist.count -and ($isCred -eq $false))
                {
                    $isCred = (($PsBoundParameters.($pList[$i])).getType().name -eq 'PSCredential')
                    $i++;
                }; # end while
                
                $pListHash=@{
                    Vault=$script:vaultName;
                    Name=$cfgString;
                    ErrorAction='Stop';
                }; # end if
                if ($isCred)
                {
                    $pListHash.Add('Secret',($PsBoundParameters.($pList[($i-1)])));   
                    [void]$pList.Remove($pList[($i-1)]);           
                }; # end if                
                $mdata=(loadMetadata -MetaData $cfgData.Metadata);                
                switch ($PsBoundParameters) # set attributes
                {
                    {$_.ContainsKey('Description')} {
                        $mData.($script:ConfigDesc)=$Description;
                    }; # end Description
                    {$_.ContainsKey('AllowPushNotification')} {
                        $mData.($script:CfgAllowPush)=[int]$AllowPushNotification;
                    }; # end AllowPushNotification
                }; # end switch
                for ($j=0; $j -lt $pList.Count;$j++)
                {
                    $mData.($pList[$j])=$PsBoundParameters.($pList[$j]);
                }; # end for  
                writeLogOutput -LogString ('Saving changes to ' + ($cfgString.Replace($script:ConfigPrefix,'')));
                $errMsg=('Failed to save the changes to ' + $cfgString);  
                
                $mData.($Script:ChangeDate)=([System.DateTime]::Now); # set change date
                if ($isCred)
                {
                    Set-Secret @pListHash -Metadata $mData;
                } # end if
                else {
                    Set-SecretInfo @pListHash -Metadata $mData;
                }; # end else
                if ($mData.$script:CfgAllowPush -and ($pList.Count -gt 0 -or $isCred))
                {
                    $tmpArr=$cfgString.split($Script:EntrySep);
                    $mList = Get-Variable -Name ($script:ExportVarPrfx + $tmpArr[2]+'*'); # search for vars with method UpdateConfig
                    foreach ($moduleVar in $mList)
                    {
                        try {
                            (Get-Variable -Name $moduleVar.Name -ValueOnly).UpdateConfig($tmpArr[4],$tmpArr[3]);
                        } # end try
                        catch {
                            writeLogOutput -LogString ('Method UpdateConfig for ' + $moduleVar.Name + ' not found') -LogType Warning;
                        }; # end catch
                    }; # end foreach
                }; # end if
                if ($SetAsDefaultConfiguration.IsPresent -or $RemoveAsDefaultConfiguration.IsPresent)
                {
                    writeLogOutput -LogString ('Setting/removing configuration ' + $Configuration + ' as default.')
                    $errMsg=('Failed to set/remove configuration ' + $Configuration + ' as default.'); 
                    setCfgAsDefault -Template ($mData.($Script:BaseTemplate)) -Configuration $cfgString -SetAsDefault:($SetAsDefaultConfiguration.IsPresent);
                }; # end if
            } # end if
            else {
                writeLogOutput -LogString ('Configuration ' + $cfgString + ' not found.') -LogType Error;
            }; # end else
        } # end try
        catch {
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_;
        }; # end catch
    }; # end process

    end {

    }; # end END
}; # end function Set-Configuration

function Get-Configuration
{
<#
.SYNOPSIS 
List configurations.
.DESCRIPTION
The command lists configuration. If neither of the parameters
-	ModuleName
-	Template
-	Configuration
is used, all configurations are listed.

.PARAMETER ModuleName
The parameter is optional. Data type string.
If the parameter is used, only configurations for a particular module are listed (configurations for all templates for a particular module). The parameter supports tab-complete (ArgumentCompleter).

.PARAMETER Template
The parameter is optional. Data type string.
If the parameter is used, only configurations for a particular template are. The parameter supports tab-complete (ArgumentCompleter).

.PARAMETER Configuration
The parameter is optional. Data type string.
The data for a particular configuration is listed. The parameter supports tab-complete (ArgumentCompleter).

.PARAMETER Format
The parameter is optional. Data type enum OutFormat
If the parameter is omitted the data is returned in table format. The parameter accepts the following formats:
-	Table (default)
-	List
-	PassValue (returns unformatted data)

.EXAMPLE
Get-CMMConfiguration -Template OLX_0.0.0.0
List all configurations for the template OLX_0.0.0.0
.EXAMPLE
Get-CMMConfiguration -ModuleName OLX -Configuration Cust1Lab_Ver:0.0.0.0
Lists the configuration Cust1Lab for the module OLX.
#>
[cmdletbinding(DefaultParametersetName='__AllParameter')]    
param([Parameter(ParameterSetName='Module',Mandatory = $false, Position = 0)][ArgumentCompleter( {
    param ( $CommandName,
        $ParameterName,
        $WordToComplete,
        $CommandAst,
        $FakeBoundParameters )  
        $cfgList=$__CMM_ModuleData.GetModuleList();
        $tmpList=@();
    if ($FakeBoundParameters.Configuration)
    {
        ('''Parameter ModuleName not allowed''')
    } # end if
    else {        
        foreach ($item in $cfgList)
        {
            if ($item.contains(' '))
            {
                $tmpList+="'"+$item+"'";
            } # end if
            else {
                $tmpList+=$item;
            }; # end else
        };
        $tmpList.Where({ $_ -like "$wordToComplete*" });
    }; #end else
    } )][string]$ModuleName,

      [Parameter(ParameterSetName='Cfg',Mandatory = $false, Position = 0)][ArgumentCompleter( {
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
            $cfgList=$__CMM_ModuleData.GetCfgOrTemplateList($null,$true);
            $tmpList=@();
        foreach ($item in $cfgList)
        {
            if ($item.contains(' '))
            {
                $tmp+="'"+$item+"'";
            } # end if
            else {
                $tmpList+=$item;
            }; # end else
        };        
        $tmpList.Where({ $_ -like "$wordToComplete*" });    
        } )][string]$Template,
        [Parameter(ParameterSetName='Module',Mandatory = $false, Position = 0)][ArgumentCompleter( {
          param ( $CommandName,
              $ParameterName,
              $WordToComplete,
              $CommandAst,
              $FakeBoundParameters )  
              $tmpList=$__CMM_ModuleData.GetCfgOrTemplateList($FakeBoundParameters.Modulename,$false);              
          if ($FakeBoundParameters.ModuleName)
          {
             for ($i=0;$i -lt $tmpList.Count;$i++)
             {
               $tVar=$tmpList[$i].split($__CMM_ModuleData.entrySep);
               $tmpList[$i]=($tVar[1]+$__CMM_ModuleData.entrySep+$__CMM_ModuleData.verStr+$tVar[0]);
             }; # end for
          } # end if
          else {
            for ($i=0;$i -lt $tmpList.Count;$i++)
            {
               $tVar=$tmpList[$i].split($__CMM_ModuleData.entrySep);
               $tmpList[$i]=($tVar[0]+$__CMM_ModuleData.entrySep+$tVar[2]+$__CMM_ModuleData.entrySep+$__CMM_ModuleData.verStr+$tVar[1]);
            }; # end for
          }; # end else
          for ($i=0;$i -lt $tmpList.Count;$i++)
          {
            if ($tmpList[$i].contains(' '))
            {
                $tmpList[$i]="'"+$tmpList[$i]+"'";
            } # end if
            else {
                $tmpList[$i]=$tmpList[$i];
            }; # end else
          }; # end for          
          $tmpList.Where({ $_ -like "$wordToComplete*" });
          } )][string]$Configuration,      
      [Parameter(Mandatory = $false, Position = 2)][OutFormat]$Format='Table'      
     )

    begin {
        $errMsg='Unhandled exeption';
    }; # end begin

    process {
        try {            
            $errMsg='Failed to read templates';            
            switch ($PSCmdlet.ParameterSetName)
            {
                'Module'    {
                    # calculate string to list configurations
                    if ($PSBoundParameters.ContainsKey('Configuration'))
                    {
                        if ($PSBoundParameters.ContainsKey('ModuleName'))
                        {
                            $tStr=$Configuration.Split($Script:EntrySep);
                            $nameStr=($script:ConfigPrefix + $ModuleName + $tStr[1].Replace($script:VerStr,$Script:EntrySep) + $Script:EntrySep + $tStr[0]);
                        } # end if
                        else {
                            $tStr=$Configuration.Split($Script:EntrySep);
                            $nameStr=$script:ConfigPrefix+$tStr[0]+$Script:EntrySep+$tStr[2].Replace($script:VerStr,'')+$Script:EntrySep+$tStr[1];
                        }; # end else
                    } # end if
                    else {
                        $nameStr=$script:ConfigPrefix + $ModuleName + '*'
                    }; # end if
                    $tData=(Get-SecretInfo -Vault $script:vaultName -Name ($nameStr) -ErrorAction Stop);
                    break;
                }; # end Module
                'Cfg'  {
                    $tData=(Get-SecretInfo -Vault $script:vaultName -Name ($script:ConfigPrefix + $Template + '*') -ErrorAction Stop);
                    break;
                }; # end Template
                default     {
                    $tData=(Get-SecretInfo -Vault $script:vaultName -Name ($script:ConfigPrefix + '*') -ErrorAction Stop);
                }; # end default
            }; # end switch
        } # end try
        catch {
            writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_; 
            return;
        }; # end catch
        
        if ($null -ne $tData) # verify if data was found
        {
            $tblFieldList=@(@('Name',[String]),
                            @('Min. Ver.',[String]),
                            @('Description',[String]),                            
                            @('PushNotify',[String]),
                            @('Parameter',[string]),
                            @('Value',[string]));
            $outTbl=createNewTable -TableName 'Output' -FieldList $tblFieldList
            foreach ($entry in $tData)
            {
                [array]$tParamList=($entry.Metadata.Keys).Where({! $_.StartsWith($script:cfgNamePrefix)}); # get list of parameters
                $tParamList=[System.Linq.Enumerable]::OrderBy([string[]]$tParamList,[Func[string,string]] {$args[0]}); # sort list of parameters
                $errMsg=('Failed to read the template ' + ($entry.Metadata.($script:BaseTemplate)));                
                $baseTemplateMetadata=(Get-SecretInfo -Vault $script:vaultName -Name $entry.Metadata.($script:BaseTemplate) -ErrorAction Stop).Metadata; # get template
                $templateKeyList=getListOfCfgParameters -EntryList ($baseTemplateMetadata.Keys);
                try {
                    $cred=Get-Secret -Vault $script:vaultName -Name $entry.name;
                } # end try
                catch {
                    $errMsg=('Failed to read credential for config ' + $entry.Metadata.$script:ObjectName);                
                    writeLogError -ErrorMessage $errMsg -PSErrMessage ($_.Exception.Message) -PSErrStack $_; 
                }; # end catch
                $tblRow=@(    # fill table row
                    $entry.Metadata.$script:ObjectName,
                    $entry.Name.split($Script:EntrySep)[3],
                    $entry.Metadata.$script:ConfigDesc,
                    [boolean][int]($entry.Metadata.$script:CfgAllowPush)
                    $entry.Metadata.$script:CredVarName,
                    $cred.UserName
                ); # end tblRow
                [void]$outTbl.rows.Add($tblRow); # add data
                for ($i=0;$i -lt $tParamList.count;$i++) # add remaining parameters
                {
                    $keyVal=$templateKeyList.Where({$_ -like ('*'+$script:paramPrefixSep)+$tParamList[$i]});
                    $dType=($baseTemplateMetadata.[string]$keyVal[0]).Split(($Script:sepVal))[0];
                    [void]$outTbl.rows.Add($null,$null,$null,$null,($tParamList[$i]),($entry.Metadata.($tParamList[$i]) -As [System.Type]$dType));
                }; # end for
            }; # end foreach 
            switch ($Format)
            {
                'Table' {
                    $outTbl | Format-Table #$fieldList;
                    break;
                }; # format tabel
                'List'  {
                    $outTbl | Format-List #$fieldList;
                    break;
                }; # format list
                'PassValue'  {
                    $outTbl;
                    break;
                }; # format list                                        
            }; # end switch                      
        } # end if
        else {
            writeLogOutput -LogString 'No configuration found' -LogType Warning;
        }; # end else
        
        
    }; # end process

    end {

    }; # end END
    
}; # end funciton Get-Configuration

#endregion config/profile