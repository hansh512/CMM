################################################################################
# Code Written by Hans Halbmayr
# Created On: 13.06.2013
# Last change on: 22.06.2019
#
# Module: CMM
#
# Version 0.90
#
# Purpos: Function library for several modules
################################################################################  

function newDynamicParam
{	      
Param([Parameter(Mandatory = $true, Position = 0)][string]$ParamName,
      [Parameter(Mandatory = $false, Position = 1)][int]$Position=0,
      [Parameter(Mandatory = $false, Position = 3)][string]$ParameterSetName,
      [Parameter(Mandatory = $false, Position = 4)][switch]$IsMandatory=$false,
      [Parameter(Mandatory = $false, Position = 5)][switch]$ValueFromPipeline=$false,
      [Parameter(Mandatory = $false, Position = 6)]$ParmType=([String]),
      [Parameter(Mandatory = $false, Position = 7)][Array]$ValidateSet
     )

    $dynPList = @()   # array for properties of the dynamic parameters
    $dynParamSet = @{
        Position=$Position  
        Mandatory=$IsMandatory
        ParamName=$ParamName        
        ValueFromPipeline=$ValueFromPipeline
        ParamType=$ParmType
    } # end dynParamSet

    if ($PSBoundParameters.ContainsKey('ValidateSet'))
    {
        $dynParamSet.Add('OptionList',$ValidateSet);
    } # end if
    if ($PSBoundParameters.ContainsKey('ParameterSetName'))
    {
        $dynParamSet.Add('ParameterSetName',$ParameterSetName);
    } # end if
    
    $dynPList += (addDynamicParameterProperties @dynParamSet);
    newDynamicParameters -ParamList $dynPList
		
} # end funciton newGetADDomainsFromTenantParams

function setTemplateDynamicParamList
{	      
Param([Parameter(Mandatory = $true, Position = 0)][hashtable]$ParamHash,
      [Parameter(Mandatory = $false, Position = 1)][int]$FirstParamPosition=0  
     )

    $dynPList = [System.Collections.ArrayList]::new();   # array for properties of the dynamic parameters
    [array]$pKeys=[System.Linq.Enumerable]::OrderBy([string[]]$ParamHash.Keys, [Func[string,string]] { $args[0] });
    $dynParamSet=@{
        #ParamName=(($pKeys[0]).Split($script:paramPrefixSep)[1])+($script:paramPrefixSep + $Script:HelpMsgString);
        ParamName=(($pKeys[0]).Split($script:paramPrefixSep)[1])+($script:paramSepDynVar + $Script:HelpMsgString);
        Mandatory=$false; 
        Position=(0+$FirstParamPosition);
        ParamType=[System.String];
    }; # end dynParamSet;
    [void]$dynPList.Add((addDynamicParameterProperties @dynParamSet)); # add parameters to list
    $paramPos=$i+$FirstParamPosition;
    for ($i=1;$i -lt $pKeys.count;$i++)
    {
        $pAttribs=@(($ParamHash[$pKeys[$i]]).Split(($Script:sepVal)));
        $dynParamSet=@{
            #ParamName=(($pKeys[$i]).Split($script:paramPrefixSep)[1])+($script:paramPrefixSep + $Script:HelpMsgString);
            ParamName=(($pKeys[$i]).Split($script:paramPrefixSep)[1])+($script:paramSepDynVar + $Script:HelpMsgString);
            Mandatory=$false; 
            Position=($paramPos);
            ParamType=[System.String];
        }; # end dynParamSet; 
        $paramPos++;               
        $dynPList += (addDynamicParameterProperties   @dynParamSet); # add parameters to list
        $dynParamSet=@{
            #ParamName=(($pKeys[$i]).Split($script:paramPrefixSep)[1])+($script:paramPrefixSep + $Script:DefValString);
            ParamName=(($pKeys[$i]).Split($script:paramPrefixSep)[1])+($script:paramSepDynVar + $Script:DefValString);
            Mandatory=$false; 
            Position=($paramPos);
            ParamType=[System.Type]($pAttribs[0]);
        }; # end dynParamSet;
        $paramPos++;
        $dynPList += (addDynamicParameterProperties @dynParamSet); # add parameters to list 
        if ($i -gt 1) # oly for parameters added with Add-CMMModuleTemplateVariable
        {
            $dynParamSet=@{
                #ParamName=(($pKeys[$i]).Split($script:paramPrefixSep)[1])+($script:paramPrefixSep + $Script:MandatoryString);
                ParamName=(($pKeys[$i]).Split($script:paramPrefixSep)[1])+($script:paramSepDynVar + $Script:MandatoryString);
                Mandatory=$false; 
                Position=($paramPos);
                ParamType=[bool];
            }; # end dynParamSet;
            $paramPos++;
            $dynPList += (addDynamicParameterProperties @dynParamSet); # add parameters to list 
        }; # end if           
    }; # end for
    newDynamicParameters -ParamList $dynPList;
}; # end funciton setTemplateDynamicParamList

function newDynamicParamList
{	      
Param([Parameter(Mandatory = $true, Position = 0)][hashtable]$ParamHash,
      [Parameter(Mandatory = $false, Position = 1)][int]$FirstParamPosition=0,
      [Parameter(Mandatory = $false, Position = 2)][switch]$SetMandatoryToFalse=$false   
     )

    $dynPList = @()   # array for properties of the dynamic parameters
    [array]$pKeys=[System.Linq.Enumerable]::OrderBy([string[]]$ParamHash.Keys, [Func[string,string]] { $args[0] });
    
    for ($i=0;$i -lt $pKeys.count;$i++)
    {
        #$pAttribs=@(($ParamHash[$pKeys[$i]]).Split(($Script:sepVal)));
        $pAttribs=@(($ParamHash[$pKeys[$i]]).Split(($Script:pPropSepVal)));
        $dynParamSet=@{
            ParamName=(($pKeys[$i]).Split($script:paramPrefixSep)[1]);
            #Mandatory=(([bool][int]$pAttribs[1]) -and (! ($setMandatoryToFalse.IsPresent))); # check for overwrite of isMandatory
            # verify if mandatory                   verify if default value                                 verify if setMandatoryToFalse
            Mandatory=([bool][int]$pAttribs[1] -and ([System.String]::IsNullOrEmpty($pAttribs[3])) -and (! ($setMandatoryToFalse.IsPresent))); 
            Position=($i+$FirstParamPosition);
            ParamType=[System.Type]($pAttribs[0]);
        }; # end dynParamSet;
        
        if (!([System.String]::IsNullOrEmpty($pAttribs[2])))
        {
            $dynParamSet.Add('HelpMessage',$pAttribs[2]);
        }; # end if        
        $dynPList += (addDynamicParameterProperties   @dynParamSet); # add parameters to list
    
    }; # end for
    newDynamicParameters -ParamList $dynPList;
}; # end funciton newDynamicParamList
function newDynamicParameters
{
[cmdletbinding()]
Param(
		[Parameter(Mandatory = $true, Position = 0)][Array]$ParamList
     )

    $pEntries = $ParamList.Count   # get number of parameters
    
    $paramDictionary = new-object -Type System.Management.Automation.RuntimeDefinedParameterDictionary

    for ($i = 0; $i -lt $pEntries; $i++)  # for every parameter
    {
        $optionsAvail = $ParamList[$i].Keys -contains "OptionList"
        if ($optionsAvail)
        {
            $options = $ParamList[$i].OptionList    # set options
        }
        $ParamList[$i].Remove("OptionList")     # remove options from list
        $paramName = $ParamList[$i].ParamName   # set parmeter name
        $ParamList[$i].Remove("ParamName")      # remove parameter name from list
        [object]$ParamType = $ParamList[$i].ParamType
        $ParamList[$i].Remove("ParamType")


        $attributes = new-object System.Management.Automation.ParameterAttribute   # init object for attributes
        if ($optionsAvail)
        {
            $ParamOptions = New-Object System.Management.Automation.ValidateSetAttribute -ArgumentList $options
        }

        foreach ($p in $ParamList[$i].keys)          # iterate through the remaining attributes
        {            
            $attributes.$p = $ParamList[$i][$p]   # assign value for the attribute                
        }
        $attributeCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]   # create attribute collection
        $attributeCollection.Add($attributes)       # add the attributes
        if ($optionsAvail)
        {
            $attributeCollection.Add($ParamOptions)     # add the options
        } # end if        

        $dynParam = new-object -Type System.Management.Automation.RuntimeDefinedParameter($ParamName,$ParamType, $attributeCollection)
        $paramDictionary.Add($ParamName, $dynParam)

    } # end for loop
    
    return $paramDictionary	     

}; # end function newDynamicParameters

function addDynamicParameterProperties
{
[cmdletbinding()]
Param(
		[Parameter(Mandatory = $false, Position = 0)][Array]$OptionList,
        [Parameter(Mandatory = $true, Position = 1)][string]$ParamName,
        [Parameter(Mandatory = $true, Position = 2)][Object]$ParamType,
        [Parameter(Mandatory = $false, Position = 3)][int]$Position,
        [Parameter(Mandatory = $false, Position = 4)][bool]$Mandatory,
		[Parameter(Mandatory = $false, Position = 5)][string]$ParameterSetName,
        [Parameter(Mandatory = $false, Position = 6)][bool]$ValueFromPipeline,
        [Parameter(Mandatory = $false, Position = 7)][bool]$ValueFromPipelineByPropertyName,
        [Parameter(Mandatory = $false, Position = 8)][bool]$ValueFromRemainingArguments,
        [Parameter(Mandatory = $false, Position = 9)][string]$HelpMessage,
		[Parameter(Mandatory = $false, Position = 10)][string]$HelpMessageBaseName,
        [Parameter(Mandatory = $false, Position = 11)][string]$HelpMessageResourceId,
        [Parameter(Mandatory = $false, Position = 12)][string]$PresetVal

	 )

    $DynParam = @{}   # create hashtable for the properties of the parameter
    

    foreach ($element in $PSBoundParameters.Keys)   # iterate through the bound parameters
    {        
        $DynParam.add($element,$PSBoundParameters.$element)  # add name and value to the hashtable
    }

    return $DynParam

} # end function addDynamicParameterProperties
#endregion dyn params
