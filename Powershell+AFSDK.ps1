# Load AFSDK  
$AFSDKAssembly = [Reflection.Assembly]::LoadWithPartialName("OSIsoft.AFSDK")

# PI Data Archive
"[Let's play with PI Data Archive]"
$piServers = New-Object OSIsoft.AF.PI.PIServers
$piServer = $piServers.DefaultPIServer
Write-host("PI Data Archive Name: {0}" -f $piServer.Name)

# Get PIPoint
$pt=[OSIsoft.AF.PI.PIPoint]::FindPIPoint($piServer,"sinusoid")
Write-host("TagName: {0}" -f $pt.Name)

# Current Value
$snap = $pt.CurrentValue()
Write-host("Current Value: {0}, Timestamp: {1}" -f $snap.Value, $snap.Timestamp)
"`r`n"

# Recorded Values
$timerange = New-Object OSIsoft.AF.Time.AFTimeRange("*-3h", "*")
$recorded = $pt.RecordedValues($timerange,[OSIsoft.AF.Data.AFBoundaryType]::Inside,"",$false,1000)
"Recoded Values :"
Write-host ($recorded| select Timestamp, value)
"`r`n"

# Interpolated Values  
$span = New-Object OSIsoft.AF.Time.AFTimeSpan(New-TimeSpan -hours 1) 
$interpolated = $pt.InterpolatedValues($timerange, $span, "", $false)  
"Interpolated Values :"
Write-host ($interpolated| select Timestamp, value)
"`r`n"

# Summaries Values  
$summaries = $pt.Summaries($timerange, $span, 
[OSIsoft.AF.Data.AFSummaryTypes]::Average,
[OSIsoft.AF.Data.AFCalculationBasis]::TimeWeighted,
[OSIsoft.AF.Data.AFTimestampCalculation]::Auto)  
"Summaries Values :"
Write-host($summaries.Keys)
Write-host($summaries.Values)
"`r`n"

#writeValue
"Write value to the PI Tag :"
$writept = [OSIsoft.AF.PI.PIPoint]::FindPIPoint($piServer,"test999")
$aftime = New-Object OSIsoft.AF.Time.AFTime("t+9h")
$writeval = New-Object OSIsoft.AF.Asset.AFValue(10,$aftime)
$writept.UpdateValue($Writeval,[OSIsoft.AF.Data.AFUpdateOption]::Replace,[OSIsoft.AF.Data.AFBufferOption]::BufferIfPossible)
Write-host("Tag : {0} Wrote Value: {1} Timestamp: {2}" -f $writept.Name, $writeval.Value, $writeval.Timestamp.LocalTime.ToString())
"`r`n"

"[Let's play with PI AF]"
# AFServer
$afServers = New-Object OSIsoft.AF.PISystems  
$afServer = $afServers.DefaultPISystem
Write-host("AFServer Name: {0}" -f $afServers.Name)

# AFDatabase
$DB = $afServer.Databases["NuGreen"]
Write-host("DataBase Name: {0}" -f $DB.Name)# Line Break
"`r`n"
$DB.UndoCheckOut($true)

# Get Element
$Element = $DB.Elements["NuGreen"].Elements["Little Rock"].Elements["Extruding Process"].Elements["Equipment"].Elements["K-435"]
Write-host("Element Name: {0}" -f $Element.Name)

# Get attribute
$Attribute = $Element.Attributes["Steam Flow"]
Write-host("Attribute Name: {0}" -f $Attribute.Name)
$Attval = $Attribute.GetValue()
Write-host("Value: {0}, Timestamp: {1}" -f $Attval.Value, $Attval.Timestamp)
"`r`n"

# Create new element with attribute  
"Try to create new Element with Attribute :"
if (!$DB.Elements["Powershell New Element"])
{
    $newelement = $DB.Elements.Add("Powershell New Element")
    $newelement.Description = "Created element from Powershell"
    $newattribute = $newelement.Attributes.Add("Powershell Attribute")  
    $newattribute.DataReferencePlugIn = $afServer.DataReferencePlugIns["PI Point"]
    $newattribute.DataReference.ConfigString = "cdt158"  
    $DB.CheckIn()
    Write-host("Created new Element : {0}" -f $newelement.Name)
}else{
    "The element is already created"
}

