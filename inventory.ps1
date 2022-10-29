#############################
$createurl = 'https://jirasm.atol.ru/rest/assets/1.0/object/create'
$updateurl = 'https://jirasm.atol.ru/rest/assets/1.0/object/'
$allurl = 'https://jirasm.atol.ru/rest/assets/1.0/aql/objects?resultPerPage=999999'
$userurl='https://jirasm.atol.ru/rest/api/2/user/search?username='
$objectSchemaKey='AS'
$objsoft=112
$softaatr=@(991, 1000, 1171)
####################################
ver='3.3.1'
#########################
cls
$sleep = Get-Random -Maximum 900
start-sleep $sleep
$badadapters=@('TAP','Cisco AnyConnect','Bluetooth','Fibocom','VirtualBox')
$virtvendor=@('VMware','Microsoft')
$mac=''
$hostsoft=''
$allobj=''
$invnumber='n\\a'
$jirasmsoft=''
$tdisk=''
$compatt=''
$upt=''
Get-Command '*json'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$compinfo = Get-CimInstance -ClassName Win32_ComputerSystem
$uptime = (get-date) - (gcim Win32_OperatingSystem).LastBootUpTime 
$upt=[math]::Round($uptime.TotalHours,1) -replace ",","."

$network = Get-NetConnectionProfile
$network | ConvertTo-Json
$network.InterfaceAlias
$localip = Get-NetIPAddress -InterfaceAlias $network.InterfaceAlias
$localip = $localip.IPv4Address
$localip

$manuname = Get-CimInstance -ClassName Win32_ComputerSystem
$manuname | ConvertTo-Json
$manuname.Manufacturer
$manuname.SystemFamily
$manuname.Model
$serial = Get-wmiobject win32_bios | ForEach-Object {$_.serialnumber}
$manuname.DNSHostName
$manuname.Domain
$manuname.UserName
$manuname.NumberOfProcessors
$manuname.NumberOfLogicalProcessors
$motherboard = Get-CimInstance -Class Win32_BaseBoard #| Format-Table Manufacturer, Product, SerialNumber, Version -Auto
$motherboard = $motherboard.Manufacturer+' P\\N:'+$motherboard.Product+' S\\N:'+$motherboard.SerialNumber+' Ver:'+$motherboard.Version

$compfqdn=[System.Net.Dns]::GetHostByName($env:computerName)
$compfqdn.HostName


if ($manuname.PCSystemType -eq '') {$PCSystemType = ''}
if ($manuname.PCSystemType -eq '6') {$PCSystemType = 'Appliance PC'}
if ($manuname.PCSystemType -eq '1') {$PCSystemType = 'Desktop'}
if ($manuname.PCSystemType -eq '4') {$PCSystemType = 'Enterprise Server'}
if ($manuname.PCSystemType -eq '8') {$PCSystemType = 'other'}
if ($manuname.PCSystemType -eq '2') {$PCSystemType = 'Mobile device'}
if ($manuname.PCSystemType -eq '7') {$PCSystemType = 'Performance server'}
if ($manuname.PCSystemType -eq '5') {$PCSystemType = 'SOHO Server'}
if ($manuname.PCSystemType -eq '0') {$PCSystemType = 'unspecified'}
if ($manuname.PCSystemType -eq '3') {$PCSystemType = 'Workstation'}
$memory=[math]::Round([long]$manuname.TotalPhysicalMemory/([math]::Pow(1024,3)),0)
$PCSystemType
try{
    $winver=(Get-WmiObject -class Win32_OperatingSystem).Caption+' '+(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion+' Build '+[System.Environment]::OSVersion.Version.Build
}
catch {$winver=(Get-WmiObject -class Win32_OperatingSystem).Caption}
$cpu = Get-WmiObject -Class Win32_Processor | select *
$cpu.name.Count
if ($cpu.name.Count -gt 1){$rcpu = $cpu.name[0]} else {$rcpu = $cpu.name}
$rcpu
$disk = Get-PhysicalDisk
foreach ($item in $disk){
    $tdisk = $item.MediaType+' '+$item.FriendlyName+' '+[math]::Round([long]$item.size/([math]::Pow(1024,3)),0)+'Gb'+ '; '+$tdisk
}
$disk = $tdisk
$disk
$allnet = get-netadapter
ForEach ($item in $allnet){
    if (($null -eq ($badadapters | ? { $item.InterfaceDescription -match $_ })) -and ($item.MacAddress -ne '') -and ($item.InterfaceDescription -ne $null)){
    $mac=$mac+$item.InterfaceDescription+' '+$item.MacAddress+' '
    }
}



if ($compinfo.Name.ToLower() -match 'srv') {$objectTypeId=52}

$computer='localhost'
$user = gwmi -Class win32_computersystem -ComputerName "localhost" | select -ExpandProperty username -ErrorAction Stop 

if ($user -eq $null){
    $rdp = QUERY SESSION
    $rdp = $rdp  -replace "\s+", ";"
    $rdp = $rdp  -replace 'Active', 'Активно'
    $rdp = $rdp -match 'rdp-tcp#'
    $rdp = $rdp -match 'Активно'
    $rdp = $rdp | ConvertFrom-Csv -Delimiter ';' -Header 'session','user','id','status'

    
    if (($rdp[0].user -ne '') -and ($rdp[0].user -ne $null)){
        $user = $rdp[0].user
	if ($user -match 'rdp-tcp'){$user = $rdp[0].id}
    }
}
if ($user -match 'ATOL\\'){$user = $user -replace 'ATOL\\',''}
if ($user -match 'NAGAEV\\'){$user = $user -replace 'NAGAEV\\',''}

if ($user -notmatch '@atol.ru'){$user = $user + '@atol.ru'}
#$user = $user -replace '\.',''

$userurl=$userurl+$user
#$userkey
$userkey=Invoke-RestMethod -Uri $userurl -Headers @{Authorization=('Basic {0}' -f $base64)} -ContentType 'application/json; charset=utf-8'

$userkeykey=''
if ($userkeykey.count -gt 1){
$user
    #$userkeykey=$userkey.key[1]
    foreach ($item in $userkey){
    $item.emailAddress
        if ($item.emailAddress -eq $user){$userkeykey=$item.key}
    }
}
else{$userkeykey=$userkey.key}
if ($userkeykey.count -gt 1){$userkeykey=$userkeykey[0]}

$compinfo.Name
if ($compinfo.Name -match "\d+$"){$invnumber = $compinfo.Name -match "\d+"|%{$matches[0]}}
$invnumber


if (($manuname.PCSystemType -eq 1) -or ($manuname.PCSystemType -eq 3)){#Workstation
$objectTypeId=65
$attributevar=@(564, 581, 975, 583, 584, 977, 585, 980, 976, 978, 590, 579, 981, 979, 596, 1154, 1100, 1155, 1165)
$allurlpc=$allurl+'&includeAttributes=false&qlQuery=objectType="Workstations"'
}

if ($manuname.PCSystemType -eq 2){#Laptop
$objectTypeId=66
$attributevar=@(564, 581, 975, 583, 584, 977, 585, 980, 976, 978, 590, 579, 981, 979, 596, 1154, 1100, 1155, 1165)
$allurlpc=$allurl+'&includeAttributes=false&qlQuery=objectType="Laptops"'
}

if (($manuname.PCSystemType -eq 0) -or ($manuname.PCSystemType -gt 3) -or ($compinfo.Name.ToLower() -match 'srv')){
    if ($null -eq ($virtvendor | ? { $manuname.Manufacturer -match $_ })){
        #baremetal
        $objectTypeId=102
        $allurlpc=$allurl+'&includeAttributes=false&qlQuery=objectType="BareMetal"'
    }
    else{
        #virtual
        $objectTypeId=103
        $allurlpc=$allurl+'&includeAttributes=false&qlQuery=objectType="Virtual"'
    }
$attributevar=@(564, 581, 983, 583, 584, 986, 585, 989, 985, 987, 590, 579, 984, 988, 596, 1159, 1100, 1156, 1166)
$invnumber='na'
}


$allurlpc
$allobj=Invoke-RestMethod -Uri $allurlpc -Headers @{Authorization=("Basic {0}" -f $base64)} -ContentType 'application/json; charset=utf-8'

##########find device id#############
ForEach ($item in $allobj.objectEntries){
    if ($compinfo.Name -eq $item.name){
    $item.name
    $item.id
    $deviceid=$item.id
    }
}


##########check object and create if null#############
if ($null -eq ($allobj.objectEntries.label | ? { $compinfo.Name.ToLower() -match $_ })) {
    $body='{"objectSchemaKey":"'+$objectSchemaKey+'", "objectTypeId":'+$objectTypeId+',"attributes": [{"objectTypeAttributeId":'+$attributevar[0]+',"objectAttributeValues": [{"value": "'+$compinfo.Name.ToLower()+'"}]}]}'
    #$body = [System.Text.Encoding]::UTF8.GetBytes($body)
    Write-Host ('create new object')
    #$body = [System.Text.Encoding]::UTF8.GetBytes($body)
    Invoke-RestMethod -Uri $createurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Post' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose
}

Start-Sleep 60

$allobj=Invoke-RestMethod -Uri $allurlpc -Headers @{Authorization=("Basic {0}" -f $base64)} -ContentType 'application/json; charset=utf-8'

##########find device id#############
ForEach ($item in $allobj.objectEntries){
    if ($compinfo.Name -eq $item.name){
    $item.name
    $item.id
    $deviceid=$item.id
    }
}



##########update device id#############
$updateurl=$updateurl+$deviceid
#$body='{"objectSchemaKey":"$objectSchemaKey", "objectTypeId": $objectTypeId,"attributes": [{"objectTypeAttributeId": 342,"objectAttributeValues": [{"value": "$compinfo.Name"}]}]}'

#$nowuser = Invoke-RestMethod -Uri $updateurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Get' -ContentType 'application/json; charset=utf-8' -Verbose
#$userkeykey=$userkey.key
#if ($userkeykey.count -gt 1){$userkeykey=$userkey.key[1]}

$object = Invoke-RestMethod -Uri $updateurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Get' -ContentType 'application/json; charset=utf-8' -Verbose
ForEach ($item in $object.attributes){
    if ($item.objectTypeAttributeId -eq $attributevar[12]){
    if ($user -eq $null){$user=$item.objectAttributeValues.value}
    }
}
ForEach ($item in $object.attributes){
    if ($item.objectTypeAttributeId -eq $attributevar[11]){
    $item.objectAttributeValues.searchValue
    #$user=$item.objectAttributeValues.value
    if ($item.objectAttributeValues.searchValue -ne $null){$userkeykey=$item.objectAttributeValues.searchValue}
    }
}



if ($null -eq ($virtvendor | ? { $manuname.Manufacturer -match $_ })) {
    $compatt=$compatt+',
    {"objectTypeAttributeId":'+$attributevar[16]+',
      "objectAttributeValues": [
        {
          "value":"true"
        }
      ]}'
}
else{
$compatt=$compatt+',
    {"objectTypeAttributeId":'+$attributevar[16]+',
      "objectAttributeValues": [
        {
          "value":"false"
        }
      ]}'
}
$compatt=$compatt+',
    {"objectTypeAttributeId":'+$attributevar[17]+',
      "objectAttributeValues": [
        {
          "value":"'+$upt+'"
        }
      ]}'
$compatt=$compatt+',
    {"objectTypeAttributeId":'+$attributevar[18]+',
      "objectAttributeValues": [
        {
          "value":"'+$localip+'"
        }
      ]}'



$body='{
  "objectSchemaKey":"'+$objectSchemaKey+'",
  "objectTypeId":'+$objectTypeId+',
  "attributes": [
    {"objectTypeAttributeId":'+$attributevar[0]+',
      "objectAttributeValues": [
        {
          "value":"'+$compinfo.Name.ToLower()+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[1]+',
      "objectAttributeValues": [
        {
          "value":"'+$invnumber+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[2]+',
      "objectAttributeValues": [
        {
          "value":"'+$compfqdn.HostName+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[3]+',
      "objectAttributeValues": [
        {
          "value":"'+$manuname.Manufacturer+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[4]+',
      "objectAttributeValues": [
        {
          "value":"'+$manuname.Model+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[5]+',
      "objectAttributeValues": [
        {
          "value":"'+$memory+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[6]+',
      "objectAttributeValues": [
        {
          "value":"'+$serial+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[7]+',
      "objectAttributeValues": [
        {
          "value":"'+$winver+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[8]+',
      "objectAttributeValues": [
        {
          "value":"'+$rcpu+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[9]+',
      "objectAttributeValues": [
        {
          "value":"'+$disk+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[10]+',
      "objectAttributeValues": [
        {
          "value":"'+$mac+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[11]+',
      "objectAttributeValues": [
        {
          "value":"'+$userkeykey+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[12]+',
      "objectAttributeValues": [
        {
          "value":"'+$user+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[13]+',
      "objectAttributeValues": [
        {
          "value":"'+$motherboard+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[14]+',
      "objectAttributeValues": [
        {
          "value":"'+$PCSystemType+' '+$localip.IPAddress+'"
        }
      ]}'+$compatt+'
  ]
}'

$body=$body -replace '\\',''
$body
#$body = [System.Text.Encoding]::UTF8.GetBytes($body)
Invoke-RestMethod -Uri $updateurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Put' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose
$updateurl
$i=0
$c=0

$soft = New-Object System.Collections.ArrayList
$soft = Get-CimInstance -Class Win32_Product | Select-Object -Property Name, Version
$soft=$soft | ConvertTo-Csv
$soft = $soft | select -uniq
$soft = $soft | ConvertFrom-Csv -Delimiter ','
#$soft


$i=0
$allsofturl=$allurl+'&qlQuery=objectType="Software"'
$allsofturl
#$body = [System.Text.Encoding]::UTF8.GetBytes($body)
$alljsmsoft=Invoke-RestMethod -Uri $allsofturl -Headers @{Authorization=("Basic {0}" -f $base64)} -ContentType 'application/json; charset=utf-8'
if ($alljsmsoft.iql -ne $null) {
#$alljsmsoft
foreach ($jsmitem in $alljsmsoft.objectEntries){
    #$jsmitem.name.Trim()
    #$jsmitem.attributes.objectAttributeValues.value[1]
    $jirasmsoft=$jirasmsoft+$jsmitem.name.Trim()+','+$jsmitem.attributes.objectAttributeValues.value[2]+','+$jsmitem.objectKey+"`n"
}

$jirasmsoft = $jirasmsoft | ConvertFrom-Csv -Delimiter ',' -Header 'name','version','objkey'
#$jirasmsoft
#$soft


foreach ($softitem in $soft){
$i=0
$softnv=$softitem.name.Trim()+','+$softitem.version
    foreach ($jsmitem in $jirasmsoft){
        $jsmnv=$jsmitem.name.Trim()+','+$jsmitem.version
            #$softnv
            #$jsmnv
            if($softnv -eq $jsmnv){
            #$softnv
            Write-Host('dont create, exit')
            $i=$i+1
            #$i    
            #$jsmitem.objkey
            $hostsoft=$jsmitem.objkey+';'+$hostsoft           
            break
            }
    }


$i
if ($i -eq 0){
#$softnv
$body='{
	            "objectSchemaKey": "'+$objectSchemaKey+'",
	            "objectTypeId": '+$objsoft+',
	            "attributes": [{
			        "objectTypeAttributeId": '+$softaatr[0]+',
			            "objectAttributeValues": [{
				            "value": "'+$softitem.name.Trim()+'"
			            }]
		            },
		            {
			            "objectTypeAttributeId": '+$softaatr[1]+',
			            "objectAttributeValues": [{
				            "value": "'+$softitem.version+'"
			            }]
		            },
		            {
			            "objectTypeAttributeId": '+$softaatr[2]+',
			            "objectAttributeValues": [{
				            "value": "'+$compfqdn.HostName+'"
			            }]
		            }
	            ]
            }'
    Write-Host('Create object')
    $body
    #$body = [System.Text.Encoding]::UTF8.GetBytes($body)
    Invoke-RestMethod -Uri $createurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Post' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose
    }
}



$attach = $hostsoft -split ';'
$attach = $attach | select -uniq

$att=''
foreach ($item in $attach){
    $att=$att+'{
				"value":"'+$item+'"
			},'

}
$att=$att -replace ".{1}$"

$body='{
	"objectSchemaKey": "'+$objectSchemaKey+'",
	"attributes": [{
		"objectTypeAttributeId": "'+$attributevar[15]+'",
		"objectAttributeValues": ['+$att+'
		]
	}]
}'
$body
#$body = [System.Text.Encoding]::UTF8.GetBytes($body)
Invoke-RestMethod -Uri $updateurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Put' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose
}
