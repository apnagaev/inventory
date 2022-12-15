#############################
$createurl = $jiraserver+'/rest/assets/1.0/object/create'
$updateurl = $jiraserver+'/rest/assets/1.0/object/'
$allurl = $jiraserver+'/rest/assets/1.0/aql/objects?resultPerPage=999999'
$userurl=$jiraserver+'/rest/api/2/user/search?username='
$updateurlclear=$jiraserver+'/rest/assets/1.0/object/'
$objsoft=112
$softaatr=@(991, 1000, 1171)
####################################
$ver='Loader:'+$ver+' Script:3.5.2'
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
#$upt=[math]::Round($uptime.TotalHours,1) -replace ",","."
$upt=((gcim Win32_OperatingSystem).LastBootUpTime).toString("yyyy-MM-ddTHH:mm:ssZ")

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
$compfqdn.HostName=$compfqdn.HostName.ToLower()


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



if ($compfqdn.HostName -match 'srv') {$objectTypeId=52}

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

$compfqdn.HostName
if ($compfqdn.HostName -match "00\d+"){$invnumber = $compfqdn.HostName -match "00\d+" |%{$matches[0]}}
$invnumber


if (($manuname.PCSystemType -eq 1) -or ($manuname.PCSystemType -eq 3)){#Workstation
$objectTypeId=65
$attributevar=@(564, 581, 975, 583, 584, 977, 585, 980, 976, 978, 590, 579, 981, 979, 596, 1154, 1100, 1178, 1165, 1182)
$allurlpc=$allurl+'&includeAttributes=false&qlQuery=objectType="Workstations"'
$bm=$true
}

if ($manuname.PCSystemType -eq 2){#Laptop
$objectTypeId=66
$attributevar=@(564, 581, 975, 583, 584, 977, 585, 980, 976, 978, 590, 579, 981, 979, 596, 1154, 1100, 1178, 1165, 1182)
$allurlpc=$allurl+'&includeAttributes=false&qlQuery=objectType="Laptops"'
$bm=$true
}

if (($manuname.PCSystemType -eq 0) -or ($manuname.PCSystemType -gt 3) -or ($compfqdn.HostName -match 'srv')){
    if ($null -eq ($virtvendor | ? { $manuname.Manufacturer -match $_ })){
        #baremetal
        $objectTypeId=102
        $allurlpc=$allurl+'&includeAttributes=false&qlQuery=objectType="BareMetal"'
        $bm=$true
    }
    else{
        #virtual
        $objectTypeId=103
        $allurlpc=$allurl+'&includeAttributes=false&qlQuery=objectType="Virtual"'
        $bm=$false
    }
$attributevar=@(564, 581, 983, 583, 584, 986, 585, 989, 985, 987, 590, 579, 984, 988, 596, 1159, 1100, 1179, 1166, 1183)
$invnumber=''
}

if ($null -ne ($virtvendor | ? { $manuname.Manufacturer -match $_ })){
    $objectTypeId=103
    $allurlpc=$allurl+'&includeAttributes=false&qlQuery=objectType="Virtual"'
    $attributevar=@(564, 581, 983, 583, 584, 986, 585, 989, 985, 987, 590, 579, 984, 988, 596, 1159, 1100, 1179, 1166, 1183)
    $invnumber='na'
    $bm=$false
}


$allurlpc
$allobj=Invoke-RestMethod -Uri $allurlpc -Headers @{Authorization=("Basic {0}" -f $base64)} -ContentType 'application/json; charset=utf-8'

##########find device id#############
ForEach ($item in $allobj.objectEntries){
    if ($compfqdn.HostName -eq $item.name){
    $item.name
    $item.id
    $deviceid=$item.id
    $compobg=$item.objectKey
    }
}#$compinfo.Name



##########check object and create if null#############
if (($null -eq ($allobj.objectEntries.name | ? { $compfqdn.HostName -match $_ })) -and ($bm))  {
    $body='{"objectSchemaKey":"'+$objectSchemaKey+'", "objectTypeId":'+$objectTypeId+',"attributes": [{"objectTypeAttributeId":'+$attributevar[0]+',"objectAttributeValues": [{"value": "'+$compfqdn.HostName+'"}]}]}'
    #$body = [System.Text.Encoding]::UTF8.GetBytes($body)
    
    Write-Host ('create new object')
    #$body = [System.Text.Encoding]::UTF8.GetBytes($body)
    Invoke-RestMethod -Uri $createurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Post' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose
}

Start-Sleep 60

$allobj=Invoke-RestMethod -Uri $allurlpc -Headers @{Authorization=("Basic {0}" -f $base64)} -ContentType 'application/json; charset=utf-8'

##########find device id#############
ForEach ($item in $allobj.objectEntries){
    if ($compfqdn.HostName -eq $item.name){
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




$osname=(Get-WmiObject -class Win32_OperatingSystem).Caption
$osname = $osname -replace 'Майкрософт','Microsoft'
$osname = $osname -replace 'Профессиональная','Pro'
$osname = $osname -replace 'Корпоративная','Corp'
$osname = $osname -replace ' (Registered Trademark)',''
$allosurl=$allurl+'&qlQuery=objectType="OS"'
$allosurl
$allos=Invoke-RestMethod -Uri $allosurl -Headers @{Authorization=("Basic {0}" -f $base64)} -ContentType 'application/json; charset=utf-8'
if ($allos.iql -ne $null) {
    #$alljsmsoft
        foreach ($ositem in $allos.objectEntries){
            $ositem.name
            if ($allos.objectEntries.name -notcontains $osname){
                                    $body='{
	                                "objectSchemaKey": "'+$objectSchemaKey+'",
	                                "objectTypeId": 355,
	                                "attributes": [{
			                            "objectTypeAttributeId":991,
			                                "objectAttributeValues": [{
				                                "value": "'+$osname+'"
			                                }]
		                                }
	                                ]
                                }'
                        Write-Host('Create object')
                        $body
                        #$body = [System.Text.Encoding]::UTF8.GetBytes($body)
                        Invoke-RestMethod -Uri $createurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Post' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose
    
        }
        else{
            if ($ositem.name -eq $osname){
            $ositem.obgkey
                $body='{
	            "objectSchemaKey": "'+$objectSchemaKey+'",
	            "attributes": [{
		            "objectTypeAttributeId":'+$attributevar[19]+',
		            "objectAttributeValues": [{
				"value":"'+$ositem.objectKey+'"
			            }
		            ]
	            }]
            }'
            $body
            #$body = [System.Text.Encoding]::UTF8.GetBytes($body)
            Invoke-RestMethod -Uri $updateurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Put' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose

        
            }
        }

    }
}







if ($null -eq ($virtvendor | ? { $manuname.Manufacturer -match $_ })) {
    $compatt=$compatt+',
    {"objectTypeAttributeId":'+$attributevar[16]+',
      "objectAttributeValues": [
        {
          "value":"true"
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
    {"objectTypeAttributeId":'+$attributevar[8]+',
      "objectAttributeValues": [
        {
          "value":"'+$rcpu+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[7]+',
      "objectAttributeValues": [
        {
          "value":"'+$winver+'"
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
    {"objectTypeAttributeId":'+$attributevar[13]+',
      "objectAttributeValues": [
        {
          "value":"'+$motherboard+'"
        }
      ]},
    {"objectTypeAttributeId":'+$attributevar[18]+',
      "objectAttributeValues": [
        {
          "value":"'+$localip+'"
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
#$compatt=$compatt+''



$body='{
  "objectSchemaKey":"'+$objectSchemaKey+'",
  "objectTypeId":'+$objectTypeId+',
  "attributes": [
    {"objectTypeAttributeId":'+$attributevar[0]+',
      "objectAttributeValues": [
        {
          "value":"'+$compfqdn.HostName+'"
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
    {"objectTypeAttributeId":'+$attributevar[12]+',
      "objectAttributeValues": [
        {
          "value":"'+$user+'"
        }
      ]}'+$compatt+'
  ]
}'

$body=$body -replace '\\',''
$body
if ($PSVersionTable.PSVersion.Major -lt 5){$body = [System.Text.Encoding]::UTF8.GetBytes($body)}
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



$checkmon=0
if ($localip -match ('10.177.\d+.\d+')){$checkmon=1}
if ($localip -match ('192.168.144.\d+')){$checkmon=1}
if ($localip -match ('10.63.\d+.\d+')){$checkmon=1}
if ($localip -match ('10.197.\d+.\d+')){$checkmon=1}
if ($localip -match ('10.97.\d+.\d+')){$checkmon=1}
if ($localip -match ('10.99.\d+.\d+')){$checkmon=1}
if ($localip -match ('10.77.\d+.\d+')){$checkmon=1}
if ($localip -match ('10.150.\d+.\d+')){$checkmon=1}
if (!$bm){$checkmon=0}

if ($checkmon -eq 1){
$allurlmon=$allurl+'&includeAttributes=false&qlQuery=objectType="Monitors"'
$allmon=Invoke-RestMethod -Uri $allurlmon -Headers @{Authorization=("Basic {0}" -f $base64)} -ContentType 'application/json; charset=utf-8'
$progress = 1
foreach ($monitor in (Get-WmiObject WmiMonitorID -Namespace root/wmi)) {
#Write-Host "Monitor #$($progress):" -ForegroundColor Green
#Write-Host "Manufactur: $(($monitor.ManufacturerName | ForEach-Object {[char]$_}) -join '')"
$monmanufact=$(($monitor.ManufacturerName | ForEach-Object {[char]$_}) -join '')
#Write-Host "PN: " ($($monitor.UserFriendlyName | ForEach-Object {[char]$_}) -join '')
$monpn=($($monitor.UserFriendlyName | ForEach-Object {[char]$_}) -join '')
#Write-Host "SN: " ($($monitor.SerialNumberID | ForEach-Object {[char]$_}) -join '')
$monsn=($($monitor.SerialNumberID | ForEach-Object {[char]$_}) -join '')
#Write-Host "WeekOfManufacture: $($monitor.WeekOfManufacture)"
#Write-Host "YearOfManufacture: $($monitor.YearOfManufacture)"

$monmanufact = $monmanufact -replace "\W",""
$monpn = $monpn -replace "\W",""
$monsn = $monsn -replace "\W",""
$monpn = $monpn -replace $monmanufact,''

if ($monmanufact -eq 'ACI'){$monmanufact='ASUS'}
if ($monmanufact -eq 'BNQ'){$monmanufact='BenQ'}
if ($monmanufact -eq 'DEL'){$monmanufact='DELL'}
if ($monmanufact -eq 'IVM'){$monmanufact='ProLite'}
if ($monmanufact -eq 'LEN'){$monmanufact='Lenovo'}
if ($monmanufact -eq 'PHL'){$monmanufact='Philips'}
if ($monmanufact -eq 'BK'){$monmanufact='LG'}
if ($monmanufact -eq 'VSC'){$monmanufact='ViewSonic'}
if ($monmanufact -eq 'ACR'){$monmanufact='Acer'}
if ($monmanufact -eq 'GSM'){$monmanufact='LG'}
if ($monmanufact -eq 'MEI'){$monmanufact='Panasonic'}
if ($monmanufact -eq 'SAM'){$monmanufact='Samsung'}
if ($monmanufact -eq 'HPN'){$monmanufact='HP'}


$monpn = $monpn -replace $monmanufact,''
$monobj=$monmanufact+' '+$monpn+' '+$monsn

$locationurl=$updateurlclear+$compobg
$locref=Invoke-RestMethod -Uri $locationurl -Headers @{Authorization=("Basic {0}" -f $base64)} -ContentType 'application/json; charset=utf-8'

foreach ($locitem in $locref.attributes){
    if  ($locitem.objectTypeAttribute.id -eq '594'){
    $locitem.objectAttributeValues.Value
    $locitem.objectAttributeValues.referencedObject.objectKey
    $location=$locitem.objectAttributeValues.referencedObject.objectKey
    }  
}



$localip
    if (($allmon.objectEntries.name -notcontains $monobj) -and ($monpn -ne '') -and ($monsn -ne '0') -and ($checkmon -eq 1)){
        $body='{
  "objectSchemaKey":"'+$objectSchemaKey+'",
  "objectTypeId":74,
  "attributes": [
    {"objectTypeAttributeId":564,
      "objectAttributeValues": [
        {
          "value":"'+$monobj+'"
        }
      ]},
    {"objectTypeAttributeId":583,
      "objectAttributeValues": [
        {
          "value":"'+$monmanufact+'"
        }
      ]},
    {"objectTypeAttributeId":584,
      "objectAttributeValues": [
        {
          "value":"'+$monpn+'"
        }
      ]},
    {"objectTypeAttributeId":585,
      "objectAttributeValues": [
        {
          "value":"'+$monsn+'"
          }
        ]},
    {"objectTypeAttributeId":1261,
      "objectAttributeValues": [
        {
          "value":"'+$compobg+'"
        }
      ]},
    {"objectTypeAttributeId":1100,
      "objectAttributeValues": [
        {
          "value":"true"
        }
      ]},
    {"objectTypeAttributeId":594,
      "objectAttributeValues": [
        {
          "value":"'+$location+'"
        }
      ]}
  ]
}'
$body
 Invoke-RestMethod -Uri $createurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Post' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose
 
    }


$allmon.objectEntries.name -notcontains $monobj
$monpn -ne ''
$monsn -ne '0'
$checkmon -eq 1

$progress++



##########update device id#############
if (($allmon.objectEntries.name -contains $monobj) -and ($checkmon -eq 1)){

        ##########find device id#############
        ForEach ($item in $allmon.objectEntries){
            if ($monobj -eq $item.name){
            $item.name
            $item.id
            $monid=$item.id

            }
        }
$updatemonurl=$updateurlclear+$monid



if (($allmon.objectEntries.name -contains $monobj) -and ($checkmon -eq 1)){
        $body='{
  "objectSchemaKey":"'+$objectSchemaKey+'",
  "objectTypeId":74,
  "attributes": [
    {"objectTypeAttributeId":1261,
      "objectAttributeValues": [
        {
          "value":"'+$compobg+'"
        }
      ]},
    {"objectTypeAttributeId":594,
      "objectAttributeValues": [
        {
          "value":"'+$location+'"
        }
      ]},
    {"objectTypeAttributeId":584,
      "objectAttributeValues": [
        {
          "value":"'+$monpn+'"
        }
      ]}
  ]
}'
$body
$updatemonurl
 #Invoke-RestMethod -Uri $updatemonurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Put' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose

    }




}
}
}

