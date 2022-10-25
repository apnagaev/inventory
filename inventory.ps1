#############################
$createurl = 'https://jirasm.atol.ru/rest/assets/1.0/object/create'
$updateurl = 'https://jirasm.atol.ru/rest/assets/1.0/object/'
$allurl = 'https://jirasm.atol.ru/rest/assets/1.0/aql/objects?ncludeAttributes=false&resultPerPage=999999'
$userurl='https://jirasm.atol.ru/rest/api/2/user/search?username='
####################################
ver='2.5'
cls
$badadapters=@('TAP-Windows','Cisco AnyConnect','Bluetooth','Fibocom')
$mac=''
$hostsoft=''
Get-Command '*json'
$compinfo = Get-CimInstance -ClassName Win32_ComputerSystem

'

$localip = Get-NetIPAddress -InterfaceAlias $network.InterfaceAlias
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
$winver='Windows '+[System.Environment]::OSVersion.Version.Major+' '+(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name DisplayVersion).DisplayVersion+' Build '+[System.Environment]::OSVersion.Version.Build
}
catch {$winver=(Get-WmiObject -class Win32_OperatingSystem).Caption}
$cpu = Get-WmiObject -Class Win32_Processor | select *
$cpu.name.Count
if ($cpu.name.Count -gt 1){$rcpu = $cpu.name[0]} else {$rcpu = $cpu.name}
$rcpu
$disk = Get-PhysicalDisk
$disk=$disk.MediaType+' '+$disk.FriendlyName+' '+[math]::Round([long]$disk.size/([math]::Pow(1024,3)),0)+'Gb'
$allnet = get-netadapter
ForEach ($item in $allnet){
    if (($null -eq ($badadapters | ? { $item.InterfaceDescription -match $_ })) -and ($item.MacAddress -ne '') -and ($item.InterfaceDescription -ne $null)){
    $mac=$mac+$item.InterfaceDescription+' '+$item.MacAddress+' '
    }
}



if ($manuname.PCSystemType -eq 1){$objectTypeId=41}
if ($manuname.PCSystemType -eq 2){$objectTypeId=42}
if ($manuname.PCSystemType -eq 3){$objectTypeId=41}
if ($manuname.PCSystemType -eq 4){$objectTypeId=52}
if ($manuname.PCSystemType -eq 5){$objectTypeId=52}
if ($manuname.PCSystemType -eq 6){$objectTypeId=52}
if ($manuname.PCSystemType -eq 7){$objectTypeId=52}
if ($manuname.PCSystemType -eq 8){$objectTypeId=52}
if ($manuname.PCSystemType -eq 0){$objectTypeId=52}

if ($compinfo.Name.ToLower() -match 'srv') {$objectTypeId=52}

$computer='localhost'
$user = gwmi -Class win32_computersystem -ComputerName "localhost" | select -ExpandProperty username -ErrorAction Stop 

if ($user -eq $null){
    $rdp = QUERY SESSION
    $rdp = $rdp  -replace "\s+", ";"
    $rdp = $rdp  -replace "Active", "Активно"
    $rdp = $rdp -match 'rdp-tcp#'
    $rdp = $rdp -match 'Активно'
    $rdp = $rdp | ConvertFrom-Csv -Delimiter ';' -Header 'session','user','id','status'

    
    if (($rdp[0].user -ne '') -and ($rdp[0].user -ne $null)){
        $user = $rdp[0].user
    }
}
if ($user -match 'ATOL\\'){$user = $user -replace 'ATOL\\',''}
if ($user -match 'NAGAEV\\'){$user = $user -replace 'NAGAEV\\',''}

if ($user -notmatch '@atol.ru'){$user = $user + '@atol.ru'}
#$user = $user -replace '\.',''

$userurl=$userurl+$user
$userkey=Invoke-RestMethod -Uri $userurl -Headers @{Authorization=("Basic {0}" -f $base64)}



$invnumber = $compinfo.Name -match "\d+"|%{$matches[0]}

if ($objectTypeId -eq 42){
$attributevar=@(342, 386, 395, 410, 408, 411, 397, 414, 406, 412, 458, 385, 409, 413, 460, 562)
$allurlpc=$allurl+'&qlQuery=objectType="Laptops"'
}

if ($objectTypeId -eq 41){
$attributevar=@(338, 353, 396, 415, 416, 419, 417, 422, 418, 420, 437, 355, 423, 421, 439, 561)
$allurlpc=$allurl+'&qlQuery=objectType="Computers"'
}

if ($objectTypeId -eq 52){
$attributevar=@(463, 481, 482, 483, 484, 487, 485, 491, 486, 488, 490, 478, 480, 489, 494)
$invnumber='N\\A'
$allurlpc=$allurl+'&qlQuery=objectType="Servers"'
}
$allobj=Invoke-RestMethod -Uri $allurlpc -Headers @{Authorization=("Basic {0}" -f $base64)}


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
    $body
    Invoke-RestMethod -Uri $createurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Post' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose
}




##########update device id#############
$updateurl=$updateurl+$deviceid
#$body='{"objectSchemaKey":"$objectSchemaKey", "objectTypeId": $objectTypeId,"attributes": [{"objectTypeAttributeId": 342,"objectAttributeValues": [{"value": "$compinfo.Name"}]}]}'

#$nowuser = Invoke-RestMethod -Uri $updateurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Get' -ContentType 'application/json' -Verbose
$userkeykey=$userkey.key

$object = Invoke-RestMethod -Uri $updateurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Get' -ContentType 'application/json; charset=utf-8' -Verbose
ForEach ($item in $object.attributes){
    if ($item.objectTypeAttributeId -eq 409){
    if ($user -eq $null){$user=$item.objectAttributeValues.value}
    }
}
ForEach ($item in $object.attributes){
    if ($item.objectTypeAttributeId -eq 385){
    $item.objectAttributeValues.searchValue
    #$user=$item.objectAttributeValues.value
    if ($item.objectAttributeValues.searchValue -ne $null){$userkeykey=$item.objectAttributeValues.searchValue}
    }
}



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
          "value":"'+$compinfo.Name+'"
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
      ]}
  ]
}'

Invoke-RestMethod -Uri $updateurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Put' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose
$updateurl
$body
$i=0
$c=0

$soft = wmic product get name,version /format:csv
$soft = $soft | ConvertFrom-Csv -Delimiter ',' #-Header 'name','version'
#$allurl=$allurl+'&qlQuery=objectType="Software"'
#$allsoft=Invoke-RestMethod -Uri $allurl -Headers @{Authorization=("Basic {0}" -f $base64)}

$soft



    #$allurl=$allurl+'&qlQuery=Name="'+$item.name+'"'
    #$alleqsoft=Invoke-RestMethod -Uri $allurl -Headers @{Authorization=("Basic {0}" -f $base64)}

    foreach ($item in $soft){
        $item.name = $item.name -replace '\+',''
        $item.name = $item.name -replace '\)','|'
        $item.name = $item.name -replace '\(','|'
        $allequrl=$allurl+'&qlQuery=Name="'+$item.name+'"'
        $allequrl
        $alleqsoft=Invoke-RestMethod -Uri $allequrl -Headers @{Authorization=("Basic {0}" -f $base64)}
            if($null -eq ($alleqsoft.objectEntries.name | ? { $item.name -match $_ })){
            $body='{
	            "objectSchemaKey": "'+$objectSchemaKey+'",
	            "objectTypeId": 53,
	            "attributes": [{
			        "objectTypeAttributeId": 467,
			            "objectAttributeValues": [{
				            "value": "'+$item.name+'"
			            }]
		            },
		            {
			            "objectTypeAttributeId": 558,
			            "objectAttributeValues": [{
				            "value": "'+$item.version+'"
			            }]
		            }
	            ]
            }'
            write-host('create name')
            $c=$c+1
            Invoke-RestMethod -Uri $createurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Post' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose
                }
            }

foreach ($item in $soft){
        $item.name = $item.name -replace '\+',''
        $item.name = $item.name -replace '\)','|'
        $item.name = $item.name -replace '\(','|'
        $allequrl=$allurl+'&qlQuery=Name="'+$item.name+'"'
        $alleqsoft=Invoke-RestMethod -Uri $allequrl -Headers @{Authorization=("Basic {0}" -f $base64)}

               foreach ($itm in $alleqsoft.objectEntries){
                            $itmtmp = $itm.attributes -match 558
                            #$itmtmp.objectAttributeValues.value
                    if($item.name -eq $itm.name){
                        if($null -eq ($itm.attributes.objectAttributeValues.value | ? { $item.version -match $_ })){
                        $body='{
	            "objectSchemaKey": "'+$objectSchemaKey+'",
	            "objectTypeId": 53,
	            "attributes": [{
			        "objectTypeAttributeId": 467,
			            "objectAttributeValues": [{
				            "value": "'+$item.name+'"
			            }]
		            },
		            {
			            "objectTypeAttributeId": 558,
			            "objectAttributeValues": [{
				            "value": "'+$item.version+'"
			            }]
		            }
	            ]
            }'
                    write-host('create version')
                    $c=$c+1
                    #Invoke-RestMethod -Uri $createurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Post' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose
            
            }
                    }
               
               
               }
    }





$c

    

    foreach ($item in $soft){
        $item.name = $item.name -replace '\+',''
        $item.name = $item.name -replace '\)','|'
        $item.name = $item.name -replace '\(','|'
        $allequrl=$allurl+'&qlQuery=Name="'+$item.name+'"'
        #$allequrl
        $alleqsoft=Invoke-RestMethod -Uri $allequrl -Headers @{Authorization=("Basic {0}" -f $base64)}
            foreach ($itm in $alleqsoft.objectEntries){
                if ($item.name -eq $itm.name){
                    $itmtmp = $itm.attributes -match 558
                        foreach ($it in $itmtmp.objectAttributeValues.value){
                            if ($item.version -eq $it){
                                $itm.objectKey
                                $hostsoft=$itm.objectKey+';'+$hostsoft
                            
                            
                            
                                    }
                        }
                    
                
                }


            }


        }




















$body='{
	"objectSchemaKey": "'+$objectSchemaKey+'",
	"attributes": [{
		"objectTypeAttributeId": "'+$attributevar[15]+'",
		"objectAttributeValues": [{
				"value": "SCHINV-143"
			},
			{
				"value": "SCHINV-150"
			}
		]
	}]
}'
#

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
Invoke-RestMethod -Uri $updateurl -Headers @{Authorization=("Basic {0}" -f $base64)} -Method 'Put' -Body $body -ContentType 'application/json; charset=utf-8' -Verbose
$c
