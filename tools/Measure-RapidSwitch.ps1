param([Parameter(Mandatory=$true)][string]$Directory,[string]$AssertCase='',[switch]$AssertAll)
$ErrorActionPreference='Stop'
$trials=@(Import-Csv (Join-Path $Directory 'driver.csv'))
if(!$trials.Count -or !$trials[0].token){throw 'Tagged trial data required'}
$events=@(Get-ChildItem (Join-Path $Directory 'client-*.csv')|ForEach-Object {Import-Csv $_.FullName -Header t,client,event,shift,async,token})
$down=@{};$up=@{};$duplicates=@{}
foreach($e in $events){
 if($e.event -in @('256:70','260:70')){if($down.ContainsKey($e.token)){$duplicates[$e.token]=$true};$down[$e.token]=$e}
 if($e.event -in @('257:70','261:70')){if($up.ContainsKey($e.token)){$duplicates[$e.token]=$true};$up[$e.token]=$e}
}
$detail=foreach($t in $trials){
 $d=$down[$t.token];$u=$up[$t.token]
 [pscustomobject]@{Case=$t.case;Trial=[int]$t.trial;Token=$t.token;Expected=[int]$t.expected;Receiver=if($d){[int]$d.client}else{0};Missing=(!$d);Wrong=([bool]$d -and $d.client -ne $t.expected);MissingShift=([bool]$d -and $d.shift -ne '1');SystemKey=([bool]$d -and $d.event -eq '260:70');MissingUp=(!$u);Split=([bool]$d -and [bool]$u -and $d.client -ne $u.client)}
}
$summary=@($detail|Group-Object Case|ForEach-Object {
 [pscustomobject]@{Case=$_.Name;Trials=$_.Count;WrongReceiver=@($_.Group|Where-Object Wrong).Count;MissingKeyDown=@($_.Group|Where-Object Missing).Count;MissingShift=@($_.Group|Where-Object MissingShift).Count;SystemKeyDown=@($_.Group|Where-Object SystemKey).Count;MissingKeyUp=@($_.Group|Where-Object MissingUp).Count;SplitKeyPair=@($_.Group|Where-Object Split).Count}
})
$summary|Export-Csv (Join-Path $Directory 'summary.csv') -NoTypeInformation
$detail|Export-Csv (Join-Path $Directory 'delivery.csv') -NoTypeInformation
$summary|Format-Table
if($AssertAll){
 $clients=@(Import-Csv (Join-Path $Directory 'windows.csv') -Header client,hwnd).Count
 if($clients -lt 12){throw 'At least twelve clients are required'}
 $required=@('keep1-held-short','keep1-held-long','keep1-held-30ms','keep1-released','keep0-held-short')
 foreach($case in $required){if(@($detail|Where-Object Case -eq $case).Count -ne $clients*4){throw "Incomplete case: $case"}}
 if(@($trials.token|Select-Object -Unique).Count -ne $trials.Count){throw 'Duplicate trial tokens'}
 $manual=@(Import-Csv (Join-Path $Directory 'manual.csv'))
 foreach($keep in @('0','1')){if(@($manual|Where-Object keep -eq $keep).Count -ne $clients*2){throw 'Incomplete manual selection checks'}}
 if(@($manual|Where-Object {$_.actual -ne $_.expected}).Count){throw 'Manual selection violated group position behavior'}
 $selected=@($detail|Where-Object {$_.Case -in $required})
} elseif($AssertCase){
 $selected=@($detail|Where-Object Case -eq $AssertCase)
}
if($AssertAll -or $AssertCase){
 if(!$selected.Count){throw 'No matching case'}
 # WM_SYSKEYDOWN can be delivered to the active window while keyboard focus
 # settles. Count it separately, but judge routing by recipient and Shift state;
 # this harness cannot assert how EVE consumes either Windows message type.
 $bad=@($selected|Where-Object {$_.Wrong -or $_.Missing -or $_.MissingShift -or $_.MissingUp -or $_.Split -or $duplicates.ContainsKey($_.Token)})
 if($bad.Count){Write-Output "FAIL: $($bad.Count)/$($selected.Count) incorrect deliveries";exit 1}
 Write-Output "PASS: $($selected.Count) correct deliveries in $AssertCase"
}
