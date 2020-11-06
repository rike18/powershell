#Создание сессии PS
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://100-MAIL04/powershell
$k = Import-PSSession $session -DisableNameChecking -AllowClobber

$back = "3"

#Дата
[datetime]$rundate = Get-Date
#День
#$startdate = $rundate.AddDays(-$back)
#Месяц
$startdate = $rundate.AddMonths(-$back)

#Выходной файл
$outfile_date = $rundate.ToShortDateString()
$dloutfile = "DistributionStats_" + "$outfile_date" + ".csv"

#Список серверов
$mbx_servers = Get-ExchangeServer | Where { $_.serverrole -match "Mailbox" } | Foreach { $_.fqdn }

#Временный массив 
#$mbx_servers = @('100-MAIL04', '100-MAIL05')
  
$dl = @{} 
$dl_date = @{}

#Статус бар
function time_pipeline { 
    param ($increment = 1000) 
    begin { $i = 0; $timer = [diagnostics.stopwatch]::startnew() } 
    process { 
        $i++ 
        if (!($i % $increment)) { Write-host `rProcessed $i in $($timer.elapsed.totalseconds) seconds -nonewline } 
        $_ 
    } 
    end { 
        write-host `rProcessed $i log records in $($timer.elapsed.totalseconds) seconds 
        Write-Host "   Average rate: $([int]($i/$timer.elapsed.totalseconds)) log recs/sec." 
    } 
} 

foreach ($mail in $mbx_servers) { 
 
    Write-Host "`nStarted processing $mail" 
 
    get-messagetrackinglog -Server $mail -Start "$startdate" -End "$rundate" -resultsize unlimited | 
    time_pipeline | % {
        if ($_.eventid -eq "expand") { 
            $dl[($_.relatedrecipientaddress).Trim()] ++
            $curent_date = [datetime]$_.Timestamp
            if (($dl_date[$_.relatedrecipientaddress]) -gt ($curent_date)) {
            } else {
                $dl_date[$_.relatedrecipientaddress] = $curent_date
            }
        } 
    }        
} 

If (Test-Path $dloutfile) { 
    $dl_stats = Import-Csv $dloutfile 
    $dl_list = $dl_stats | Foreach { $_.address } 
}
else { 
    $dl_list = @() 
    $dl_stats = @() 
} 
 
$dl_stats | Foreach { 
    if ($dl[$_.address]) { 
        if ($_.lastused -lt $dl_date.($_.address)) {  
            $_.lastused = $dl_date.($_.address)
        } 
    } 
} 

$dl.keys | Foreach { 
    if ($dl_list -notcontains $_) { 
        $new_rec = "" | select Address, Used, LastUsed 
        $new_rec.address = $_ 
        $new_rec.used = $dl[$_] 
        $new_rec.lastused = $dl_date.$_ 
        $dl_stats += @($new_rec) 
    } 
} 
 
$dl_stats | Export-Csv $dloutfile -NoTypeInformation -Append -Encoding utf8BOM

Write-Host "Статистика $dloutfile готова" -ForegroundColor Green