# Дашборд ОПУ - интерактивный HTML

# ── Конфигурация: секреты из переменных окружения (облако, GitHub Actions)
#    или из локальных файлов (рядом со скриптом либо в папке ОПУ_Автоматизация).
#    В публичный репозиторий секреты НЕ попадают - только этот код-«движок». ──
function Get-OpuSecret($envName, $fileName) {
    $ev = [Environment]::GetEnvironmentVariable($envName)
    if (-not [string]::IsNullOrWhiteSpace($ev)) { return $ev.Trim() }
    $paths = @()
    if ($PSScriptRoot) { $paths += (Join-Path $PSScriptRoot $fileName) }
    $paths += (Join-Path 'C:\Users\sandievde\Documents\ОПУ_Автоматизация' $fileName)
    foreach ($p in $paths) { if (Test-Path $p) { return (Get-Content $p -Raw).Trim() } }
    return $null
}

$sheetId = Get-OpuSecret 'OPU_SHEET_ID' 'opu_sheet_id.txt'
if ([string]::IsNullOrWhiteSpace($sheetId)) { throw "Не задан ID таблицы: переменная OPU_SHEET_ID или файл opu_sheet_id.txt" }
$dashPwd = Get-OpuSecret 'OPU_PASSWORD' 'opu_password.txt'
if ([string]::IsNullOrWhiteSpace($dashPwd)) { throw "Не задан пароль дашборда: переменная OPU_PASSWORD или файл opu_password.txt" }

$exportUrl  = "https://docs.google.com/spreadsheets/d/$sheetId/export?format=xlsx"
$tmpFile    = "$env:TEMP\opu_dash.xlsx"
$outputFile = "C:\Users\sandievde\Desktop\Справки ОПУ\Дашборд_ОПУ.html"
# В облаке (GitHub Actions) путь вывода задаётся переменной окружения OPU_OUTPUT
if ($env:OPU_OUTPUT) { $outputFile = $env:OPU_OUTPUT }
$utf8bom    = New-Object System.Text.UTF8Encoding $true
$today      = [datetime]::Today
$monthStart = [datetime]::new($today.Year, $today.Month, 1)

function Strip-Initials($name) {
    if (-not $name) { return $name }
    ($name -replace '\s+([А-ЯЁ]\.){1,3}$', '').Trim()
}

function Get-StatusBadge($s) {
    if (-not $s) { return "" }
    $bg,$fg = switch -Regex ($s) {
        "работ|процесс" { "#dbeafe","#1e40af"; break }
        "ожидан|провер|согласован" { "#fef3c7","#92400e"; break }
        "приостан|заморож|отложен" { "#f3e8ff","#6b21a8"; break }
        default                    { "#f1f5f9","#475569" }
    }
    return "<span style='background:$bg;color:$fg;border-radius:4px;padding:2px 8px;font-size:11px;white-space:nowrap;font-weight:600'>$([System.Net.WebUtility]::HtmlEncode($s))</span>"
}

if (-not (Test-Path (Split-Path $outputFile))) { New-Item -ItemType Directory -Path (Split-Path $outputFile) | Out-Null }
Invoke-WebRequest -Uri $exportUrl -OutFile $tmpFile -UseBasicParsing
Add-Type -Assembly System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($tmpFile)

$ssReader = New-Object System.IO.StreamReader($zip.GetEntry("xl/sharedStrings.xml").Open())
[xml]$ssXml = $ssReader.ReadToEnd(); $ssReader.Close()
$strings = @()
foreach ($si in $ssXml.sst.si) {
    if ($null -ne $si.t) { $strings += [string]$si.t }
    elseif ($null -ne $si.r) { $text=""; $si.r | ForEach-Object { $text += [string]$_.t }; $strings += $text }
    else { $strings += "" }
}

$wbReader = New-Object System.IO.StreamReader($zip.GetEntry("xl/workbook.xml").Open())
[xml]$wbXml = $wbReader.ReadToEnd(); $wbReader.Close()
$rReader  = New-Object System.IO.StreamReader($zip.GetEntry("xl/_rels/workbook.xml.rels").Open())
[xml]$rXml = $rReader.ReadToEnd(); $rReader.Close()
$ridToFile = @{}
foreach ($rel in $rXml.Relationships.Relationship) { $ridToFile[$rel.Id] = "xl/" + $rel.Target }
$rNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
$sheetFiles = @{}
foreach ($s in $wbXml.workbook.sheets.sheet) {
    $rid = $s.GetAttribute("id", $rNS)
    if ($rid -and $ridToFile.ContainsKey($rid)) { $sheetFiles[$s.name] = $ridToFile[$rid] }
}

function Get-CV($c) {
    if ($null -eq $c) { return "" }
    $t = $c.GetAttribute("t"); $v = if ($null -ne $c.v) { $c.v } else { "" }
    if ($t -eq "s") { return $strings[[int]$v] }; return $v
}
function To-Date($s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    try { return ([datetime]"1899-12-30").AddDays([double]$s) } catch { return $null }
}
function Esc($s) { ([string]$s) -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' }

# Реестр заявок
$sh1File = $sheetFiles["Реестр заявок"]
$sh1R = New-Object System.IO.StreamReader($zip.GetEntry($sh1File).Open())
[xml]$sh1 = $sh1R.ReadToEnd(); $sh1R.Close()

# Гиперссылки из rels-файла листа (строка -> URL)
$linkMap = @{}
$shName   = [System.IO.Path]::GetFileName($sh1File)
$shDir    = [System.IO.Path]::GetDirectoryName($sh1File) -replace '\\','/'
$relsPath = "$shDir/_rels/$shName.rels"
$relsEntry = $zip.GetEntry($relsPath)
if ($null -ne $relsEntry) {
    $rlRdr = New-Object System.IO.StreamReader($relsEntry.Open())
    [xml]$rlXml = $rlRdr.ReadToEnd(); $rlRdr.Close()
    $ridToUrl = @{}
    foreach ($rel in $rlXml.Relationships.Relationship) {
        if ($rel.Type -like "*hyperlink*") { $ridToUrl[$rel.Id] = $rel.Target }
    }
    foreach ($hl in $sh1.worksheet.hyperlinks.hyperlink) {
        $rid = $hl.GetAttribute("id", $rNS)
        if ($hl.ref -match "^F(\d+)$" -and $ridToUrl.ContainsKey($rid)) {
            $linkMap[[int]$matches[1]] = $ridToUrl[$rid]
        }
    }
}

$records = @()
foreach ($row in $sh1.worksheet.sheetData.row) {
    $rNum = [int]$row.r
    if ($rNum -lt 3) { continue }
    $cells = @{}
    foreach ($c in $row.c) { if ($c.r -match '^([A-Z]+)') { $cells[$matches[1]] = $c } }
    $fmt=(Get-CV $cells["A"]); $id=(Get-CV $cells["F"]); $dept=(Get-CV $cells["C"])
    $stat=(Get-CV $cells["D"]); $comm=(Get-CV $cells["H"]); $nm=(Get-CV $cells["G"]); $ex=(Get-CV $cells["K"]); $dlRaw=(Get-CV $cells["M"]); $tzRaw=(Get-CV $cells["O"])
    if (($fmt -ne "Заявка" -and $fmt -ne "Предпроект") -or $stat -eq "Отложена" -or [string]::IsNullOrWhiteSpace($id)) { continue }
    $dot=$id.IndexOf("."); $idStr=if($dot -gt 0){$id.Substring(0,$dot)}else{$id}
    $dl=To-Date $dlRaw
    $tzVal=0.0; try{if(-not [string]::IsNullOrWhiteSpace($tzRaw)){$tzVal=[double]$tzRaw}}catch{}
    $records += [PSCustomObject]@{
        ID=$idStr; Fmt=$fmt; Dept=$dept; Status=$stat; Name=$nm; Exec=$ex; Comm=$comm; DlDate=$dl; Row=$rNum; TZ=$tzVal
        DaysOverdue=if($null -ne $dl -and $dl -lt $today){[int]($today-$dl).TotalDays}else{$null}
        DaysLeft=if($null -ne $dl -and $dl -ge $today){[int]($dl-$today).TotalDays}else{$null}
    }
}

$g5=@($records | Where-Object { $_.Fmt -eq "Предпроект" })
$nonPP=@($records | Where-Object { $_.Fmt -ne "Предпроект" })
$g1=@($nonPP | Where-Object { $_.Status -eq "Заявка принята" })
$g2=@($nonPP | Where-Object { $_.Status -eq "ТЗ на разработке" -or $_.Dept -eq "ОА" })
$g3=@($nonPP | Where-Object { $_.Dept -eq "ОРПО" })
$g4=@($nonPP | Where-Object { $_.Status -eq "Тестирование" })
$inGroupRows = @($g1 + $g2 + $g3 + $g4 + $g5 | Select-Object -ExpandProperty Row)
$gu=@($records | Where-Object { $inGroupRows -notcontains $_.Row })
$overdue=@($records | Where-Object { $null -ne $_.DaysOverdue } | Sort-Object DaysOverdue -Descending)
$upcoming=@($records | Where-Object { $null -ne $_.DaysLeft -and $_.DaysLeft -le 14 } | Sort-Object DaysLeft)
$noDl=@($records | Where-Object { $null -eq $_.DlDate } | Sort-Object Dept,Name)
$byExec=@($records | Where-Object { $_.Exec -ne "" } | Group-Object Exec | Sort-Object Count -Descending | Select-Object -First 8)
$peakExecs=@($records | Where-Object { $_.Exec -ne "" -and $null -ne $_.DaysLeft -and $_.DaysLeft -le 7 } | Group-Object Exec | Sort-Object Count -Descending)
$statusGroups=@($records | Group-Object Status | Sort-Object Count -Descending)
$execOverdue = "[" + (($byExec | ForEach-Object { @($_.Group | Where-Object { $null -ne $_.DaysOverdue }).Count }) -join ",") + "]"
$execUrgent  = "[" + (($byExec | ForEach-Object { @($_.Group | Where-Object { $null -eq $_.DaysOverdue -and $null -ne $_.DaysLeft -and $_.DaysLeft -le 7 }).Count }) -join ",") + "]"
$execNormal  = "[" + (($byExec | ForEach-Object { @($_.Group | Where-Object { $null -eq $_.DaysOverdue -and ($null -eq $_.DaysLeft -or $_.DaysLeft -gt 7) }).Count }) -join ",") + "]"
$allExecs = @($records | Where-Object { $_.Exec -ne "" } | Select-Object -ExpandProperty Exec -Unique | Sort-Object)
$allDepts = @($records | Where-Object { $_.Dept -ne "" } | Select-Object -ExpandProperty Dept -Unique | Sort-Object)
$execOpts = ($allExecs | ForEach-Object { $s=Strip-Initials $_; "<option value='$(Esc $s)'>$(Esc $s)</option>" }) -join ""
$deptOpts = ($allDepts | ForEach-Object { "<option value='$(Esc $_)'>$(Esc $_)</option>" }) -join ""
$byDept=@($records | Where-Object { $_.Dept -ne "" } | Group-Object Dept | Sort-Object Count -Descending | Select-Object -First 8)
$deptLabels="[" + (($byDept | ForEach-Object { '"' + (Esc $_.Name) + '"' }) -join ",") + "]"
$deptOver="[" + (($byDept | ForEach-Object { @($_.Group | Where-Object { $null -ne $_.DaysOverdue }).Count }) -join ",") + "]"
$deptUrg ="[" + (($byDept | ForEach-Object { @($_.Group | Where-Object { $null -eq $_.DaysOverdue -and $null -ne $_.DaysLeft -and $_.DaysLeft -le 7 }).Count }) -join ",") + "]"
$deptNorm="[" + (($byDept | ForEach-Object { @($_.Group | Where-Object { $null -eq $_.DaysOverdue -and ($null -eq $_.DaysLeft -or $_.DaysLeft -gt 7) }).Count }) -join ",") + "]"
$stale=@($records | Where-Object { $null -ne $_.DaysOverdue -and $_.DaysOverdue -ge 30 })
$todayTasks=@($records | Where-Object { $null -ne $_.DaysLeft -and $_.DaysLeft -eq 0 })
$endOfMonth=[datetime]::new($today.Year,$today.Month,[datetime]::DaysInMonth($today.Year,$today.Month))
$workDaysLeft=0; $d=$today.AddDays(1)
while ($d -le $endOfMonth) { if ([int]$d.DayOfWeek -ne 0 -and [int]$d.DayOfWeek -ne 6){$workDaysLeft++}; $d=$d.AddDays(1) }
$todayTasksHtml=""
foreach ($r in $todayTasks) {
    $url=if($linkMap.ContainsKey($r.Row)){$linkMap[$r.Row]}else{""}
    $idLink=if($url -ne ""){"<a href='$(Esc $url)' target='_blank' style='color:#1e3a5f;font-weight:700'>$(Esc $r.ID)</a>"}else{"<b>$(Esc $r.ID)</b>"}
    $exSpan = if ($r.Exec -ne "") { "<span class='td-exec'>$(Esc (Strip-Initials $r.Exec))</span>" } else { "" }
    $todayTasksHtml += "<div class='td-task'>$idLink — $(Esc $r.Name)$exSpan</div>"
}
if ($todayTasksHtml -eq "") { $todayTasksHtml = "<span class='td-none'>нет задач с дедлайном сегодня</span>" }

# ── Сверка с Битриксом: ID задач из ссылок «Реестра заявок» + запрос по дедлайну/исполнителю ──
$sverkaJson = "[]"; $sverkaMsg = ""; $sverkaSrc = @()
try {
    # Токен вебхука: переменная окружения BITRIX_WEBHOOK (облако) или файл bitrix_webhook.txt
    $bxHook = Get-OpuSecret 'BITRIX_WEBHOOK' 'bitrix_webhook.txt'
    if ([string]::IsNullOrWhiteSpace($bxHook)) { throw "нет ни переменной BITRIX_WEBHOOK, ни файла bitrix_webhook.txt" }
    if ($bxHook -notmatch '/$') { $bxHook += '/' }

    foreach ($row in $sh1.worksheet.sheetData.row) {
        $rNum = [int]$row.r
        if ($rNum -lt 3) { continue }
        if (-not $linkMap.ContainsKey($rNum)) { continue }
        $u = [string]$linkMap[$rNum]
        if ($u -notmatch 'task/view/(\d+)') { continue }
        $bxId = $matches[1]
        $cells = @{}
        foreach ($c in $row.c) { if ($c.r -match '^([A-Z]+)') { $cells[$matches[1]] = $c } }
        $idTxt=(Get-CV $cells["F"]); $stat=(Get-CV $cells["D"]); $nm=(Get-CV $cells["G"]); $ex=(Get-CV $cells["K"]); $dlR=(Get-CV $cells["M"])
        $dot=$idTxt.IndexOf("."); if($dot -gt 0){$idTxt=$idTxt.Substring(0,$dot)}
        $sverkaSrc += [PSCustomObject]@{ BxId=$bxId; Reg=$idTxt; Name=$nm; RegStatus=$stat; RegExec=$ex; RegDl=(To-Date $dlR); Url=$u }
    }

    if ($sverkaSrc.Count -gt 0) {
        $ids = @($sverkaSrc | Select-Object -ExpandProperty BxId -Unique | Sort-Object)
        $bx = @{}
        for ($i=0; $i -lt $ids.Count; $i += 40) {
            $chunk = @($ids[$i..([Math]::Min($i+39, $ids.Count-1))])
            $qs = "?"; foreach ($id in $chunk) { $qs += "filter[ID][]=$id&" }
            $qs += "select[]=ID&select[]=TITLE&select[]=STATUS&select[]=DEADLINE&select[]=RESPONSIBLE_ID"
            $resp = Invoke-RestMethod -Uri ($bxHook + "tasks.task.list.json" + $qs) -Method Get -TimeoutSec 60
            foreach ($t in $resp.result.tasks) {
                $bx[[string]$t.id] = [PSCustomObject]@{
                    Status   = [string]$t.status
                    Deadline = if ($t.deadline) { [datetime]$t.deadline } else { $null }
                    Resp     = if ($t.responsible) { [string]$t.responsible.name } else { "" }
                }
            }
            Start-Sleep -Milliseconds 250
        }

        $sverkaOut = @()
        foreach ($s in $sverkaSrc) {
            $b = $bx[$s.BxId]
            $found = ($null -ne $b)
            $bxDl = if ($found -and $b.Deadline) { $b.Deadline } else { $null }
            $dlDiff = ($found -and $null -ne $s.RegDl -and $null -ne $bxDl -and $s.RegDl.Date -ne $bxDl.Date)
            $regParts = @((Strip-Initials $s.RegExec) -split '\s+' | Where-Object { $_ })
            $regSur = if ($regParts.Count) { [string]$regParts[0] } else { "" }
            $bxParts = if ($found) { @([string]$b.Resp -split '\s+' | Where-Object { $_ }) } else { @() }
            $bxSur = if ($bxParts.Count) { [string]$bxParts[0] } else { "" }
            $execDiff = ($found -and $regSur -ne "" -and $bxSur -ne "" -and $regSur.ToLower() -ne $bxSur.ToLower())
            $sverkaOut += [PSCustomObject]@{
                bxId=$s.BxId; reg=$s.Reg; name=$s.Name; url=$s.Url
                regDl=if($s.RegDl){$s.RegDl.ToString('dd.MM.yyyy')}else{""}
                bxDl=if($bxDl){$bxDl.ToString('dd.MM.yyyy')}else{""}
                regExec=$regSur; bxExec=$bxSur
                found=$found; dlDiff=$dlDiff; execDiff=$execDiff
            }
        }
        $sverkaJson = @($sverkaOut) | ConvertTo-Json -Depth 4 -Compress
        if ($null -eq $sverkaJson -or $sverkaJson -eq "") { $sverkaJson = "[]" }
        $sverkaJson = $sverkaJson -replace '</','<\/'
    }
} catch {
    $sverkaMsg = "Не удалось получить данные Битрикса: $($_.Exception.Message)"
    $sverkaJson = "[]"
}
Write-Host "Сверка Битрикс: заявок со ссылкой $(@($sverkaSrc).Count)$(if($sverkaMsg){' (ошибка: '+$sverkaMsg+')'})"

# Реестр проектов
$shPR=New-Object System.IO.StreamReader($zip.GetEntry($sheetFiles["Реестр проектов"]).Open())
[xml]$shProj=$shPR.ReadToEnd(); $shPR.Close()
$projects=@(); $cPN=""; $cPID=""; $cTasks=[System.Collections.Generic.List[object]]::new()
$doneSt=@("Сделано","Выполнено","Выполнена","Завершён","Завершено")
foreach ($row in ($shProj.worksheet.sheetData.row | Where-Object { [int]$_.r -ge 3 })) {
    $cells=@{}; foreach ($c in $row.c) { if ($c.r -match '^([A-Z]+)') { $cells[$matches[1]]=$c } }
    $a=(Get-CV $cells["A"]); $bv=(Get-CV $cells["B"]); $st=(Get-CV $cells["C"]); $updRawT=(Get-CV $cells["D"]); $tn=(Get-CV $cells["F"]); $ex=(Get-CV $cells["G"]); $stRawT=(Get-CV $cells["H"]); $dlRawT=(Get-CV $cells["I"]); $tzRawT=(Get-CV $cells["K"])
    # Новый проект - когда название в A СМЕНИЛОСЬ (а не просто «A непустой»).
    # Так парсинг работает и со старой структурой (A только в первой строке проекта),
    # и с новой (A продублирован в каждой строке задачи).
    if ($a -ne "" -and $a -ne $cPN) {
        if ($cPN -ne "" -and $cTasks.Count -gt 0) {
            $dn=@($cTasks | Where-Object { $doneSt -contains $_.S }).Count
            $projects += [PSCustomObject]@{ Name=$cPN; ID=$cPID; Total=$cTasks.Count; Done=$dn; Tasks=$cTasks.ToArray() }
        }
        $cPN=$a; $dot=$bv.IndexOf("."); $cPID=if($dot -gt 0){$bv.Substring(0,$dot)}else{$bv}
        $cTasks=[System.Collections.Generic.List[object]]::new()
    }
    if (($tn -ne "" -or $st -ne "") -and $cPN -ne "") { $cTasks.Add([PSCustomObject]@{S=$st; N=$tn; Ex=$ex; StartRaw=$stRawT; DlRaw=$dlRawT; UpdRaw=$updRawT; TZ=$tzRawT}) }
}
if ($cPN -ne "" -and $cTasks.Count -gt 0) {
    $dn=@($cTasks | Where-Object { $doneSt -contains $_.S }).Count
    $projects += [PSCustomObject]@{ Name=$cPN; ID=$cPID; Total=$cTasks.Count; Done=$dn; Tasks=$cTasks.ToArray() }
}

# Напоминалка: окно дат с учётом выходных
$remDates = [System.Collections.Generic.List[PSCustomObject]]::new()
$dow = [int]$today.DayOfWeek   # 0=Вс,1=Пн,...,5=Пт,6=Сб
if ($dow -eq 1) {              # Понедельник: добавляем пропущенные Сб и Вс
    $remDates.Add([PSCustomObject]@{ Date=$today.AddDays(-2); Type="missed" })
    $remDates.Add([PSCustomObject]@{ Date=$today.AddDays(-1); Type="missed" })
}
$remDates.Add([PSCustomObject]@{ Date=$today; Type="today" })
if ($dow -eq 5) {              # Пятница: Сб, Вс (назначить заранее) + Пн
    $remDates.Add([PSCustomObject]@{ Date=$today.AddDays(1); Type="weekend" })
    $remDates.Add([PSCustomObject]@{ Date=$today.AddDays(2); Type="weekend" })
    $remDates.Add([PSCustomObject]@{ Date=$today.AddDays(3); Type="upcoming" })
} elseif ($dow -ne 6 -and $dow -ne 0) {   # Остальные рабочие дни: следующие 3 рабочих
    $nd = $today.AddDays(1); $wc = 0
    while ($wc -lt 3) {
        $ndow = [int]$nd.DayOfWeek
        if ($ndow -ne 0 -and $ndow -ne 6) {
            $remDates.Add([PSCustomObject]@{ Date=$nd; Type="upcoming" }); $wc++
        }
        $nd = $nd.AddDays(1)
    }
}
$ruCult = [System.Globalization.CultureInfo]::GetCultureInfo("ru-RU")
$remCount = 0; $remHtml = ""
foreach ($rd in $remDates) {
    $dateRows = ""; $dateCount = 0
    foreach ($p in $projects) {
        foreach ($t in $p.Tasks) {
            if ($doneSt -contains $t.S) { continue }
            $tStart = To-Date $t.StartRaw
            if ($null -ne $tStart -and $tStart.Date -eq $rd.Date.Date) {
                $exFmt = if ($t.Ex -ne "") { Strip-Initials $t.Ex } else { "—" }
                $dlFmt = if ($null -ne (To-Date $t.DlRaw)) { (To-Date $t.DlRaw).ToString("dd.MM") } else { "—" }
                $rkey = ($rd.Date.ToString("yyyyMMdd") + "_" + $p.Name + "_" + $t.N) -replace '[^a-zA-Z0-9а-яА-ЯёЁ]','_' -replace '_{2,}','_'
                if ($rkey.Length -gt 80) { $rkey = $rkey.Substring(0,80) }
                if ($rd.Type -eq "today" -or $rd.Type -eq "weekend") {
                    $chkCell = "<td style='width:36px;text-align:center'><input type='checkbox' class='rem-chk' onchange='remCheck(this)'></td>"
                } else {
                    $chkCell = "<td style='width:36px'></td>"
                }
                $stBadge = Get-StatusBadge $t.S
                $dateRows += "<tr class='rem-row' data-rkey='$(Esc $rkey)'><td>$(Esc $p.Name)</td><td>$(Esc $t.N)</td><td>$exFmt</td><td>$dlFmt</td><td>$stBadge</td>$chkCell</tr>"
                $dateCount++
                if ($rd.Type -eq "today") { $remCount++ }
            }
        }
    }
    if ($dateCount -eq 0) { continue }
    $rdow = [int]$rd.Date.DayOfWeek
    $hdrLabel = switch ($rd.Type) {
        "missed"   { "Пропущено" }
        "today"    { "Сегодня" }
        "weekend"  { if ($rdow -eq 6) { "Суббота - назначить заранее" } else { "Воскресенье - назначить заранее" } }
        "upcoming" { "Скоро" }
    }
    $dateFmt = $rd.Date.ToString("dd MMMM", $ruCult) + " (" + $rd.Date.ToString("ddd", $ruCult) + ")"
    $remHtml += "<div class='rem-grp rem-$($rd.Type)'><div class='rem-grp-hdr'><span class='rem-grp-lbl'>$hdrLabel</span><span class='rem-grp-date'>$dateFmt</span></div><table><thead><tr><th>Проект</th><th>Этап</th><th>Ответственный</th><th>Срок</th><th>Статус</th><th style='width:36px'></th></tr></thead><tbody>$dateRows</tbody></table></div>"
}
if ($remHtml -eq "") { $remHtml = "<div class='empty' style='padding:20px;text-align:center;color:#94a3b8'>Нет предстоящих назначений</div>" }

# Архив
$shAR=New-Object System.IO.StreamReader($zip.GetEntry($sheetFiles["Архив"]).Open())
[xml]$shArch=$shAR.ReadToEnd(); $shAR.Close()
$zip.Dispose()
$doneThisMonth=0
foreach ($row in $shArch.worksheet.sheetData.row) {
    if ([int]$row.r -lt 3) { continue }
    $cells=@{}; foreach ($c in $row.c) { if ($c.r -match '^([A-Z]+)') { $cells[$matches[1]]=$c } }
    $dd=To-Date (Get-CV $cells["K"])
    if ($null -ne $dd -and $dd -ge $monthStart) { $doneThisMonth++ }
}

# ── Расчёт выполнимости к дедлайнам (Этап 1: модель загрузки) ────────────────
$planDays = 14
$planEnd  = $today.AddDays($planDays)
$availWorkDays = 0
$d = $today.AddDays(1)
while ($d -le $planEnd) { if ([int]$d.DayOfWeek -ne 0 -and [int]$d.DayOfWeek -ne 6) { $availWorkDays++ }; $d = $d.AddDays(1) }

$execTZ = @{}   # имя -> суммарные трудозатраты
$execTC = @{}   # имя -> число задач с TZ > 0
$execTA = @{}   # имя -> всего задач в окне

# Из реестра заявок
foreach ($r in $records) {
    if ($r.Exec -eq "" -or $null -eq $r.DlDate) { continue }
    if ($r.DlDate -ge $today -and $r.DlDate -le $planEnd) {
        $nm = Strip-Initials $r.Exec
        if (-not $execTZ.ContainsKey($nm)) { $execTZ[$nm]=0.0; $execTC[$nm]=0; $execTA[$nm]=0 }
        $execTA[$nm]++
        if ($r.TZ -gt 0) { $execTZ[$nm]+=$r.TZ; $execTC[$nm]++ }
    }
}
# Из реестра проектов
$excludeStProj = @("Сделано","Выполнено","Выполнена","Завершён","Завершено","Отложено","Приостановлено")
foreach ($p in $projects) {
    foreach ($t in $p.Tasks) {
        if ($excludeStProj -contains $t.S -or $t.Ex -eq "") { continue }
        $tDl = To-Date $t.DlRaw
        if ($null -eq $tDl -or $tDl -lt $today -or $tDl -gt $planEnd) { continue }
        $nm = Strip-Initials $t.Ex
        if (-not $execTZ.ContainsKey($nm)) { $execTZ[$nm]=0.0; $execTC[$nm]=0; $execTA[$nm]=0 }
        $execTA[$nm]++
        $tTz=0.0; try{if(-not [string]::IsNullOrWhiteSpace($t.TZ)){$tTz=[double]$t.TZ}}catch{}
        if ($tTz -gt 0) { $execTZ[$nm]+=$tTz; $execTC[$nm]++ }
    }
}

# Строим HTML-таблицу
$loadRows = ""
if ($execTZ.Count -gt 0) {
    $sortedExecs = $execTZ.Keys | Sort-Object { if ($availWorkDays -gt 0) { $execTZ[$_]/$availWorkDays } else { 0 } } -Descending
    foreach ($nm in $sortedExecs) {
        $req   = $execTZ[$nm]
        $total = $execTA[$nm]
        $withTZ = $execTC[$nm]
        $coeff = if ($availWorkDays -gt 0 -and $req -gt 0) { [math]::Round($req/$availWorkDays,2) } else { $null }
        $coeffCell = if ($null -ne $coeff) {
            $cls = if ($coeff -gt 1.0) { "red" } else { "green" }
            $lbl = if ($coeff -gt 1.0) { "Не успевает" } else { "Успевает" }
            "<td class='$cls' style='font-weight:600'>$coeff</td><td class='$cls'>$lbl</td>"
        } else {
            "<td style='color:#94a3b8'>—</td><td style='color:#94a3b8'>нет данных</td>"
        }
        $tzDisplay = if ($req -gt 0) { [math]::Round($req,1) } else { "—" }
        $covNote = if ($withTZ -lt $total) { "<span style='color:#94a3b8;font-size:11px'> ($withTZ из $total)</span>" } else { "" }
        $loadRows += "<tr><td>$(Esc $nm)</td><td>$tzDisplay$covNote</td><td>$availWorkDays</td>$coeffCell</tr>"
    }
} else {
    $loadRows = "<tr><td colspan='5' class='empty'>Нет задач с дедлайном в ближайшие $planDays дней</td></tr>"
}
$loadHtml = "<table><thead><tr><th>Исполнитель</th><th>Трудозатраты (дн.)</th><th>Рабочих дней</th><th>Коэффициент</th><th>Статус</th></tr></thead><tbody>$loadRows</tbody></table>"
$loadHtml += "<div style='font-size:11px;color:#94a3b8;margin-top:8px'>Окно: ближайшие $planDays дн. ($availWorkDays рабочих). К&gt;1.0 = объём задач превышает доступное время. В скобках — доля задач с заполненными трудозатратами.</div>"

# ── Данные для JS сценарного движка ──────────────────────────────────────────
$sdTasks = [System.Collections.Generic.List[object]]::new()
foreach ($r in $records) {
    if ($r.Exec -eq "" -or $null -eq $r.DlDate) { continue }
    if ($r.DlDate -lt $today.AddDays(-7)) { continue }
    $sdTasks.Add([PSCustomObject]@{ id=$r.ID; name=$r.Name; exec=(Strip-Initials $r.Exec); dl=$r.DlDate.ToString("yyyy-MM-dd"); tz=$r.TZ; src="req"; proj="" })
}
foreach ($p in $projects) {
    foreach ($t in $p.Tasks) {
        if ($excludeStProj -contains $t.S) { continue }
        $tDl = To-Date $t.DlRaw
        if ($null -eq $tDl -or $tDl -lt $today.AddDays(-7)) { continue }
        $nm2 = Strip-Initials $t.Ex; if ($nm2 -eq "") { continue }
        $tTz2=0.0; try{if(-not [string]::IsNullOrWhiteSpace($t.TZ)){$tTz2=[double]$t.TZ}}catch{}
        $sdTasks.Add([PSCustomObject]@{ id=$p.ID; name=$t.N; exec=$nm2; dl=$tDl.ToString("yyyy-MM-dd"); tz=$tTz2; src="proj"; proj=$p.Name })
    }
}
$sdExecList = @($sdTasks | Select-Object -ExpandProperty exec -Unique | Where-Object { $_ -ne "" } | Sort-Object)
$sdProjList = @($sdTasks | Select-Object -ExpandProperty proj -Unique | Where-Object { $_ -ne "" } | Sort-Object)
$sdObj = [PSCustomObject]@{ avail=$availWorkDays; planDays=$planDays; tasks=$sdTasks.ToArray(); execs=$sdExecList; projs=$sdProjList }
$sdJson = ConvertTo-Json $sdObj -Depth 3 -Compress

# ── Контекст для ИИ-ассистента ───────────────────────────────────────────────
$aiLines = [System.Collections.Generic.List[string]]::new()
# Топ-20 заявок по срочности: просроченные вперед, затем ближайший дедлайн
$urgentRecs = $records | Sort-Object {
    if ($null -ne $_.DaysOverdue) { -$_.DaysOverdue }
    elseif ($null -ne $_.DaysLeft) { $_.DaysLeft + 1000 }
    else { 9999 }
} | Select-Object -First 20
$aiLines.Add("РЕЕСТР ЗАЯВОК (активных: $($records.Count), показано топ-$($urgentRecs.Count) срочных):")
foreach ($r in $urgentRecs) {
    $dl = if ($null -ne $r.DaysOverdue) { "ПРОСРОЧКА $($r.DaysOverdue) дн." }
          elseif ($null -ne $r.DaysLeft) { "через $($r.DaysLeft) дн. (до $($r.DlDate.ToString('dd.MM')))" }
          else { "дедлайн не указан" }
    $aiLines.Add("[$($r.ID)] $($r.Name) | $($r.Status) | $(Strip-Initials $r.Exec) | $($r.Dept) | $dl")
}
$aiLines.Add("")
$aiLines.Add("ПРОГРЕСС ПРОЕКТОВ:")
foreach ($p in $projects) {
    $pct = if ($p.Total -gt 0) { [int]($p.Done/$p.Total*100) } else { 0 }
    $aiLines.Add("[$($p.ID)] $($p.Name): $($p.Done)/$($p.Total) этапов ($pct%)")
    # Только активные (незавершенные) задачи
    foreach ($t in ($p.Tasks | Where-Object { $doneSt -notcontains $_.S })) {
        $tDlP = To-Date $t.DlRaw
        $dlInfo = if ($null -ne $tDlP) { $tDlP.ToString("dd.MM") } else { "---" }
        $aiLines.Add("  $($t.S)|$($t.N)|$(Strip-Initials $t.Ex)|$dlInfo")
    }
}
# Жесткий лимит 7000 символов (~3500 токенов для русского текста)
$aiText = $aiLines -join "`n"
if ($aiText.Length -gt 7000) { $aiText = $aiText.Substring(0, 7000) + "...(обрезано)" }
$aiContextJson = ConvertTo-Json $aiText

# ── Аналитика ────────────────────────────────────────────────────────────────
# п.1 — Пульс системы
$noExecCount=@($records | Where-Object { $_.Exec -eq "" }).Count
$healthScore=[Math]::Max(0,100-[int]($overdue.Count*100/[Math]::Max($records.Count,1)*0.5)-[Math]::Min($stale.Count*5,20)-[Math]::Min($noExecCount*2,10)-[Math]::Min($noDl.Count,10))
$healthColor=if($healthScore -ge 80){"#16a34a"}elseif($healthScore -ge 60){"#d97706"}else{"#dc2626"}
$healthBg   =if($healthScore -ge 80){"#dcfce7"}elseif($healthScore -ge 60){"#fef3c7"}else{"#fee2e2"}
$healthLabel=if($healthScore -ge 80){"Хорошо"}elseif($healthScore -ge 60){"Внимание"}else{"Критично"}

# ── HTML-функции для групп ────────────────────────────────────────────────────
$vaultEncoded   = [Uri]::EscapeDataString("PMD")
$obsNotesRoot   = "C:\Knowledge\PMD\08_заявки"
$obsidianNoteMap = @{}
if (Test-Path $obsNotesRoot) {
    Get-ChildItem "$obsNotesRoot\*.md" | ForEach-Object {
        if ($_.BaseName -match '^Заявка_(\d+)_') {
            $obsidianNoteMap[$Matches[1]] = "08_заявки/" + $_.BaseName
        }
    }
}
function Find-ObsidianNote($id) {
    if (-not $id -or $obsidianNoteMap.Count -eq 0) { return $null }
    if ($obsidianNoteMap.ContainsKey($id)) { return $obsidianNoteMap[$id] }
    return $null
}
function Format-TaskHtml($r, $showDept) {
    $url = if ($linkMap.ContainsKey($r.Row)) { $linkMap[$r.Row] } else { "" }
    if ($url -ne "") {
        $idHtml = "<a class='t-id' href='$(Esc $url)' target='_blank'>$(Esc $r.ID)</a>"
    } else {
        $idHtml = "<span class='t-id'>$(Esc $r.ID)</span>"
    }
    $deptHtml = if ($showDept -and $r.Dept -ne "") { "<span class='t-dept'>$(Esc $r.Dept)</span>" } else { "" }
    $nameHtml = "<span class='t-name'>$(Esc $r.Name)</span>"
    $exHtml   = if ($r.Exec -ne "") { "<span class='t-exec' onclick='filterByExec(this)' title='Фильтр по исполнителю'>$(Esc (Strip-Initials $r.Exec))</span>" } else { "" }
    $dlHtml = ""; $dlCls = ""
    if ($null -ne $r.DaysOverdue) {
        $dlHtml = "просрочка $($r.DaysOverdue) дн."; $dlCls = "dl-red"
    } elseif ($null -ne $r.DaysLeft) {
        $dlStr = $r.DlDate.ToString("dd.MM")
        if ($r.DaysLeft -eq 0)     { $dlHtml = "дедлайн сегодня"; $dlCls = "dl-red" }
        elseif ($r.DaysLeft -eq 1) { $dlHtml = "дедлайн завтра";  $dlCls = "dl-orange" }
        elseif ($r.DaysLeft -le 7) { $dlHtml = "до $dlStr ($($r.DaysLeft) дн.)"; $dlCls = "dl-orange" }
        else                       { $dlHtml = "до $dlStr"; $dlCls = "dl-ok" }
    }
    $dlSpan = if ($dlHtml -ne "") { "<span class='t-dl $dlCls'>$dlHtml</span>" } else { "" }
    $commHtml = ""
    if ($r.Comm -ne "") {
        $commHtml = "<span class='t-comm'>$(Esc $r.Comm)</span>"
    }
    $dn = ($r.Name -replace "'","&#39;"); $de = ((Strip-Initials $r.Exec) -replace "'","&#39;"); $dd = ($r.Dept -replace "'","&#39;")
    $dc = ((Esc $r.Comm) -replace "'","&#39;")
    $notePath = Find-ObsidianNote $r.ID
    if ($notePath) {
        $eparts  = ($notePath -split '/') | ForEach-Object { [Uri]::EscapeDataString($_) }
        $obsUrl  = "obsidian://open?vault=$vaultEncoded&file=" + ($eparts -join '/')
        $obsBtnHtml = "<a class='obs-btn no-print' href='$(Esc $obsUrl)' title='Открыть заметку в Obsidian'><svg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 24 24' fill='currentColor'><polygon points='12,2 2,9 12,22 22,9'/></svg>Заметка</a>"
    } else { $obsBtnHtml = "" }
    $watchBtn = "<button class='watch-btn no-print' onclick='toggleWatch(this)' data-id='$(Esc $r.ID)' title='Добавить в наблюдение'>&#9734;</button>"
    $aiAskBtn = "<button class='ai-ask-btn no-print' onclick='askAIAbout(this)' title='Спросить ИИ об этой заявке'><svg xmlns='http://www.w3.org/2000/svg' width='11' height='11' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='11' width='18' height='10' rx='2'/><circle cx='12' cy='5' r='2'/><path d='M12 7v4'/><circle cx='8' cy='16' r='1.5' fill='currentColor' stroke='none'/><circle cx='16' cy='16' r='1.5' fill='currentColor' stroke='none'/></svg></button>"
    return "<div class='task' data-name='$(Esc $dn)' data-id='$(Esc $r.ID)' data-exec='$(Esc $de)' data-dept='$(Esc $dd)' data-comm='$dc'>$idHtml$deptHtml$nameHtml$exHtml$dlSpan$commHtml$obsBtnHtml$watchBtn$aiAskBtn</div>"
}

function Build-GroupHtml($title, $accentCls, $tasks, $showDept) {
    $cnt = $tasks.Count
    if ($cnt -eq 0) {
        $body = "<div class='task-empty'>нет заявок</div>"
    } else {
        $body = ($tasks | ForEach-Object { Format-TaskHtml $_ $showDept }) -join ""
    }
    $copyBtn = "<button class='grp-copy no-print' onclick='copyGroup(this)' title='Скопировать список'><svg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='9' y='9' width='13' height='13' rx='2'/><path d='M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1'/></svg></button>"
    return "<div class='grp'><div class='grp-hdr $accentCls'><span>$title</span><div style='display:flex;align-items:center;gap:6px'>$copyBtn<span class='grp-cnt'>$cnt</span></div></div><div class='grp-body'>$body</div></div>"
}

$g5html = Build-GroupHtml "0. Предпроект"            "ga-slate"  $g5 $true
$g1html = Build-GroupHtml "1. Заявка принята"       "ga-blue"   $g1 $true
$g2html = Build-GroupHtml "2. Разработка ТЗ / ОА"  "ga-purple" $g2 $true
$g3html = Build-GroupHtml "3. В работе в ОРПО"      "ga-amber"  $g3 $false
$g4html = Build-GroupHtml "4. Тестирование"         "ga-teal"   $g4 $false

$guHtml = Build-GroupHtml "Прочие / Уточнение" "ga-gray" $gu $true

# ── Тепловая карта дедлайнов ─────────────────────────────────────────────────
$dlMap = @{}
foreach ($r in $records) {
    if ($null -ne $r.DlDate) {
        $key = $r.DlDate.ToString("yyyy-MM-dd")
        if (-not $dlMap.ContainsKey($key)) { $dlMap[$key] = [System.Collections.Generic.List[string]]::new() }
        $short = if ($r.Name.Length -gt 30) { $r.Name.Substring(0,27)+"..." } else { $r.Name }
        $dlMap[$key].Add("$($r.ID) - $short")
    }
}
foreach ($p in $projects) {
    foreach ($t in $p.Tasks) {
        $tDl = To-Date $t.DlRaw
        if ($null -ne $tDl) {
            $key = $tDl.ToString("yyyy-MM-dd")
            if (-not $dlMap.ContainsKey($key)) { $dlMap[$key] = [System.Collections.Generic.List[string]]::new() }
            $tn = $t.N; if ($tn.Length -gt 25) { $tn = $tn.Substring(0,22)+"..." }
            $dlMap[$key].Add("[пр] $($p.Name): $tn")
        }
    }
}
$ruMon   = @("","янв","фев","мар","апр","май","июн","июл","авг","сен","окт","ноя","дек")
$dowOff  = if ([int]$today.DayOfWeek -eq 0) { -6 } else { 1 - [int]$today.DayOfWeek }
$calStart = $today.AddDays($dowOff)
$calEnd   = $calStart.AddDays(62)
$calHtml  = "<div class='cal-hdr'><span>Пн</span><span>Вт</span><span>Ср</span><span>Чт</span><span>Пт</span><span>Сб</span><span>Вс</span></div><div class='cal-grid'>"
$calCur   = $calStart
while ($calCur -le $calEnd) {
    $key     = $calCur.ToString("yyyy-MM-dd")
    $cnt     = if ($dlMap.ContainsKey($key)) { $dlMap[$key].Count } else { 0 }
    $isToday = $calCur -eq $today
    $isPast  = $calCur -lt $today
    $isFar   = $calCur -gt $today.AddDays(60)
    $cls = if ($isToday) { "cal-today" } elseif ($isPast -or $isFar) { "cal-muted" } elseif ($cnt -eq 0) { "cal-zero" } elseif ($cnt -le 2) { "cal-low" } elseif ($cnt -le 4) { "cal-mid" } else { "cal-high" }
    $tipAttr = ""
    if ($cnt -gt 0 -and -not $isPast) {
        $tip = ($dlMap[$key] -join "&#10;") -replace "'","&#39;"
        $tipAttr = " title='$tip'"
    }
    $isWknd  = ([int]$calCur.DayOfWeek -eq 0 -or [int]$calCur.DayOfWeek -eq 6)
    $wkndCls = if ($isWknd) { " cal-wknd" } else { "" }
    $monSpan = if ($calCur.Day -eq 1) { "<span class='cal-mon'>$($ruMon[$calCur.Month])</span>" } else { "" }
    $cntSpan = if ($cnt -gt 0 -and -not $isPast) { "<span class='cal-cnt'>$cnt</span>" } else { "" }
    $calHtml += "<div class='cal-day $cls$wkndCls'$tipAttr>$monSpan<span class='cal-d'>$($calCur.Day)</span>$cntSpan</div>"
    $calCur = $calCur.AddDays(1)
}
$calHtml += "</div><div class='cal-legend'><span class='cl-i cl-low'>1-2 задачи</span><span class='cl-i cl-mid'>3-4 задачи</span><span class='cl-i cl-high'>5+ задач</span><span class='cl-i cl-td'>сегодня</span></div>"

# ── Матрица рисков ────────────────────────────────────────────────────────────
$riskRed = @(); $riskOrange = @(); $riskBlue = @()
$riskSrc = @(@{y=1;rs=$g1},@{y=2;rs=$g2},@{y=3;rs=$g3},@{y=4;rs=$g4})
foreach ($grp in $riskSrc) {
    $gy = $grp.y
    foreach ($r in @($grp.rs)) {
        $gx = if ($null -ne $r.DaysLeft) { $r.DaysLeft } elseif ($null -ne $r.DaysOverdue) { -$r.DaysOverdue } else { $null }
        if ($null -eq $gx) { continue }
        $gnm = $r.Name; if ($gnm.Length -gt 28) { $gnm = $gnm.Substring(0,25)+"..." }
        $gnm = $gnm -replace '"','\"'
        $pt  = "{""x"":$gx,""y"":$gy,""id"":""$(Esc $r.ID)"",""nm"":""$gnm"",""ex"":""$(Esc $r.Exec)""}"
        if ($null -ne $r.DaysOverdue)                                     { $riskRed    += $pt }
        elseif ($null -ne $r.DaysLeft -and $r.DaysLeft -le 7) { $riskOrange += $pt }
        else                                                               { $riskBlue   += $pt }
    }
}
$riskRedJson    = "[" + ($riskRed    -join ",") + "]"
$riskOrangeJson = "[" + ($riskOrange -join ",") + "]"
$riskBlueJson   = "[" + ($riskBlue   -join ",") + "]"

# ── Остальные блоки ───────────────────────────────────────────────────────────
$overdueRows=""
foreach ($r in $overdue) {
    $url = if ($linkMap.ContainsKey($r.Row)) { $linkMap[$r.Row] } else { "" }
    $idCell = if ($url -ne "") { "<a href='$(Esc $url)' target='_blank' style='color:#1e3a5f;font-weight:700'>$(Esc $r.ID)</a>" } else { "<b>$(Esc $r.ID)</b>" }
    $staleCls = if ($r.DaysOverdue -ge 30) { " class='tr-stale'" } else { "" }
    $overdueRows += "<tr$staleCls><td>$idCell</td><td>$(Esc $r.Name)</td><td>$(Esc (Strip-Initials $r.Exec))</td><td class='red' data-v='$($r.DaysOverdue)'>$($r.DaysOverdue) дн.</td></tr>"
}
if ($overdueRows -eq "") { $overdueRows='<tr><td colspan="4" class="empty">нет просроченных</td></tr>' }

$upcomingRows=""
foreach ($r in $upcoming) {
    $dl=if($null -ne $r.DlDate){$r.DlDate.ToString("dd.MM")}else{"-"}
    $cls=if($r.DaysLeft -le 2){"red"}elseif($r.DaysLeft -le 7){"orange"}else{"green"}
    $lft=if($r.DaysLeft -eq 0){"сегодня"}elseif($r.DaysLeft -eq 1){"завтра"}else{"$($r.DaysLeft) дн."}
    $url = if ($linkMap.ContainsKey($r.Row)) { $linkMap[$r.Row] } else { "" }
    $idCell = if ($url -ne "") { "<a href='$(Esc $url)' target='_blank' style='color:#1e3a5f;font-weight:700'>$(Esc $r.ID)</a>" } else { "<b>$(Esc $r.ID)</b>" }
    $nameCell = "$(Esc $r.Name)"
    if ($r.Comm -ne "") { $nameCell += "<br><span class='t-comm' style='padding-left:0'>$(Esc $r.Comm)</span>" }
    $dv = if ($null -ne $r.DaysLeft) { $r.DaysLeft } else { 999 }
    $upcomingRows += "<tr><td>$idCell</td><td>$nameCell</td><td>$(Esc (Strip-Initials $r.Exec))</td><td>$dl</td><td class='$cls' data-v='$dv'>$lft</td></tr>"
}
if ($upcomingRows -eq "") { $upcomingRows='<tr><td colspan="5" class="empty">нет дедлайнов в ближайшие 14 дней</td></tr>' }

$noDlRows=""
foreach ($r in $noDl) {
    $url = if ($linkMap.ContainsKey($r.Row)) { $linkMap[$r.Row] } else { "" }
    $idCell = if ($url -ne "") { "<a href='$(Esc $url)' target='_blank' style='color:#1e3a5f;font-weight:700'>$(Esc $r.ID)</a>" } else { "<b>$(Esc $r.ID)</b>" }
    $noDlRows += "<tr><td>$idCell</td><td>$(Esc $r.Name)</td><td>$(Esc (Strip-Initials $r.Exec))</td><td>$(Esc $r.Dept)</td></tr>"
}
if ($noDlRows -eq "") { $noDlRows='<tr><td colspan="4" class="empty">все задачи имеют дедлайн</td></tr>' }

$ganttEntries=@()
$projBars=""
foreach ($p in $projects) {
    $pct=if($p.Total -gt 0){[int]($p.Done/$p.Total*100)}else{0}
    $clr=if($pct -ge 75){"#16a34a"}elseif($pct -ge 40){"#2563eb"}else{"#d97706"}
    $taskRows=""
    $gTasks=@()
    foreach ($t in $p.Tasks) {
        $isDone = $doneSt -contains $t.S
        $stCls  = if ($isDone) { "ts-done" } else { "ts-open" }
        $stLbl  = if ($t.S -ne "") { Esc $t.S } else { "без статуса" }
        $tName  = if ($t.N -ne "") { Esc $t.N } else { "(без названия)" }
        $exHtml = if ($t.Ex -ne "") { "<span class='t-exec' onclick='filterByExec(this)' title='Фильтр по исполнителю'>$(Esc (Strip-Initials $t.Ex))</span>" } else { "" }
        $dlHtml = ""; $dlCls = ""
        $tDl = To-Date $t.DlRaw
        if ($isDone) {
            if ($null -ne $tDl) {
                $tUpd = To-Date $t.UpdRaw
                if ($null -ne $tUpd) {
                    $dDiff = [int]($tUpd - $tDl).TotalDays
                    if ($dDiff -gt 0) { $dlHtml = "просрочка $dDiff дн."; $dlCls = "dl-red" }
                    else              { $dlHtml = "в срок";                $dlCls = "dl-done" }
                } else {
                    $dlHtml = $tDl.ToString("dd.MM"); $dlCls = "dl-ok"
                }
            }
        } elseif ($null -ne $tDl) {
            $dLeft = [int]($tDl - $today).TotalDays
            if     ($dLeft -lt 0) { $dlHtml = "просрочка $(-$dLeft) дн.";               $dlCls = "dl-red" }
            elseif ($dLeft -eq 0) { $dlHtml = "сегодня";                                $dlCls = "dl-red" }
            elseif ($dLeft -eq 1) { $dlHtml = "завтра";                                 $dlCls = "dl-orange" }
            elseif ($dLeft -le 7) { $dlHtml = "$($tDl.ToString('dd.MM')) ($dLeft дн.)"; $dlCls = "dl-orange" }
            else                  { $dlHtml = "$($tDl.ToString('dd.MM')) ($dLeft дн.)"; $dlCls = "dl-ok" }
        }
        $dlSpan = if ($dlHtml -ne "") { "<span class='t-dl $dlCls'>$dlHtml</span>" } else { "" }
        $staleHtml = ""
        if (-not $isDone) {
            $tUpd2 = To-Date $t.UpdRaw
            if ($null -ne $tUpd2) {
                $staleDays = [int]($today - $tUpd2).TotalDays
                if     ($staleDays -ge 30) { $staleHtml = "<span class='t-stale st-r' title='Без движения $staleDays дн.'>$staleDays дн.</span>" }
                elseif ($staleDays -ge 15) { $staleHtml = "<span class='t-stale st-o' title='Без движения $staleDays дн.'>$staleDays дн.</span>" }
                elseif ($staleDays -ge 8)  { $staleHtml = "<span class='t-stale st-y' title='Без движения $staleDays дн.'>$staleDays дн.</span>" }
            }
        }
        $taskRows += "<div class='trow'><span class='ts $stCls'>$stLbl</span><span class='tname'>$tName</span>$exHtml$dlSpan$staleHtml</div>"
        $gDl = if ($null -ne $tDl) { $tDl.ToString("yyyy-MM-dd") } else { "" }
        $gN = $t.N -replace '\\','\\\\' -replace '"','\"'
        $gTasks += '{"n":"' + $gN + '","dl":"' + $gDl + '","done":' + $(if($isDone){"true"}else{"false"}) + '}'
    }
    $gId=$p.ID -replace '"','\"'; $gNm=$p.Name -replace '"','\"'
    $ganttEntries += '"' + $gId + '":{"name":"' + $gNm + '","tasks":[' + ($gTasks -join ',') + ']}'
    $pOvd=@($p.Tasks | Where-Object { -not ($doneSt -contains $_.S) -and $null -ne (To-Date $_.DlRaw) -and (To-Date $_.DlRaw) -lt $today }).Count
    $pHealth=[Math]::Max(0,100-[int]((1-$pct/100)*40)-[Math]::Min($pOvd*15,40))
    $phCls=if($pHealth -ge 75){"ph-g"}elseif($pHealth -ge 50){"ph-a"}else{"ph-r"}
    $phStatus=if($pHealth -ge 75){'Норма'}elseif($pHealth -ge 50){'Внимание'}else{'Критично'}
    $phTip="Индекс здоровья: $pHealth / 100 — $phStatus. Складывается из: степени выполнения задач ($pct% выполнено) и числа просроченных этапов ($pOvd шт.). Снижается при низком прогрессе и наличии просрочек."
    $projBars += "<details class='proj'><summary class='proj-hdr'><span><span class='ph-dot $phCls' title='$phTip'></span><b>$(Esc $p.ID)</b> - $(Esc $p.Name)</span><div style='display:flex;align-items:center;gap:8px'><button class='btn-gantt no-print' onclick='event.preventDefault();openGantt(this)' title='Диаграмма Ганта'>Гант</button><span class='proj-pct'>$($p.Done)/$($p.Total) ($pct%)</span></div></summary><div class='track' style='margin:6px 0 8px'><div class='bar' style='width:${pct}%;background:$clr'></div></div><div class='tlist'>$taskRows</div></details>"
}

$ganttJson = "{" + ($ganttEntries -join ",") + "}"
$execLabels="[" + (($byExec | ForEach-Object { '"' + (Esc (Strip-Initials $_.Name)) + '"' }) -join ",") + "]"
$execData="[" + (($byExec | ForEach-Object { $_.Count }) -join ",") + "]"
$stLabels="[" + (($statusGroups | ForEach-Object { '"' + (Esc $_.Name) + '"' }) -join ",") + "]"
$stData="[" + (($statusGroups | ForEach-Object { $_.Count }) -join ",") + "]"

$dateStr=$today.ToString("dd.MM.yyyy")
$monthName=(Get-Culture).DateTimeFormat.GetMonthName($today.Month)
$sb=New-Object System.Text.StringBuilder

$genTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$genTimeDisplay = Get-Date -Format "dd.MM.yyyy HH:mm"
$sb.AppendLine("<!DOCTYPE html><!-- Сгенерировано: $genTime --><html lang=""ru""><head><meta charset=""UTF-8""><meta name=""viewport"" content=""width=device-width,initial-scale=1""><title>Дашборд ОПУ</title>") | Out-Null
$sb.AppendLine('<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>') | Out-Null
$sb.AppendLine('<style>') | Out-Null
$sb.AppendLine('*{box-sizing:border-box;margin:0;padding:0}body{font-family:"Segoe UI",Arial,sans-serif;background:#eef2f7;color:#1e293b;padding:20px;font-size:14px}') | Out-Null
$sb.AppendLine('h1{font-size:22px;font-weight:700;color:#1e3a5f}h2{font-size:12px;font-weight:700;color:#475569;margin-bottom:14px;text-transform:uppercase;letter-spacing:.6px}') | Out-Null
$sb.AppendLine('.hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;padding:16px 20px;background:#fff;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.07)}') | Out-Null
$sb.AppendLine('.hdr-sub{font-size:12px;color:#64748b;margin-top:3px}.date-badge{background:#1e3a5f;color:#fff;padding:6px 16px;border-radius:6px;font-size:13px;font-weight:600}') | Out-Null
$sb.AppendLine('.hdr-right{display:flex;align-items:center;gap:12px}.btn-refresh{display:inline-flex;align-items:center;justify-content:center;color:#64748b;border:1px solid #cbd5e1;border-radius:6px;padding:6px 8px;text-decoration:none;background:#f8fafc;line-height:1}.btn-refresh:hover{background:#f1f5f9;color:#1e3a5f;border-color:#94a3b8}') | Out-Null
$sb.AppendLine('.btn-ar{display:inline-flex;align-items:center;gap:5px;color:#64748b;border:1px solid #cbd5e1;border-radius:6px;padding:5px 8px;background:#f8fafc;cursor:pointer;font-size:12px;font-weight:600;line-height:1}.btn-ar:hover{background:#f1f5f9;color:#1e3a5f}.btn-ar.ar-active{background:#eff6ff;border-color:#93c5fd;color:#1d4ed8}') | Out-Null
$sb.AppendLine('.grp-copy{background:none;border:none;cursor:pointer;opacity:.4;padding:3px;border-radius:4px;display:inline-flex;align-items:center;color:inherit;line-height:1}.grp-copy:hover{opacity:1;background:rgba(255,255,255,.25)}') | Out-Null
$sb.AppendLine('.obs-btn{display:inline-flex;align-items:center;gap:3px;padding:2px 7px;background:#7c3aed;color:#fff!important;border-radius:4px;font-size:10px;font-weight:600;text-decoration:none!important;white-space:nowrap;vertical-align:middle;margin-left:6px;line-height:1.6}.obs-btn:hover{background:#6d28d9}') | Out-Null
$sb.AppendLine('.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px}') | Out-Null
$sb.AppendLine('.card{background:#fff;border-radius:10px;padding:18px 20px;box-shadow:0 1px 4px rgba(0,0,0,.07);border-left:4px solid #2563eb}') | Out-Null
$sb.AppendLine('.card.red{border-left-color:#dc2626}.card.green{border-left-color:#16a34a}.card.amber{border-left-color:#d97706}') | Out-Null
$sb.AppendLine('.num{font-size:38px;font-weight:700;line-height:1}.lbl{font-size:12px;color:#64748b;margin-top:5px}') | Out-Null
$sb.AppendLine('.row{display:grid;gap:14px;margin-bottom:18px}.r2{grid-template-columns:1fr 1fr}.r32{grid-template-columns:3fr 2fr}') | Out-Null
$sb.AppendLine('.panel{background:#fff;border-radius:10px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,.07);position:relative}') | Out-Null
$sb.AppendLine('.ch{position:relative;height:220px}.ch-h{position:relative;height:200px}') | Out-Null
$sb.AppendLine('details.proj{margin-bottom:8px;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden}details.proj:last-child{margin-bottom:0}') | Out-Null
$sb.AppendLine('details.proj summary{list-style:none;cursor:pointer;background:#f8fafc}details.proj summary::-webkit-details-marker{display:none}') | Out-Null
$sb.AppendLine('details.proj summary:hover{background:#f1f5f9}details[open].proj summary{background:#f1f5f9;border-bottom:1px solid #e2e8f0}') | Out-Null
$sb.AppendLine('.proj-hdr{display:flex;align-items:center;justify-content:space-between;font-size:13px;gap:8px;padding:10px 14px;user-select:none}') | Out-Null
$sb.AppendLine('.proj-hdr::before{content:"▶";font-size:10px;color:#94a3b8;margin-right:2px;flex-shrink:0;transition:transform .15s}details[open]>.proj-hdr::before{transform:rotate(90deg)}') | Out-Null
$sb.AppendLine('.proj-pct{color:#64748b;white-space:nowrap;font-size:12px}.track{height:6px;background:#e2e8f0;border-radius:0;overflow:hidden}.bar{height:100%}') | Out-Null
$sb.AppendLine('.tlist{padding:4px 0 6px 14px;background:#fff}.trow{display:flex;align-items:baseline;gap:8px;padding:4px 14px 4px 0;font-size:12px;border-bottom:1px solid #f1f5f9}') | Out-Null
$sb.AppendLine('.trow:last-child{border-bottom:none}.ts{flex-shrink:0;border-radius:4px;padding:1px 7px;font-size:11px;font-weight:500;min-width:80px;text-align:center}') | Out-Null
$sb.AppendLine('.ts-done{background:#dcfce7;color:#166534}.ts-open{background:#f1f5f9;color:#475569}.tname{color:#334155;flex:1}') | Out-Null
$sb.AppendLine('table{width:100%;border-collapse:collapse;font-size:13px}th{text-align:left;padding:8px 10px;background:#f8fafc;color:#475569;font-weight:600;border-bottom:2px solid #e2e8f0}') | Out-Null
$sb.AppendLine('td{padding:7px 10px;border-bottom:1px solid #f1f5f9;vertical-align:top}tr:last-child td{border-bottom:none}tr:hover td{background:#f8fafc}') | Out-Null
$sb.AppendLine('.red{color:#dc2626;font-weight:600}.orange{color:#d97706;font-weight:600}.green{color:#16a34a;font-weight:600}.empty{text-align:center;color:#94a3b8;padding:20px!important}') | Out-Null
$sb.AppendLine('.grps{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px}') | Out-Null
$sb.AppendLine('.grp{border-radius:8px;overflow:hidden;border:1px solid #e2e8f0}') | Out-Null
$sb.AppendLine('.grp-hdr{display:flex;justify-content:space-between;align-items:center;padding:8px 14px;font-weight:600;font-size:13px}') | Out-Null
$sb.AppendLine('.grp-cnt{background:rgba(255,255,255,0.35);border-radius:20px;padding:1px 9px;font-size:12px;font-weight:700}') | Out-Null
$sb.AppendLine('.ga-blue{background:#1d4ed8;color:#eff6ff}.ga-purple{background:#7c3aed;color:#f5f3ff}.ga-amber{background:#d97706;color:#fffbeb}.ga-teal{background:#0f766e;color:#f0fdfa}.ga-slate{background:#475569;color:#f8fafc}.ga-gray{background:#64748b;color:#f1f5f9}') | Out-Null
$sb.AppendLine('.grp-body{padding:4px 0}') | Out-Null
$sb.AppendLine('.task{display:flex;align-items:baseline;flex-wrap:wrap;gap:4px 6px;padding:5px 12px;border-bottom:1px solid #f1f5f9;font-size:12.5px}') | Out-Null
$sb.AppendLine('.task:last-child{border-bottom:none}.task:hover{background:#f8fafc}') | Out-Null
$sb.AppendLine('.t-id{font-weight:700;color:#1e3a5f;min-width:32px;flex-shrink:0;text-decoration:none}') | Out-Null
$sb.AppendLine('a.t-id:hover{color:#2563eb;text-decoration:underline}') | Out-Null
$sb.AppendLine('.t-dept{background:#e0e7ff;color:#3730a3;border-radius:4px;padding:1px 6px;font-size:11px;flex-shrink:0}') | Out-Null
$sb.AppendLine('.t-name{flex:1;min-width:160px;color:#1e293b}') | Out-Null
$sb.AppendLine('.t-exec{color:#64748b;font-size:11.5px;flex-shrink:0}') | Out-Null
$sb.AppendLine('.t-comm{color:#94a3b8;font-size:11px;font-style:italic;flex-basis:100%;padding-left:38px}') | Out-Null
$sb.AppendLine('.t-dl{font-size:11px;flex-shrink:0;border-radius:4px;padding:1px 6px}') | Out-Null
$sb.AppendLine('.dl-red{background:#fee2e2;color:#b91c1c;font-weight:600}.dl-orange{background:#fef3c7;color:#92400e;font-weight:600}.dl-ok{color:#64748b}.dl-done{background:#dcfce7;color:#15803d;font-weight:600}') | Out-Null
$sb.AppendLine('.task-empty{padding:10px 14px;color:#94a3b8;font-size:12px;font-style:italic}') | Out-Null
$sb.AppendLine('.grp-uточ{margin-top:8px;border-top:1px dashed #e2e8f0;padding-top:8px}.gu-lbl{display:block;font-size:11px;font-weight:600;color:#64748b;padding:2px 12px 6px;text-transform:uppercase;letter-spacing:.5px}') | Out-Null
$sb.AppendLine('.filter-bar{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap}') | Out-Null
$sb.AppendLine('.filter-bar input{flex:1;min-width:160px;padding:6px 10px;border:1px solid #e2e8f0;border-radius:6px;font-size:12.5px;outline:none}.filter-bar input:focus{border-color:#2563eb}') | Out-Null
$sb.AppendLine('.filter-bar select{padding:6px 8px;border:1px solid #e2e8f0;border-radius:6px;font-size:12.5px;background:#fff;outline:none;cursor:pointer}.filter-bar select:focus{border-color:#2563eb}') | Out-Null
$sb.AppendLine('.btn-reset{padding:6px 12px;border:1px solid #e2e8f0;border-radius:6px;font-size:12px;background:#f8fafc;cursor:pointer;color:#64748b;white-space:nowrap}.btn-reset:hover{background:#f1f5f9;color:#1e3a5f}') | Out-Null
$sb.AppendLine('@media(max-width:900px){.cards{grid-template-columns:1fr 1fr}.r2,.r32,.grps{grid-template-columns:1fr}}') | Out-Null
$sb.AppendLine('.ch-xl{position:relative;height:260px}') | Out-Null
$sb.AppendLine('.cal-hdr{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;margin-bottom:3px;text-align:center;font-size:10px;color:#94a3b8;font-weight:600;padding:0 1px}') | Out-Null
$sb.AppendLine('.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px}') | Out-Null
$sb.AppendLine('.cal-day{border-radius:5px;padding:3px 2px;text-align:center;font-size:11px;cursor:default;min-height:34px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:1px}') | Out-Null
$sb.AppendLine('.cal-muted{background:#f8fafc;color:#cbd5e1}.cal-zero{background:#f1f5f9;color:#94a3b8}.cal-today{outline:2px solid #2563eb;background:#eff6ff;color:#1e3a5f;font-weight:700}') | Out-Null
$sb.AppendLine('.cal-low{background:#dcfce7;color:#166534}.cal-mid{background:#fef3c7;color:#92400e}.cal-high{background:#fee2e2;color:#b91c1c;font-weight:700}') | Out-Null
$sb.AppendLine('.cal-d{font-size:12px;line-height:1}.cal-cnt{font-size:9px;font-weight:700;line-height:1;opacity:.85}.cal-mon{font-size:8px;font-weight:600;text-transform:uppercase;letter-spacing:.3px;opacity:.75;line-height:1}') | Out-Null
$sb.AppendLine('.cal-wknd{opacity:.35}') | Out-Null
$sb.AppendLine('.cal-legend{display:flex;gap:14px;margin-top:8px;flex-wrap:wrap;font-size:11px;color:#64748b}') | Out-Null
$sb.AppendLine('.cl-i{display:flex;align-items:center;gap:5px}.cl-i::before{content:"";display:inline-block;width:12px;height:12px;border-radius:3px;flex-shrink:0}') | Out-Null
$sb.AppendLine('.cl-low::before{background:#dcfce7}.cl-mid::before{background:#fef3c7}.cl-high::before{background:#fee2e2}.cl-td::before{background:#eff6ff;outline:2px solid #2563eb}') | Out-Null
$sb.AppendLine('th.srt{cursor:pointer;user-select:none}th.srt:hover{background:#eef2f7}th.srt::after{content:" ⇅";font-size:9px;opacity:.4}th.srt-a::after{content:" ▲";opacity:1}th.srt-d::after{content:" ▼";opacity:1}') | Out-Null
$sb.AppendLine('details.nodl{margin-top:14px;border:1px solid #fef3c7;border-radius:8px;overflow:hidden}') | Out-Null
$sb.AppendLine('details.nodl summary{list-style:none;cursor:pointer;padding:10px 16px;background:#fffbeb;font-size:13px;font-weight:600;color:#92400e;display:flex;align-items:center;gap:8px}') | Out-Null
$sb.AppendLine('details.nodl summary::-webkit-details-marker{display:none}details.nodl summary::before{content:"▶";font-size:10px;transition:transform .15s}details[open].nodl summary::before{transform:rotate(90deg)}') | Out-Null
$sb.AppendLine('details.nodl summary:hover{background:#fef9c3}details[open].nodl summary{border-bottom:1px solid #fef3c7}') | Out-Null
$sb.AppendLine('.btn-print{display:inline-flex;align-items:center;justify-content:center;color:#64748b;border:1px solid #cbd5e1;border-radius:6px;padding:6px 8px;background:#f8fafc;cursor:pointer;line-height:1}.btn-print:hover{background:#f1f5f9;color:#1e3a5f}') | Out-Null
$sb.AppendLine('@media print{.hdr-right,.filter-bar,.no-print{display:none!important}body{background:#fff;padding:8px}.panel{box-shadow:none;border:1px solid #e2e8f0;break-inside:avoid;margin-bottom:10px}.ch{height:160px}details.proj{break-inside:avoid}details.nodl{break-inside:avoid}}') | Out-Null
$sb.AppendLine('.tab-nav{display:flex;gap:6px;margin-bottom:18px}') | Out-Null
$sb.AppendLine('.tab-btn{padding:9px 20px;border:none;border-radius:8px;background:#e2e8f0;color:#475569;font-size:13px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:6px;transition:background .15s}') | Out-Null
$sb.AppendLine('.tab-btn.active{background:#1e3a5f;color:#fff}') | Out-Null
$sb.AppendLine('.tab-btn:not(.active):hover{background:#cbd5e1}') | Out-Null
$sb.AppendLine('.btn-urgent{padding:9px 16px;border:none;border-radius:8px;background:#e2e8f0;color:#475569;font-size:13px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:6px;transition:background .15s;margin-left:auto}') | Out-Null
$sb.AppendLine('.btn-urgent.active{background:#dc2626;color:#fff}.btn-urgent:not(.active):hover{background:#cbd5e1}') | Out-Null
$sb.AppendLine('.stale-banner{display:none;align-items:center;gap:10px;background:#fef3c7;border:1px solid #fbbf24;border-radius:8px;padding:10px 16px;margin-bottom:14px;font-size:13px;color:#92400e}.stale-banner.on{display:flex}') | Out-Null
$sb.AppendLine('.g-search-bar{display:flex;align-items:center;gap:14px;margin-bottom:18px}') | Out-Null
$sb.AppendLine('.g-search-wrap{position:relative;display:flex;align-items:center;flex:1;max-width:520px}') | Out-Null
$sb.AppendLine('#g-srch{width:100%;padding:9px 30px 9px 34px;border:1.5px solid #e2e8f0;border-radius:8px;font-size:13.5px;outline:none;transition:border-color .15s,box-shadow .15s;font-family:inherit;background:#fff}') | Out-Null
$sb.AppendLine('#g-srch:focus{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.08)}') | Out-Null
$sb.AppendLine('.g-srch-ico{position:absolute;left:10px;color:#94a3b8;pointer-events:none;display:flex}') | Out-Null
$sb.AppendLine('.g-srch-clear{position:absolute;right:8px;border:none;background:none;color:#94a3b8;cursor:pointer;font-size:13px;padding:3px;line-height:1;display:none;border-radius:3px}.g-srch-clear.on{display:block}.g-srch-clear:hover{color:#475569;background:#f1f5f9}') | Out-Null
$sb.AppendLine('#g-srch-cnt{font-size:12px;white-space:nowrap;min-width:120px}') | Out-Null
$sb.AppendLine('.t-stale{font-size:10px;font-weight:700;border-radius:4px;padding:1px 6px;margin-left:4px;white-space:nowrap}') | Out-Null
$sb.AppendLine('.st-y{background:#fef3c7;color:#92400e}.st-o{background:#ffedd5;color:#9a3412}.st-r{background:#fee2e2;color:#b91c1c}') | Out-Null
$sb.AppendLine('.watch-btn{border:none;background:none;cursor:pointer;font-size:14px;color:#cbd5e1;padding:2px 4px;line-height:1;border-radius:4px;transition:color .15s,transform .1s;margin-left:auto}.watch-btn:hover{color:#f59e0b}.watch-btn.watching{color:#f59e0b}') | Out-Null
$sb.AppendLine('.watch-item{display:flex;align-items:center;gap:8px;padding:7px 0;border-bottom:1px solid #f1f5f9;flex-wrap:wrap}.watch-item:last-child{border-bottom:none}') | Out-Null
$sb.AppendLine('.watch-rm{border:none;background:none;color:#94a3b8;cursor:pointer;font-size:12px;padding:2px 6px;border-radius:4px;margin-left:auto}.watch-rm:hover{background:#fee2e2;color:#b91c1c}') | Out-Null
$sb.AppendLine('.ai-msgs{height:380px;overflow-y:auto;padding:4px 0;display:flex;flex-direction:column;gap:10px;margin-bottom:12px;scroll-behavior:smooth}') | Out-Null
$sb.AppendLine('.ai-msg{padding:10px 14px;border-radius:10px;font-size:13px;line-height:1.6;max-width:88%;word-break:break-word}') | Out-Null
$sb.AppendLine('.ai-msg b{font-weight:700}.ai-msg ul{margin:4px 0 4px 18px;padding:0}.ai-msg li{margin-bottom:2px}') | Out-Null
$sb.AppendLine('.ai-user{background:#1e3a5f;color:#fff;align-self:flex-end;border-bottom-right-radius:3px}') | Out-Null
$sb.AppendLine('.ai-assistant{background:#f1f5f9;color:#1e293b;align-self:flex-start;border-bottom-left-radius:3px}') | Out-Null
$sb.AppendLine('.ai-err{background:#fee2e2;color:#b91c1c;align-self:flex-start}') | Out-Null
$sb.AppendLine('.ai-typing span{display:inline-block;width:6px;height:6px;border-radius:50%;background:#94a3b8;margin:0 2px;animation:aiDot 1.2s infinite}.ai-typing span:nth-child(2){animation-delay:.2s}.ai-typing span:nth-child(3){animation-delay:.4s}') | Out-Null
$sb.AppendLine('@keyframes aiDot{0%,80%,100%{transform:scale(0.6);opacity:.4}40%{transform:scale(1);opacity:1}}') | Out-Null
$sb.AppendLine('.ai-input-row{display:flex;gap:8px}.ai-input-row input{flex:1;padding:9px 14px;border:1.5px solid #e2e8f0;border-radius:8px;font-size:13px;outline:none;font-family:inherit;transition:border-color .15s}') | Out-Null
$sb.AppendLine('.ai-input-row input:focus{border-color:#1e3a5f}.ai-send-btn{padding:9px 20px;border:none;border-radius:8px;background:#1e3a5f;color:#fff;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap}.ai-send-btn:hover{background:#1e4976}.ai-send-btn:disabled{background:#94a3b8;cursor:not-allowed}') | Out-Null
$sb.AppendLine('.ai-chips{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px}.ai-chip{padding:5px 12px;border:1px solid #e2e8f0;border-radius:20px;font-size:12px;background:#f8fafc;cursor:pointer;color:#475569;transition:background .15s}.ai-chip:hover{background:#e2e8f0;color:#1e3a5f}.ai-chip-scen{border-color:#bfdbfe;background:#eff6ff;color:#1d4ed8;font-weight:600}.ai-chip-scen:hover{background:#dbeafe;color:#1e3a5f}.ai-scen-hint{font-size:11px;color:#94a3b8;margin-bottom:10px;line-height:1.5}') | Out-Null
$sb.AppendLine('.sd-panel{background:#f8fafc;border:1.5px solid #bfdbfe;border-radius:10px;padding:14px 16px;margin-bottom:12px}.sd-panel-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;font-weight:700;font-size:13px;color:#1e3a5f}.sd-panel-hdr button{background:none;border:none;cursor:pointer;color:#94a3b8;font-size:16px;line-height:1;padding:0 2px}.sd-panel-hdr button:hover{color:#475569}') | Out-Null
$sb.AppendLine('.sd-form{display:grid;grid-template-columns:1fr 1fr;gap:10px 14px}.sd-row{display:flex;flex-direction:column;gap:3px}.sd-row label{font-size:11px;color:#64748b;font-weight:600}.sd-row select,.sd-row input[type=number],.sd-row input[type=date]{padding:7px 10px;border:1.5px solid #e2e8f0;border-radius:6px;font-size:12px;font-family:inherit;outline:none;background:#fff}.sd-row select:focus,.sd-row input:focus{border-color:#1e3a5f}.sd-calc-btn{grid-column:1/-1;padding:9px;background:#1e3a5f;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;margin-top:4px}.sd-calc-btn:hover{background:#1e4976}') | Out-Null
$sb.AppendLine('.sd-comp{margin:6px 0 10px;background:#f0f7ff;border:1px solid #bfdbfe;border-radius:8px;padding:12px}.sd-comp-lbl{font-size:12px;font-weight:700;color:#1e3a5f;margin-bottom:8px;display:flex;align-items:center;gap:6px}.sd-comp-tbl{width:100%;border-collapse:collapse;font-size:12px}.sd-comp-tbl th{padding:5px 10px;text-align:left;color:#475569;font-weight:600;background:#e0effe;border-bottom:1.5px solid #bfdbfe}.sd-comp-tbl td{padding:5px 10px;border-bottom:1px solid #dbeafe}.sd-warn{font-size:10px;color:#d97706;background:#fef3c7;border-radius:3px;padding:1px 5px;margin-left:3px}') | Out-Null
$sb.AppendLine('.plan-nav{display:flex;align-items:center;gap:10px;margin-bottom:14px;flex-wrap:wrap}.plan-nav button{padding:7px 14px;border:1.5px solid #e2e8f0;border-radius:8px;background:#fff;color:#475569;font-size:13px;font-weight:600;cursor:pointer}.plan-nav button:hover{background:#f1f5f9;border-color:#cbd5e1}.plan-range{font-size:14px;font-weight:700;color:#1e3a5f;margin-left:6px}.plan-wrap{overflow-x:auto;padding-bottom:6px}') | Out-Null
$sb.AppendLine('.plan-grid{border-collapse:collapse;width:100%;min-width:900px}.plan-grid th,.plan-grid td{border:1px solid #e2e8f0;vertical-align:top}.plan-grid thead th{background:#f8fafc;padding:6px 4px;font-size:11px;font-weight:600;color:#475569;text-align:center;white-space:nowrap}.plan-grid thead th.wknd{background:#f1f5f9;color:#94a3b8}.plan-grid thead th.today{background:#1e3a5f;color:#fff}.plan-day-num{font-size:15px;font-weight:700;display:block}') | Out-Null
$sb.AppendLine('.plan-exec-col{background:#fff;padding:8px 10px;font-size:12px;font-weight:600;color:#1e3a5f;white-space:nowrap;position:sticky;left:0;z-index:2;min-width:130px}.plan-cell{padding:3px;min-width:74px}.plan-cell.wknd{background:#fafbfc}.plan-cell.today{background:#eff6ff}') | Out-Null
$sb.AppendLine('.plan-blk{display:block;margin:2px 0;padding:3px 6px;border-radius:5px;font-size:10.5px;font-weight:600;line-height:1.25;border-left:3px solid;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:130px}.plan-blk.req{background:#eff6ff;border-color:#3b82f6;color:#1e40af}.plan-blk.proj{background:#f0fdfa;border-color:#14b8a6;color:#0f766e}.plan-blk.warn{background:#fffbeb;border-color:#f59e0b;color:#b45309}.plan-blk.over{background:#fef2f2;border-color:#ef4444;color:#b91c1c}') | Out-Null
$sb.AppendLine('.plan-load{padding:6px 10px;white-space:nowrap;min-width:130px}.plan-load-txt{font-size:11px;font-weight:600;color:#475569}.plan-load-bar{height:6px;border-radius:3px;background:#e2e8f0;overflow:hidden;margin-top:4px;width:110px}.plan-load-fill{height:100%;border-radius:3px}') | Out-Null
$sb.AppendLine('.plan-tip{position:fixed;z-index:9999;background:#0f172a;color:#f1f5f9;padding:9px 12px;border-radius:8px;font-size:12px;line-height:1.55;max-width:280px;box-shadow:0 8px 24px rgba(0,0,0,.3);pointer-events:none;display:none}.plan-tip b{color:#fff;font-size:12.5px}.plan-tip .pt-meta{color:#94a3b8;font-size:11px}') | Out-Null
$sb.AppendLine('.plan-legend{display:flex;gap:16px;flex-wrap:wrap;margin-top:14px;font-size:11px;color:#64748b}.plan-legend span{display:inline-flex;align-items:center;gap:5px}.plan-legend i{width:14px;height:12px;border-radius:3px;border-left:3px solid;display:inline-block}.plan-empty{padding:40px;text-align:center;color:#94a3b8;font-size:14px}') | Out-Null
$sb.AppendLine('.ai-setup{text-align:center;padding:30px 20px}.ai-key-row{display:flex;gap:8px;max-width:440px;margin:14px auto}.ai-key-row input{flex:1;padding:8px 12px;border:1.5px solid #e2e8f0;border-radius:8px;font-size:13px;outline:none;font-family:inherit}.ai-key-row input:focus{border-color:#1e3a5f}.ai-key-save{padding:8px 18px;border:none;border-radius:8px;background:#1e3a5f;color:#fff;font-size:13px;font-weight:600;cursor:pointer}.ai-hint{font-size:11px;color:#94a3b8;margin-top:8px}') | Out-Null
$sb.AppendLine('.ai-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}.ai-hdr h2{margin:0}.ai-hdr-btns{display:flex;gap:6px}.ai-hdr-btn{padding:5px 12px;border:1px solid #e2e8f0;border-radius:6px;font-size:11px;background:#f8fafc;cursor:pointer;color:#64748b}.ai-hdr-btn:hover{background:#f1f5f9;color:#1e3a5f}') | Out-Null
$sb.AppendLine('.ai-mode-btn.active{background:#1e3a5f!important;color:#fff!important;border-color:#1e3a5f!important}') | Out-Null
$sb.AppendLine('.ai-assistant{position:relative}.ai-copy-btn{position:absolute;right:6px;top:6px;background:none;border:none;cursor:pointer;color:#94a3b8;font-size:12px;padding:2px 5px;border-radius:4px;opacity:0;transition:opacity .15s;line-height:1}.ai-assistant:hover .ai-copy-btn{opacity:1}') | Out-Null
$sb.AppendLine('.ai-timer{font-size:11px;color:#94a3b8;margin-left:10px;font-style:italic}.ai-hist-sep{text-align:center;color:#94a3b8;font-size:11px;padding:6px 0 8px;border-bottom:1px solid #f1f5f9;margin-bottom:4px}') | Out-Null
$sb.AppendLine('.ai-ask-btn{border:none;background:none;cursor:pointer;color:#94a3b8;font-size:10px;padding:2px 5px;border-radius:4px;margin-left:4px;white-space:nowrap;vertical-align:middle}.ai-ask-btn:hover{background:#eff6ff;color:#1e3a5f}') | Out-Null
$sb.AppendLine('.t-exec{cursor:pointer;transition:color .15s}.t-exec:hover{color:#1e3a5f;text-decoration:underline}') | Out-Null
$sb.AppendLine('.t-exec.exec-active{color:#1e3a5f;font-weight:700;background:#eff6ff;border-radius:3px;padding:0 3px}') | Out-Null
$sb.AppendLine('.exec-filter-badge{display:none;align-items:center;gap:5px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:6px;padding:3px 8px;font-size:12px;color:#1e3a5f;margin-left:8px;white-space:nowrap}') | Out-Null
$sb.AppendLine('.exec-filter-badge button{border:none;background:none;cursor:pointer;color:#64748b;font-size:12px;padding:0 2px;line-height:1}.exec-filter-badge button:hover{color:#b91c1c}') | Out-Null
$sb.AppendLine('.rem-badge{background:#ef4444;color:#fff;border-radius:10px;padding:1px 7px;font-size:11px;font-weight:700;display:none}') | Out-Null
$sb.AppendLine('.rem-badge.on{display:inline}') | Out-Null
$sb.AppendLine('.rem-date{font-size:13px;color:#64748b;margin-bottom:16px}') | Out-Null
$sb.AppendLine('.rem-row{transition:opacity .25s}.rem-chk{width:16px;height:16px;cursor:pointer;accent-color:#1e3a5f}') | Out-Null
$sb.AppendLine('.rem-grp{margin-bottom:20px}') | Out-Null
$sb.AppendLine('.rem-grp-hdr{display:flex;align-items:center;gap:12px;margin-bottom:10px;padding:7px 12px;border-radius:7px;font-size:13px;font-weight:600}') | Out-Null
$sb.AppendLine('.rem-grp-date{opacity:.75;font-weight:400;font-size:12px}') | Out-Null
$sb.AppendLine('.rem-grp.rem-missed .rem-grp-hdr{background:#fee2e2;color:#b91c1c}') | Out-Null
$sb.AppendLine('.rem-grp.rem-today .rem-grp-hdr{background:#dbeafe;color:#1e40af}') | Out-Null
$sb.AppendLine('.rem-grp.rem-weekend .rem-grp-hdr{background:#fef3c7;color:#92400e}') | Out-Null
$sb.AppendLine('.rem-grp.rem-upcoming .rem-grp-hdr{background:#f1f5f9;color:#475569}') | Out-Null
$sb.AppendLine('.tr-stale td{background:#fff0f0!important}.tr-stale .red{color:#991b1b}') | Out-Null
$sb.AppendLine('.pulse{display:inline-flex;align-items:center;gap:5px;margin-top:5px;font-size:11px;font-weight:600;padding:3px 9px;border-radius:20px}') | Out-Null
$sb.AppendLine('.ph-dot{display:inline-block;width:9px;height:9px;border-radius:50%;flex-shrink:0;margin-right:5px;vertical-align:middle}.ph-g{background:#16a34a}.ph-a{background:#d97706}.ph-r{background:#dc2626}') | Out-Null
$sb.AppendLine('.ch-sm{position:relative;height:140px}.r-analytics{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:18px}') | Out-Null
$sb.AppendLine('@media(max-width:900px){.r-analytics{grid-template-columns:1fr}}') | Out-Null
$sb.AppendLine('.btn-copy{padding:5px 12px;border:1px solid #fca5a5;border-radius:6px;font-size:11px;background:#fff5f5;cursor:pointer;color:#991b1b;white-space:nowrap}.btn-copy:hover{background:#fee2e2}') | Out-Null
$sb.AppendLine('.stale-hint{font-size:11px;color:#991b1b;margin-top:2px}') | Out-Null
$sb.AppendLine('.peak-row{display:flex;align-items:center;gap:8px;padding:7px 0;border-bottom:1px solid #f1f5f9;cursor:default}.peak-row:last-child{border-bottom:none}') | Out-Null
$sb.AppendLine('.peak-bar{width:3px;height:34px;border-radius:2px;flex-shrink:0}.peak-name{flex:1;font-size:13px;color:#334155;font-weight:500}') | Out-Null
$sb.AppendLine('.peak-badge{font-size:13px;font-weight:700;min-width:22px;text-align:center}.peak-lbl{font-size:11px;font-weight:600;white-space:nowrap}') | Out-Null
$sb.AppendLine('.peak-empty{color:#94a3b8;font-size:12px;font-style:italic;padding:24px 0;text-align:center}') | Out-Null
$sb.AppendLine('.p-hide{position:absolute;top:10px;right:12px;background:none;border:none;font-size:18px;color:#e2e8f0;cursor:pointer;line-height:1;padding:0 3px;border-radius:4px}.p-hide:hover{color:#dc2626;background:#fee2e2}') | Out-Null
$sb.AppendLine('.btn-gantt{font-size:11px;font-weight:700;padding:3px 9px;border:none;border-radius:5px;background:rgba(255,255,255,.92);color:#1e3a5f;cursor:pointer;letter-spacing:.2px}.btn-gantt:hover{background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.18)}') | Out-Null
$sb.AppendLine('.gantt-box{background:#fff;border-radius:12px;width:92vw;max-width:960px;display:flex;flex-direction:column;box-shadow:0 8px 40px rgba(0,0,0,.25)}') | Out-Null
$sb.AppendLine('.gantt-hdr{display:flex;align-items:center;justify-content:space-between;padding:14px 20px;border-bottom:1px solid #e2e8f0;flex-shrink:0}') | Out-Null
$sb.AppendLine('.gantt-ttl{font-size:14px;font-weight:700;color:#1e3a5f}.gantt-close{background:none;border:none;font-size:24px;cursor:pointer;color:#94a3b8;line-height:1;padding:0 4px}.gantt-close:hover{color:#dc2626}') | Out-Null
$sb.AppendLine('.gantt-body{overflow:auto;padding:16px 20px;max-height:72vh}') | Out-Null
$sb.AppendLine('.g-dates{position:relative;height:38px;margin-left:200px;border-bottom:2px solid #e2e8f0;margin-bottom:6px}') | Out-Null
$sb.AppendLine('.g-lbl{position:absolute;font-size:10px;color:#94a3b8;transform:translateX(-50%);white-space:nowrap;bottom:4px}') | Out-Null
$sb.AppendLine('.g-today-line{position:absolute;top:0;bottom:0;width:2px;background:#ef4444;opacity:.5;transform:translateX(-50%);pointer-events:none;z-index:2}') | Out-Null
$sb.AppendLine('.g-today-lbl{position:absolute;top:2px;left:4px;font-size:9px;color:#ef4444;font-weight:700;white-space:nowrap}') | Out-Null
$sb.AppendLine('.g-rows{display:flex;flex-direction:column;position:relative}') | Out-Null
$sb.AppendLine('.g-row{display:flex;align-items:center;min-height:34px;border-bottom:1px solid #f8fafc}.g-row:last-child{border-bottom:none}') | Out-Null
$sb.AppendLine('.g-nm{width:200px;min-width:200px;font-size:12px;color:#334155;padding-right:14px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;cursor:default}') | Out-Null
$sb.AppendLine('.g-area{flex:1;position:relative;height:28px}') | Out-Null
$sb.AppendLine('.g-line{position:absolute;top:50%;height:14px;transform:translateY(-50%);border-radius:4px;left:0}') | Out-Null
$sb.AppendLine('.g-tick{position:absolute;bottom:0;width:1px;height:6px;background:#d1d5db;transform:translateX(-50%)}') | Out-Null
$sb.AppendLine('.g-dlbl{position:absolute;top:50%;transform:translateY(-50%);padding-left:4px;font-size:10px;white-space:nowrap;font-weight:600;background:transparent!important}') | Out-Null
$sb.AppendLine('.gd-done{background:#16a34a}.g-line.gd-done{background:#bbf7d0}.g-dlbl.gd-done{color:#16a34a}') | Out-Null
$sb.AppendLine('.gd-over{background:#dc2626}.g-line.gd-over{background:#fecaca}.g-dlbl.gd-over{color:#dc2626}') | Out-Null
$sb.AppendLine('.gd-urg{background:#d97706}.g-line.gd-urg{background:#fde68a}.g-dlbl.gd-urg{color:#d97706}') | Out-Null
$sb.AppendLine('.gd-ok{background:#2563eb}.g-line.gd-ok{background:#bfdbfe}.g-dlbl.gd-ok{color:#2563eb}') | Out-Null
$sb.AppendLine('.g-nodl{font-size:11px;color:#94a3b8;font-style:italic;position:absolute;top:50%;transform:translateY(-50%)}') | Out-Null
$sb.AppendLine('.g-empty{padding:30px;text-align:center;color:#94a3b8;font-style:italic}') | Out-Null
$sb.AppendLine('#restore-dock{display:none;position:fixed;bottom:20px;right:20px;z-index:500}') | Out-Null
$sb.AppendLine('#rd-list{display:none;position:absolute;bottom:calc(100% + 6px);right:0;background:#fff;border:1px solid #e2e8f0;border-radius:10px;box-shadow:0 4px 16px rgba(0,0,0,.15);padding:6px;min-width:200px}') | Out-Null
$sb.AppendLine('#toast{display:none;position:fixed;bottom:24px;right:24px;background:#1e3a5f;color:#fff;padding:10px 20px;border-radius:8px;font-size:13px;z-index:9999;box-shadow:0 4px 12px rgba(0,0,0,.2)}') | Out-Null
$sb.AppendLine('.qlinks{display:flex;align-items:center;gap:8px;margin-bottom:18px;flex-wrap:wrap}') | Out-Null
$sb.AppendLine('.ql-lbl{font-size:11px;color:#94a3b8;font-weight:600;text-transform:uppercase;letter-spacing:.5px;margin-right:4px}') | Out-Null
$sb.AppendLine('.ql-btn{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border:1px solid #e2e8f0;border-radius:6px;font-size:12px;color:#475569;text-decoration:none;background:#fff;white-space:nowrap;box-shadow:0 1px 2px rgba(0,0,0,.04)}.ql-btn:hover{background:#f1f5f9;color:#1e3a5f;border-color:#94a3b8}') | Out-Null
$sb.AppendLine('.today-panel{background:#fff;border-radius:10px;padding:14px 20px;box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:18px}') | Out-Null
$sb.AppendLine('.tp-title{font-size:11px;font-weight:700;color:#dc2626;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}.td-task{font-size:12.5px;padding:3px 0;border-bottom:1px solid #f8fafc;display:flex;align-items:baseline;gap:6px;flex-wrap:wrap}.td-task:last-child{border-bottom:none}.td-exec{color:#64748b;font-size:11.5px}.td-none{font-size:12px;color:#94a3b8;font-style:italic}') | Out-Null
$sb.AppendLine('</style></head><body>') | Out-Null
$sb.AppendLine('<div id="auth-wall" style="display:none;position:fixed;top:0;right:0;bottom:0;left:0;background:#eef2f7;z-index:9999;align-items:center;justify-content:center"><div style="background:#fff;border-radius:12px;padding:40px;box-shadow:0 4px 24px rgba(0,0,0,.12);text-align:center;width:320px"><div style="font-size:20px;font-weight:700;color:#1e3a5f;margin-bottom:6px">Дашборд ОПУ</div><div style="font-size:13px;color:#64748b;margin-bottom:24px">Введите пароль для доступа</div><input id="auth-inp" type="password" placeholder="Пароль" style="width:100%;padding:10px 14px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;outline:none;box-sizing:border-box;margin-bottom:8px"><div style="font-size:11px;color:#94a3b8;margin-bottom:12px">Запросите пароль у администратора Сандиева Д.Е.</div><div id="auth-err" style="display:none;color:#dc2626;font-size:12px;margin-bottom:10px">Неверный пароль</div><button onclick="checkAuth()" style="width:100%;padding:10px;background:#1e3a5f;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer">Войти</button></div></div>') | Out-Null
$sb.AppendLine('<script>(function(){if(localStorage.getItem("opu_auth")==="1"){return;}var w=document.getElementById("auth-wall");w.style.display="flex";document.getElementById("auth-inp").addEventListener("keydown",function(e){if(e.key==="Enter")checkAuth();});document.getElementById("auth-inp").focus();})();function checkAuth(){var v=document.getElementById("auth-inp").value;if(v===' + ($dashPwd | ConvertTo-Json) + '){localStorage.setItem("opu_auth","1");document.getElementById("auth-wall").style.display="none";}else{document.getElementById("auth-err").style.display="block";document.getElementById("auth-inp").value="";document.getElementById("auth-inp").focus();}}</script>') | Out-Null

$sb.AppendLine("<div class='hdr'><div><h1>Дашборд ОПУ</h1><div class='hdr-sub'>Реестр заявок и проектов · обновлено $genTimeDisplay</div><div class='pulse' style='background:$(Esc $healthBg);color:$(Esc $healthColor)' title='Пульс системы: комплексный показатель состояния портфеля заявок (0-100). Снижается за: просроченные задачи (до -50 баллов), зависшие 30+ дней (до -20), задачи без исполнителя (до -10), задачи без дедлайна (до -10). 80-100 = хорошо, 60-79 = внимание, ниже 60 = критично.'>&#9679; Пульс: $healthScore / 100 — $healthLabel</div></div><div class='hdr-right'><button class='btn-print no-print' onclick='window.print()' title='Распечатать'><svg xmlns='http://www.w3.org/2000/svg' width='15' height='15' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><polyline points='6 9 6 2 18 2 18 9'/><path d='M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2'/><rect x='6' y='14' width='12' height='8'/></svg></button><button id='btn-upd' class='btn-refresh' onclick='hardReload()' title='Обновить страницу — Ctrl+Shift+R'><svg xmlns='http://www.w3.org/2000/svg' width='15' height='15' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><polyline points='23 4 23 10 17 10'/><polyline points='1 20 1 14 7 14'/><path d='M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15'/></svg></button><button id='btn-ar' class='btn-ar no-print' onclick='toggleAR()' title='Автообновление каждые 5 мин — нажмите для включения'><svg xmlns='http://www.w3.org/2000/svg' width='14' height='14' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='12' r='10'/><polyline points='12 6 12 12 16 14'/></svg><span class='ar-t'></span></button><div class='date-badge'>$dateStr</div></div></div>") | Out-Null
$remBadgeClass = if ($remCount -gt 0) { "rem-badge on" } else { "rem-badge" }
$sb.AppendLine("<div class='tab-nav no-print'><button class='tab-btn active' id='tab-btn-dash' onclick='switchTab(""dash"")'>Дашборд</button><button class='tab-btn' id='tab-btn-plnr' onclick='switchTab(""plnr"")'><svg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='4' width='18' height='18' rx='2' ry='2'/><line x1='16' y1='2' x2='16' y2='6'/><line x1='8' y1='2' x2='8' y2='6'/><line x1='3' y1='10' x2='21' y2='10'/></svg>Планировщик</button><button class='tab-btn' id='tab-btn-plan' onclick='switchTab(""plan"")'><svg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2'/><circle cx='9' cy='7' r='4'/><path d='M23 21v-2a4 4 0 0 0-3-3.87'/><path d='M16 3.13a4 4 0 0 1 0 7.75'/></svg>Исполнители</button><button class='tab-btn' id='tab-btn-rem' onclick='switchTab(""rem"")'><svg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9'/><path d='M13.73 21a2 2 0 0 1-3.46 0'/></svg>Назначения сегодня<span class='$remBadgeClass'>$remCount</span></button><button class='tab-btn' id='tab-btn-ai' onclick='switchTab(""ai"")'><svg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='11' width='18' height='10' rx='2'/><circle cx='12' cy='5' r='2'/><path d='M12 7v4'/><circle cx='8' cy='16' r='1.5' fill='currentColor' stroke='none'/><circle cx='16' cy='16' r='1.5' fill='currentColor' stroke='none'/></svg>ИИ-ассистент</button><button class='tab-btn' id='tab-btn-proto' onclick='switchTab(""proto"")'><svg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z'/><polyline points='14 2 14 8 20 8'/><line x1='9' y1='15' x2='15' y2='15'/><line x1='9' y1='11' x2='12' y2='11'/></svg>Протокол &#8594; задачи</button><button class='tab-btn' id='tab-btn-mc' onclick='switchTab(""mc"")'><svg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><polyline points='23 6 13.5 15.5 8.5 10.5 1 18'/><polyline points='17 6 23 6 23 12'/></svg>Прогноз</button><button class='tab-btn' id='tab-btn-sver' onclick='switchTab(""sver"")'><svg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M8 3 4 7l4 4'/><path d='M4 7h16'/><path d='m16 21 4-4-4-4'/><path d='M20 17H4'/></svg>Сверка Битрикс</button></div>") | Out-Null
$sb.AppendLine('<div id="stale-banner" class="stale-banner no-print"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg><span>Данные могут быть неактуальны. За обновлением обратиться к <b>Сандиеву Д.Е.</b></span></div>') | Out-Null
$sb.AppendLine('<div id="tab-dash">') | Out-Null
$sb.AppendLine('<div class="g-search-bar no-print"><div class="g-search-wrap"><span class="g-srch-ico"><svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg></span><input type="text" id="g-srch" placeholder="Поиск по всему дашборду — заявки, проекты, этапы, исполнители..." oninput="globalSearch()"><button id="g-srch-clear" class="g-srch-clear" onclick="clearGsrch()" title="Очистить">&#10005;</button></div><span id="g-srch-cnt"></span></div>') | Out-Null
$sb.AppendLine('<div id="watch-panel" class="panel no-print" style="display:none;margin-bottom:18px"><h2>&#9733; Под наблюдением</h2><div id="watch-list"></div></div>') | Out-Null
$sb.AppendLine('<div class="cards">') | Out-Null
$sb.AppendLine("<div class='card'><div class='num'>$($records.Count)</div><div class='lbl'>активных заявок</div></div>") | Out-Null
$sb.AppendLine("<div class='card red'><div class='num'>$($overdue.Count)</div><div class='lbl'>просрочено</div></div>") | Out-Null
$sb.AppendLine("<div class='card amber'><div class='num'>$($g4.Count)</div><div class='lbl'>на тестировании</div></div>") | Out-Null
$sb.AppendLine("<div class='card green'><div class='num'>$doneThisMonth</div><div class='lbl'>выполнено в $monthName</div></div>") | Out-Null
$sb.AppendLine('</div>') | Out-Null
$todayTitle = if ($todayTasks.Count -gt 0) { "<div class='tp-title'>Дедлайн сегодня ($($todayTasks.Count))</div>" } else { "" }
$sb.AppendLine("<div class='today-panel'>$todayTitle$todayTasksHtml</div>") | Out-Null

$sb.AppendLine('<div class="panel" style="margin-bottom:18px">') | Out-Null
$sb.AppendLine('<h2>Этапы работы</h2>') | Out-Null
$sb.AppendLine("<div class='filter-bar'><input type='text' id='srch' placeholder='Поиск по названию или ID...' oninput='filterTasks()'><select id='f-ex' onchange='filterTasks()'><option value=''>Все исполнители</option>$execOpts</select><select id='f-dp' onchange='filterTasks()'><option value=''>Все отделы</option>$deptOpts</select><button class='btn-reset' onclick=""document.getElementById('srch').value='';document.getElementById('f-ex').value='';document.getElementById('f-dp').value='';filterTasks()"">&#10005; Сбросить</button></div>") | Out-Null
$sb.AppendLine("<div class='grps'>$g1html$g2html$g3html$g4html</div>") | Out-Null
$sb.AppendLine("<div class='grps' style='margin-top:12px'>$g5html$guHtml</div>") | Out-Null
$sb.AppendLine('</div>') | Out-Null

$sb.AppendLine('<div class="row r2" style="margin-bottom:18px">') | Out-Null
$sb.AppendLine('<div class="panel"><h2>Загрузка исполнителей</h2><div class="ch"><canvas id="eC"></canvas></div></div>') | Out-Null
$sb.AppendLine("<div class='panel'><h2>Тепловая карта дедлайнов (60 дней)</h2>$calHtml</div>") | Out-Null
$sb.AppendLine('</div>') | Out-Null
$peakRows = ""
if ($peakExecs.Count -eq 0) {
    $peakRows = "<div class='peak-empty'>Пиковой нагрузки не ожидается</div>"
} else {
    foreach ($pe in $peakExecs) {
        $cnt = $pe.Count
        $clr = if ($cnt -ge 4) { "#dc2626" } elseif ($cnt -ge 2) { "#d97706" } else { "#16a34a" }
        $bg  = if ($cnt -ge 4) { "#fee2e2" } elseif ($cnt -ge 2) { "#fef3c7" } else { "#dcfce7" }
        $lbl = if ($cnt -ge 4) { "Критично" } elseif ($cnt -ge 2) { "Высокая" } else { "Умеренная" }
        $tip = ($pe.Group | Sort-Object DaysLeft | ForEach-Object {
            $dl = if ($_.DaysLeft -eq 0) { "сегодня" } elseif ($_.DaysLeft -eq 1) { "завтра" } else { "через $($_.DaysLeft) дн." }
            "$(Esc $_.ID): $(Esc $_.Name) ($dl)"
        }) -join "&#10;"
        $peakRows += "<div class='peak-row' title='$tip'><div class='peak-bar' style='background:$clr'></div><span class='peak-name'>$(Esc (Strip-Initials $pe.Name))</span><span class='peak-badge' style='background:$bg;color:$clr'>$cnt</span><span class='peak-lbl' style='color:$clr'>$lbl</span></div>"
    }
}
$sb.AppendLine('<div class="row" style="grid-template-columns:2fr 1fr;margin-bottom:18px">') | Out-Null
$sb.AppendLine('<div class="panel"><h2>Нагрузка по отделам</h2><div class="ch"><canvas id="dC"></canvas></div></div>') | Out-Null
$sb.AppendLine("<div class='panel'><h2>Пик нагрузки — 7 дней</h2><div style='font-size:11px;color:#94a3b8;margin-bottom:12px'>Исполнители с дедлайнами на следующие 7 дней</div>$peakRows</div>") | Out-Null
$sb.AppendLine('</div>') | Out-Null

$sb.AppendLine("<div class='panel' style='margin-bottom:18px'><h2>Выполнимость к дедлайнам — ближайшие $planDays дней</h2>$loadHtml</div>") | Out-Null

$sb.AppendLine('<div class="row r32">') | Out-Null
$projLegend = "<div style='display:flex;align-items:center;gap:14px;margin-bottom:12px;flex-wrap:wrap'><span style='font-size:11px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.5px'>Индекс здоровья проекта</span><span style='display:flex;align-items:center;gap:4px;font-size:11px;color:#475569'><span class='ph-dot ph-g'></span>Норма (75-100)</span><span style='display:flex;align-items:center;gap:4px;font-size:11px;color:#475569'><span class='ph-dot ph-a'></span>Внимание (50-74)</span><span style='display:flex;align-items:center;gap:4px;font-size:11px;color:#475569'><span class='ph-dot ph-r'></span>Критично (&lt;50)</span><span style='font-size:10px;color:#94a3b8;margin-left:4px'>Наведите на точку для деталей</span></div>"
$sb.AppendLine("<div class='panel'><div style='display:flex;align-items:center;justify-content:space-between;margin-bottom:14px'><h2 style='margin-bottom:0'>Прогресс проектов</h2><button class='btn-urgent no-print' id='btn-urgent' onclick='toggleUrgent()'>Только срочное</button></div>$projLegend$projBars</div>") | Out-Null
$sb.AppendLine("<div class='panel'><h2>Ближайшие дедлайны</h2><table id='up-tbl'><thead><tr><th class='srt' onclick='sortTable(this)'>ID</th><th class='srt' onclick='sortTable(this)'>Заявка</th><th class='srt' onclick='sortTable(this)'>Исп.</th><th class='srt' onclick='sortTable(this)'>Дата</th><th class='srt' onclick='sortTable(this)'>Осталось</th></tr></thead><tbody>$upcomingRows</tbody></table></div>") | Out-Null
$sb.AppendLine('</div>') | Out-Null

$staleHint = if ($stale.Count -gt 0) { "<div class='stale-hint'>из них $($stale.Count) зависло (просрочка 30+ дней) - выделено красным фоном</div>" } else { "" }
$sb.AppendLine("<div class='panel'><div style='display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:14px'><div><h2 style='margin:0'>Просрочено ($($overdue.Count))</h2>$staleHint</div><button class='btn-copy no-print' onclick='copyOverdue()'>Скопировать список</button></div><table id='ov-tbl'><thead><tr><th class='srt' onclick='sortTable(this)'>ID</th><th class='srt' onclick='sortTable(this)'>Заявка</th><th class='srt' onclick='sortTable(this)'>Исполнитель</th><th class='srt' onclick='sortTable(this)'>Просрочка</th></tr></thead><tbody>$overdueRows</tbody></table></div>") | Out-Null
$sb.AppendLine("<details class='nodl'><summary>Задачи без дедлайна ($($noDl.Count))</summary><table><thead><tr><th>ID</th><th>Заявка</th><th>Исполнитель</th><th>Отдел</th></tr></thead><tbody>$noDlRows</tbody></table></details>") | Out-Null

$sb.AppendLine('</div>') | Out-Null
$todayFull = $today.ToString("dd MMMM yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("ru-RU"))
$sb.AppendLine("<div id='tab-rem' style='display:none'><div class='panel'><h2>Назначения</h2><div class='rem-date'>$todayFull</div>$remHtml</div></div>") | Out-Null
$sb.AppendLine('<div id="tab-ai" style="display:none"><div class="panel"><div id="ai-setup"><h2 style="margin-bottom:8px">ИИ-ассистент</h2><p style="color:#64748b;font-size:13px;margin-bottom:0">Введите API-ключ Groq для работы с ассистентом</p><div class="ai-key-row"><input type="password" id="ai-key-inp" placeholder="gsk_..."><button class="ai-key-save" onclick="saveAIKey()">Сохранить</button></div><p class="ai-hint">Ключ хранится только в вашем браузере (localStorage) и не передаётся никуда, кроме api.groq.com</p></div><div id="ai-chat" style="display:none"><div class="ai-hdr"><h2>ИИ-ассистент по портфелю</h2><div class="ai-hdr-btns"><button class="ai-hdr-btn" id="ai-mode-btn" onclick="toggleAIMode()">Кратко</button><button class="ai-hdr-btn" onclick="clearAIChat()">Очистить</button><button class="ai-hdr-btn" onclick="resetAIKey()">Сменить ключ</button></div></div><div class="ai-chips"><button class="ai-chip" onclick="aiQuick(this)">Угрозы срыва</button><button class="ai-chip" onclick="aiQuick(this)">Кто перегружен?</button><button class="ai-chip" onclick="aiQuick(this)">Статус-отчёт</button><button class="ai-chip" onclick="aiQuick(this)">Что назначить сегодня?</button><button class="ai-chip" onclick="aiQuick(this)">Просрочки по отделам</button><button class="ai-chip ai-chip-scen" onclick="sdTogglePanel()" title="Конструктор сценариев">&#9889; Что если...</button></div><div class="ai-scen-hint">Или введите текстом: «Что если снять Иванова с проекта А и передать Петрову?» — ИИ пересчитает загрузку.</div><div id="sd-panel" class="sd-panel" style="display:none"><div class="sd-panel-hdr"><span>&#9889; Сценарное моделирование</span><button onclick="document.getElementById(''sd-panel'').style.display=''none''">&#10005;</button></div><div class="sd-form"><div class="sd-row" style="grid-column:1/-1"><label>Сценарий</label><select id="sd-action" onchange="sdFormToggle()"><option value="reassign">Снять / передать задачи</option><option value="absence">Отпуск / отсутствие</option><option value="newtask">Добавить новую заявку</option><option value="shift">Сдвинуть дедлайн</option></select></div><div class="sd-row" id="sd-from-row"><label>Исполнитель</label><select id="sd-from"></select></div><div class="sd-row" id="sd-to-row"><label>Передать (необяз.)</label><select id="sd-to"><option value="">— без исполнителя —</option></select></div><div class="sd-row" id="sd-proj-row"><label>Проект (необяз.)</label><select id="sd-proj"><option value="">— все задачи —</option></select></div><div class="sd-row" id="sd-days-row" style="display:none"><label>Количество дней</label><input type="number" id="sd-days" min="1" max="60" value="5"></div><div class="sd-row" id="sd-df-row" style="display:none"><label>С даты (дд.мм)</label><input type="text" id="sd-date-from" placeholder="10.07"></div><div class="sd-row" id="sd-dt-row" style="display:none"><label>По дату (дд.мм)</label><input type="text" id="sd-date-to" placeholder="20.07"></div><button class="sd-calc-btn" onclick="runScenarioForm()">Рассчитать</button></div></div><div id="ai-msgs" class="ai-msgs"></div><div class="ai-input-row"><input type="text" id="ai-inp" placeholder="Задайте вопрос по проектному портфелю..." onkeydown="if(event.key===''Enter''&&!event.shiftKey){event.preventDefault();sendAI();}"><button id="ai-send-btn" class="ai-send-btn" onclick="sendAI()">Отправить</button></div></div></div></div>') | Out-Null
$sb.AppendLine('<div id="tab-plan" style="display:none"><div class="panel"><h2 style="margin-bottom:4px">Загрузка исполнителей по дням</h2><p style="color:#64748b;font-size:13px;margin:0 0 14px">Задачи размещены в день дедлайна. Наведите на блок для деталей.</p><div class="plan-nav"><button onclick="planShift(-1)">&#8592; Назад</button><button onclick="planShift(0)">Сегодня</button><button onclick="planShift(1)">Вперёд &#8594;</button><span class="plan-range" id="plan-range"></span></div><div class="plan-wrap"><div id="plan-grid" onmousemove="planTip(event)" onmouseleave="planTipHide()"></div></div><div class="plan-legend"><span><i style="background:#eff6ff;border-color:#3b82f6"></i>Заявка</span><span><i style="background:#f0fdfa;border-color:#14b8a6"></i>Проект</span><span><i style="background:#fffbeb;border-color:#f59e0b"></i>Дедлайн 0-2 дня</span><span><i style="background:#fef2f2;border-color:#ef4444"></i>Просрочено</span></div></div><div class="plan-tip" id="plan-tip"></div></div>') | Out-Null
$sb.AppendLine('<div id="tab-plnr" style="display:none"><iframe id="plnr-frame" title="Планировщик ОПУ" data-src="https://script.google.com/macros/s/AKfycbxtKdU0aX_VZaSXz00JjhJMJECqjM7qdeJEolofeVdpmXqFPDAuuYVr-y8BOBAss6GbxQ/exec" style="width:100%;height:calc(100vh - 160px);min-height:560px;border:1px solid #e2e8f0;border-radius:10px;background:#fff;display:block"></iframe></div>') | Out-Null
$sb.AppendLine(@'
<div id="tab-proto" style="display:none">
<div class="panel">
<h2>Протокол &#8594; задачи</h2>
<p style="color:#64748b;font-size:13px;margin:-4px 0 14px">Вставьте текст протокола совещания - ИИ выделит поручения (кто, что, срок) и подготовит их к переносу в реестр. Ключ Groq берётся из вкладки «ИИ-ассистент».</p>
<style>
#tab-proto textarea{width:100%;min-height:170px;border:1.5px solid #e2e8f0;border-radius:8px;padding:10px;font-family:inherit;font-size:13px;line-height:1.5;outline:none;resize:vertical}
#tab-proto textarea:focus{border-color:#1e3a5f}
#tab-proto .pr-row{display:flex;gap:12px;flex-wrap:wrap;align-items:flex-end;margin-top:12px}
#tab-proto .pr-row label{font-size:12px;font-weight:600;color:#475569;display:block;margin-bottom:5px}
#tab-proto input[type=date],#tab-proto select{border:1.5px solid #e2e8f0;border-radius:8px;padding:8px 10px;font-family:inherit;font-size:13px;outline:none}
#tab-proto .pr-btn{padding:9px 16px;border:1.5px solid #e2e8f0;border-radius:8px;background:#fff;color:#334155;font-size:13px;font-weight:600;cursor:pointer}
#tab-proto .pr-btn:hover{background:#f1f5f9}
#tab-proto .pr-primary{background:#1e3a5f;color:#fff;border-color:#1e3a5f}
#tab-proto .pr-primary:hover{background:#274b73}
#tab-proto .pr-status{font-size:13px;color:#64748b;margin-top:10px;min-height:18px}
#tab-proto .pr-status.err{color:#dc2626}
#tab-proto .pr-status.busy{color:#d97706}
#tab-proto table{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:8px}
#tab-proto th{background:#f8fafc;color:#475569;text-align:left;padding:8px;border-bottom:2px solid #e2e8f0;white-space:nowrap}
#tab-proto td{border-bottom:1px solid #eef2f7;padding:4px 6px;vertical-align:top}
#tab-proto td input{border:1px solid transparent;border-radius:6px;padding:5px 6px;background:transparent;width:100%;font-family:inherit;font-size:12.5px}
#tab-proto td input:hover{border-color:#e2e8f0}
#tab-proto td input:focus{border-color:#1e3a5f;background:#f8fbff}
#tab-proto .pr-chk{width:34px;text-align:center}
#tab-proto .pr-chk input{width:16px;height:16px}
#tab-proto .pr-wrap{overflow-x:auto}
</style>
<textarea id="pr-text" placeholder="Вставьте сюда текст протокола совещания..."></textarea>
<div class="pr-row">
  <div><label>Дата совещания</label><input type="date" id="pr-date"></div>
  <div><label>Модель</label><select id="pr-model"><option value="llama-3.3-70b-versatile">Llama 3.3 70B (точнее)</option><option value="llama-3.1-8b-instant">Llama 3.1 8B (быстрее)</option></select></div>
  <div><button class="pr-btn pr-primary" onclick="protoExtract()">Извлечь задачи</button> <button class="pr-btn" onclick="protoSample()">Пример</button></div>
</div>
<div class="pr-status" id="pr-status"></div>
<div id="pr-result" style="display:none;margin-top:14px">
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;flex-wrap:wrap;gap:8px">
    <span id="pr-count" style="font-size:12px;color:#64748b;font-weight:600"></span>
    <div style="display:flex;gap:8px;flex-wrap:wrap">
      <button class="pr-btn" onclick="protoCopyTSV()" title="Скопировать выбранные строки для вставки в Google-таблицу">Копировать для реестра</button>
      <button class="pr-btn" onclick="protoCopyList()">Копировать списком</button>
      <button class="pr-btn" onclick="protoCSV()">Скачать CSV</button>
    </div>
  </div>
  <div class="pr-wrap"><table><thead><tr><th class="pr-chk">&#10003;</th><th>Задача</th><th>Исполнитель</th><th>Срок</th><th>Проект / контекст</th><th>Цитата</th></tr></thead><tbody id="pr-tbody"></tbody></table></div>
  <p style="font-size:11.5px;color:#94a3b8;margin-top:8px">Проверьте и поправьте прямо в таблице. Галочка слева - включать ли строку. «Копировать для реестра» кладёт выбранные строки через табуляцию - вставьте в Google-таблицу, разложится по столбцам.</p>
</div>
</div>
</div>
'@) | Out-Null
$sb.AppendLine(@'
<div id="tab-mc" style="display:none">
<div class="panel">
<h2>Прогноз портфеля (Монте-Карло)</h2>
<p class="mc-intro">Оценки задач неточны. Модель тысячи раз «проигрывает» портфель со случайной длительностью каждой задачи (трёхточечная оценка) и показывает ВЕРОЯТНОСТЬ уложиться в сроки, а не один ответ. Крутите ползунки «лучший/худший случай» - и видно, как меняется риск.</p>
<style>
#tab-mc .mc-intro{color:#64748b;font-size:13px;margin:-4px 0 14px;max-width:860px;line-height:1.5}
#tab-mc .mc-controls{display:flex;gap:14px;flex-wrap:wrap;align-items:flex-end;background:#f8fafc;border:1px solid #eef2f7;border-radius:10px;padding:14px}
#tab-mc .mc-controls label{font-size:11px;font-weight:600;color:#475569;display:block;margin-bottom:4px}
#tab-mc .mc-controls input{width:92px;border:1.5px solid #e2e8f0;border-radius:7px;padding:7px 9px;font-family:inherit;font-size:13px;outline:none}
#tab-mc .mc-controls select{border:1.5px solid #e2e8f0;border-radius:7px;padding:7px 9px;font-family:inherit;font-size:13px;outline:none;max-width:240px}
#tab-mc .mc-controls input[type=date]{width:150px}
#tab-mc .mc-controls input:focus{border-color:#1e3a5f}
#tab-mc .mc-u{font-size:11px;color:#94a3b8;margin-left:3px}
#tab-mc .mc-btn{padding:9px 18px;border:1.5px solid #1e3a5f;border-radius:8px;background:#1e3a5f;color:#fff;font-size:13px;font-weight:600;cursor:pointer}
#tab-mc .mc-btn:hover{background:#274b73}
#tab-mc .mc-status{font-size:12.5px;color:#64748b;margin:10px 0;min-height:16px}
#tab-mc .mc-status.err{color:#dc2626}
#tab-mc .mc-status.busy{color:#d97706}
#tab-mc .mc-cards{display:flex;gap:14px;flex-wrap:wrap;margin:6px 0 16px}
#tab-mc .mc-card{flex:1;min-width:180px;background:#fff;border:1px solid #eef2f7;border-radius:12px;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,.05)}
#tab-mc .mc-num{font-size:34px;font-weight:700;line-height:1;color:#1e3a5f}
#tab-mc .mc-lbl{font-size:12px;color:#64748b;margin-top:6px}
#tab-mc .mc-verdict{margin:8px 0 18px}
#tab-mc .mc-v{position:relative;border-radius:12px;padding:16px 18px;border:1px solid}
#tab-mc .mc-v-red{background:#fef2f2;border-color:#fecaca}
#tab-mc .mc-v-amber{background:#fffbeb;border-color:#fde68a}
#tab-mc .mc-v-green{background:#f0fdf4;border-color:#bbf7d0}
#tab-mc .mc-v-head{font-size:16px;font-weight:700;margin-bottom:9px;padding-right:96px}
#tab-mc .mc-v-red .mc-v-head{color:#b91c1c}
#tab-mc .mc-v-amber .mc-v-head{color:#b45309}
#tab-mc .mc-v-green .mc-v-head{color:#15803d}
#tab-mc .mc-v-body{font-size:13px;color:#334155;line-height:1.5;margin-bottom:5px}
#tab-mc .mc-v-todo{font-size:13px;color:#1e293b;font-weight:600;margin-top:9px;padding-top:9px;border-top:1px solid rgba(0,0,0,.08)}
#tab-mc .mc-v-copy{position:absolute;top:14px;right:14px;padding:5px 12px;border:1px solid rgba(0,0,0,.15);background:rgba(255,255,255,.75);border-radius:7px;font-size:11.5px;font-weight:600;color:#334155;cursor:pointer}
#tab-mc .mc-v-copy:hover{background:#fff}
#tab-mc .mc-details{margin:6px 0 8px}
#tab-mc .mc-details summary{cursor:pointer;font-size:12.5px;color:#64748b;font-weight:600;padding:6px 0;list-style:revert}
#tab-mc .mc-details summary:hover{color:#1e3a5f}
#tab-mc .mc-details[open] summary{margin-bottom:10px}
#tab-mc .mc-hist-wrap{overflow-x:auto;margin-bottom:16px}
#tab-mc canvas{max-width:100%}
#tab-mc .mc-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
#tab-mc .mc-grid h3{font-size:13px;color:#1e3a5f;margin-bottom:8px}
#tab-mc table{width:100%;border-collapse:collapse;font-size:12.5px}
#tab-mc th{background:#f8fafc;color:#475569;text-align:left;padding:7px 8px;border-bottom:2px solid #e2e8f0;font-weight:600}
#tab-mc td{padding:6px 8px;border-bottom:1px solid #eef2f7}
#tab-mc .mc-tbl-wrap{overflow-x:auto}
#tab-mc .mc-note{font-size:11.5px;color:#94a3b8;margin-top:14px;line-height:1.5;max-width:900px}
@media(max-width:720px){#tab-mc .mc-grid{grid-template-columns:1fr}}
</style>
<div class="mc-controls">
  <div><label>Проект</label><select id="mc-proj"><option value="">Весь портфель</option></select></div>
  <div><label>Прогонов</label><input type="number" id="mc-runs" value="3000" min="500" max="20000" step="500"></div>
  <div><label>Лучший случай</label><input type="number" id="mc-opt" value="15" min="0" max="50" step="5"><span class="mc-u">% быстрее</span></div>
  <div><label>Худший случай</label><input type="number" id="mc-pess" value="60" min="0" max="200" step="10"><span class="mc-u">% дольше</span></div>
  <div><label>Оценка по умолчанию</label><input type="number" id="mc-def" value="2" min="0.5" max="30" step="0.5"><span class="mc-u">дн.</span></div>
  <div><label>Целевая дата</label><input type="date" id="mc-target"></div>
  <div><button class="mc-btn" onclick="mcRunSim()">Рассчитать</button></div>
</div>
<div class="mc-status" id="mc-status"></div>
<div id="mc-result" style="display:none">
  <div id="mc-verdict" class="mc-verdict"></div>
  <div class="mc-grid">
    <div><h3>Задачи под наибольшим риском</h3><div class="mc-tbl-wrap"><table id="mc-risk"></table></div></div>
    <div><h3>Узкие места (исполнители)</h3><div class="mc-tbl-wrap"><table id="mc-bottleneck"></table></div></div>
  </div>
  <details class="mc-details">
    <summary>Подробности: вероятности и график распределения</summary>
    <div class="mc-cards">
      <div class="mc-card"><div class="mc-num" id="mc-p-deadlines">-</div><div class="mc-lbl">уложиться во ВСЕ дедлайны сразу</div></div>
      <div class="mc-card"><div class="mc-num" id="mc-p-target">-</div><div class="mc-lbl" id="mc-target-lbl">портфель к целевой дате</div></div>
      <div class="mc-card"><div class="mc-num" id="mc-p80">-</div><div class="mc-lbl">дата «готово на 80%»</div></div>
    </div>
    <div class="mc-hist-wrap"><canvas id="mc-hist" width="820" height="230"></canvas></div>
  </details>
  <p class="mc-note" id="mc-note"></p>
</div>
</div>
</div>
'@) | Out-Null
$sb.AppendLine('<div id="gantt-modal" style="display:none;position:fixed;top:0;right:0;bottom:0;left:0;background:rgba(0,0,0,.55);z-index:2000;align-items:center;justify-content:center"><div class="gantt-box"><div class="gantt-hdr"><span id="gantt-title" class="gantt-ttl"></span><button class="gantt-close" onclick="closeGantt()">&#215;</button></div><div class="gantt-body" id="gantt-body"></div></div></div>') | Out-Null
$sb.AppendLine('<div id="toast">Список скопирован</div>') | Out-Null
$sb.AppendLine("<div id='restore-dock'><div id='rd-list'></div><button class='no-print' onclick='toggleRestoreDock()' title='Скрытые панели' style='display:flex;align-items:center;gap:6px;padding:8px 14px;background:#1e3a5f;color:#fff;border:none;border-radius:8px;cursor:pointer;font-size:12px;font-weight:600;box-shadow:0 2px 8px rgba(0,0,0,.25);white-space:nowrap'><svg xmlns='http://www.w3.org/2000/svg' width='14' height='14' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z'/><circle cx='12' cy='12' r='3'/></svg><span id='rd-lbl'>Скрытые (0)</span></button></div>") | Out-Null
$sb.AppendLine('<script>') | Out-Null
$sb.AppendLine("const ganttData=$ganttJson;") | Out-Null
$sb.AppendLine("const SD=$sdJson;") | Out-Null
$sb.AppendLine("const eL=$execLabels,eOver=$execOverdue,eUrg=$execUrgent,eNorm=$execNormal;") | Out-Null
$sb.AppendLine("const dL=$deptLabels,dOv=$deptOver,dUr=$deptUrg,dNo=$deptNorm;") | Out-Null
$sb.AppendLine('const stackedCfg=(labels,d1,d2,d3)=>({type:"bar",data:{labels,datasets:[{label:"Просрочено",data:d1,backgroundColor:"#fca5a5",borderRadius:0},{label:"Срок ≤7 дн.",data:d2,backgroundColor:"#fcd34d",borderRadius:0},{label:"В норме",data:d3,backgroundColor:"#93c5fd",borderRadius:4}]},options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom",labels:{font:{size:11},padding:6}}},scales:{x:{stacked:true,beginAtZero:true,ticks:{stepSize:1,precision:0}},y:{stacked:true}}}});') | Out-Null
$sb.AppendLine('new Chart(document.getElementById("eC"),stackedCfg(eL,eOver,eUrg,eNorm));') | Out-Null
$sb.AppendLine('new Chart(document.getElementById("dC"),stackedCfg(dL,dOv,dUr,dNo));') | Out-Null
$sb.AppendLine('function sortTable(th){') | Out-Null
$sb.AppendLine('  const tbl=th.closest("table");const tbody=tbl.querySelector("tbody");') | Out-Null
$sb.AppendLine('  const col=Array.from(th.parentElement.children).indexOf(th);') | Out-Null
$sb.AppendLine('  const asc=th.classList.contains("srt-a");') | Out-Null
$sb.AppendLine('  tbl.querySelectorAll("th.srt").forEach(t=>t.classList.remove("srt-a","srt-d"));') | Out-Null
$sb.AppendLine('  th.classList.add(asc?"srt-d":"srt-a");') | Out-Null
$sb.AppendLine('  const rows=Array.from(tbody.querySelectorAll("tr"));') | Out-Null
$sb.AppendLine('  rows.sort((a,b)=>{') | Out-Null
$sb.AppendLine('    const ca=a.children[col],cb=b.children[col];') | Out-Null
$sb.AppendLine('    const va=ca.dataset.v!==undefined?parseFloat(ca.dataset.v):ca.textContent.trim();') | Out-Null
$sb.AppendLine('    const vb=cb.dataset.v!==undefined?parseFloat(cb.dataset.v):cb.textContent.trim();') | Out-Null
$sb.AppendLine('    if(!isNaN(va)&&!isNaN(vb))return asc?vb-va:va-vb;') | Out-Null
$sb.AppendLine('    return asc?String(vb).localeCompare(String(va),"ru"):String(va).localeCompare(String(vb),"ru");') | Out-Null
$sb.AppendLine('  });') | Out-Null
$sb.AppendLine('  rows.forEach(r=>tbody.appendChild(r));') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function filterTasks(){') | Out-Null
$sb.AppendLine('  const q=document.getElementById("srch").value.toLowerCase();') | Out-Null
$sb.AppendLine('  const ex=document.getElementById("f-ex").value;') | Out-Null
$sb.AppendLine('  const dp=document.getElementById("f-dp").value;') | Out-Null
$sb.AppendLine('  document.querySelectorAll(".grp").forEach(grp=>{') | Out-Null
$sb.AppendLine('    let n=0;') | Out-Null
$sb.AppendLine('    grp.querySelectorAll(".task").forEach(t=>{') | Out-Null
$sb.AppendLine('      const ok=(!q||(t.dataset.name||"").toLowerCase().includes(q)||(t.dataset.id||"").toLowerCase().includes(q))') | Out-Null
$sb.AppendLine('               &&(!ex||t.dataset.exec===ex)&&(!dp||t.dataset.dept===dp);') | Out-Null
$sb.AppendLine('      t.style.display=ok?"":"none"; if(ok)n++;') | Out-Null
$sb.AppendLine('    });') | Out-Null
$sb.AppendLine('    const c=grp.querySelector(".grp-cnt"); if(c)c.textContent=n;') | Out-Null
$sb.AppendLine('  });') | Out-Null
$sb.AppendLine('  highlightExec(ex);updateExecBadge(ex);') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function filterByExec(el){const name=el.textContent.trim();const sel=document.getElementById("f-ex");sel.value=(sel.value===name)?"":name;filterTasks();}') | Out-Null
$sb.AppendLine('function highlightExec(name){document.querySelectorAll(".t-exec").forEach(function(el){el.classList.toggle("exec-active",!!name&&el.textContent.trim()===name);});}') | Out-Null
$sb.AppendLine('function updateExecBadge(name){var b=document.getElementById("exec-filter-badge");if(!b){b=document.createElement("span");b.id="exec-filter-badge";b.className="exec-filter-badge";var fb=document.querySelector(".filter-bar");if(fb)fb.appendChild(b);}b.style.display=name?"inline-flex":"none";if(name){b.textContent=name+" ";var btn=document.createElement("button");btn.innerHTML="&#10005;";btn.title="Сбросить фильтр";btn.onclick=clearExecFilter;b.appendChild(btn);}}') | Out-Null
$sb.AppendLine('function clearExecFilter(){document.getElementById("f-ex").value="";filterTasks();}') | Out-Null
$sb.AppendLine('function copyOverdue(){') | Out-Null
$sb.AppendLine('  const rows=document.querySelectorAll("#ov-tbl tbody tr");') | Out-Null
$sb.AppendLine('  const lines=[];') | Out-Null
$sb.AppendLine('  rows.forEach(r=>{') | Out-Null
$sb.AppendLine('    const id=r.children[0].textContent.trim();') | Out-Null
$sb.AppendLine('    const nm=r.children[1].textContent.trim();') | Out-Null
$sb.AppendLine('    const ex=r.children[2].textContent.trim();') | Out-Null
$sb.AppendLine('    const dy=r.children[3].textContent.trim();') | Out-Null
$sb.AppendLine('    lines.push(id+" | "+nm+" | "+ex+" | "+dy);') | Out-Null
$sb.AppendLine('  });') | Out-Null
$sb.AppendLine('  if(!lines.length)return;') | Out-Null
$sb.AppendLine('  navigator.clipboard.writeText(lines.join("\n")).then(()=>{') | Out-Null
$sb.AppendLine('    const t=document.getElementById("toast");') | Out-Null
$sb.AppendLine('    t.style.display="block";') | Out-Null
$sb.AppendLine('    setTimeout(()=>{t.style.display="none";},2200);') | Out-Null
$sb.AppendLine('  });') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function openGantt(btn){') | Out-Null
$sb.AppendLine('  const b=btn.closest("summary").querySelector("b");') | Out-Null
$sb.AppendLine('  const pid=b?b.textContent.trim():"";') | Out-Null
$sb.AppendLine('  const data=ganttData[pid];if(!data)return;') | Out-Null
$sb.AppendLine('  document.getElementById("gantt-title").textContent=pid+" — "+data.name;') | Out-Null
$sb.AppendLine('  document.getElementById("gantt-body").innerHTML=renderGantt(data.tasks);') | Out-Null
$sb.AppendLine('  const m=document.getElementById("gantt-modal");m.style.display="flex";document.body.style.overflow="hidden";') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function closeGantt(){document.getElementById("gantt-modal").style.display="none";document.body.style.overflow="";}') | Out-Null
$sb.AppendLine(@'
var PR_ENDPOINT="https://api.groq.com/openai/v1/chat/completions";
function protoInit(){var d=document.getElementById("pr-date");if(d&&!d.value)d.value=new Date().toISOString().slice(0,10);}
function protoEsc(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function protoStatus(t,cls){var el=document.getElementById("pr-status");el.textContent=t||"";el.className="pr-status"+(cls?" "+cls:"");}
function protoToast(t){var el=document.getElementById("toast");if(!el)return;el.textContent=t;el.style.display="block";setTimeout(function(){el.style.display="none";},2000);}
function protoMessages(text,mdate){
  var sys="Ты - ассистент руководителя проектного отдела (ОПУ). Из текста протокола совещания извлеки ТОЛЬКО поручения и задачи (action items) - то, что конкретно кто-то должен сделать. Игнорируй обсуждения, факты и решения без действия. Дата совещания: "+mdate+". Относительные сроки (к пятнице, через неделю, до конца месяца) переводи в дату формата ГГГГ-ММ-ДД относительно даты совещания; если срок не указан - пустая строка. Исполнителя указывай как в тексте; если не назван - пустая строка. Поле quote - короткая точная цитата из протокола (до 15 слов). Верни СТРОГО JSON-объект: {\"tasks\":[{\"task\":\"строка\",\"assignee\":\"строка\",\"deadline\":\"ГГГГ-ММ-ДД или пусто\",\"project\":\"строка\",\"quote\":\"строка\"}]}. Если поручений нет - {\"tasks\":[]}.";
  return [{role:"system",content:sys},{role:"user",content:text}];
}
function protoParse(content){
  var s=String(content||"").trim().replace(/^```(json)?/i,"").replace(/```$/,"").trim();
  var obj;try{obj=JSON.parse(s);}catch(e){var a=s.indexOf("{"),b=s.lastIndexOf("}");if(a>=0&&b>a){obj=JSON.parse(s.slice(a,b+1));}else{throw new Error("Модель вернула не JSON");}}
  var arr=Array.isArray(obj)?obj:(obj.tasks||obj.items||[]);
  return arr.map(function(t){return {task:(t.task||t.title||"").trim(),assignee:(t.assignee||t.exec||"").trim(),deadline:(t.deadline||t.due||"").toString().trim(),project:(t.project||t.context||"").trim(),quote:(t.quote||"").trim()};}).filter(function(t){return t.task;});
}
async function protoExtract(){
  var text=document.getElementById("pr-text").value.trim();
  if(!text){protoStatus("Вставьте текст протокола.","err");return;}
  var key=localStorage.getItem("opu_ai_key");
  if(!key){protoStatus("Сначала укажите ключ Groq во вкладке ИИ-ассистент.","err");return;}
  var mdate=document.getElementById("pr-date").value||new Date().toISOString().slice(0,10);
  var model=document.getElementById("pr-model").value;
  protoStatus("Извлекаю задачи...","busy");
  try{
    var resp=await fetch(PR_ENDPOINT,{method:"POST",headers:{"Content-Type":"application/json","Authorization":"Bearer "+key},body:JSON.stringify({model:model,temperature:0.2,max_tokens:2000,response_format:{type:"json_object"},messages:protoMessages(text,mdate)})});
    if(!resp.ok){var tx=await resp.text();throw new Error("Groq "+resp.status+": "+tx.slice(0,180));}
    var data=await resp.json();
    var content=data.choices&&data.choices[0]&&data.choices[0].message&&data.choices[0].message.content;
    var tasks=protoParse(content);
    protoRender(tasks);
    protoStatus(tasks.length?("Найдено задач: "+tasks.length):"Поручений в тексте не найдено.");
  }catch(e){protoStatus("Ошибка: "+e.message,"err");}
}
function protoCell(v){return "<td><input value=\""+protoEsc(v)+"\"></td>";}
function protoRender(tasks){
  var tb=document.getElementById("pr-tbody");
  document.getElementById("pr-result").style.display="block";
  if(!tasks.length){tb.innerHTML="";document.getElementById("pr-count").textContent="";return;}
  tb.innerHTML=tasks.map(function(t){return "<tr><td class=\"pr-chk\"><input type=\"checkbox\" checked></td>"+protoCell(t.task)+protoCell(t.assignee)+protoCell(t.deadline)+protoCell(t.project)+protoCell(t.quote)+"</tr>";}).join("");
  tb.querySelectorAll("input[type=checkbox]").forEach(function(cb){cb.addEventListener("change",protoCount);});
  protoCount();
}
function protoCount(){var rows=document.querySelectorAll("#pr-tbody tr");var sel=0;rows.forEach(function(r){if(r.querySelector("input[type=checkbox]").checked)sel++;});document.getElementById("pr-count").textContent="Задач: "+rows.length+" · выбрано: "+sel;}
function protoSelected(){var out=[];document.querySelectorAll("#pr-tbody tr").forEach(function(r){if(!r.querySelector("input[type=checkbox]").checked)return;var inp=r.querySelectorAll("td input:not([type=checkbox])");out.push([inp[0].value.trim(),inp[1].value.trim(),inp[2].value.trim(),inp[3].value.trim(),inp[4].value.trim()]);});return out;}
function protoCopyTSV(){var rows=protoSelected();if(!rows.length){protoToast("Нет выбранных строк");return;}var tsv=rows.map(function(r){return r.slice(0,4).join("\t");}).join("\n");navigator.clipboard.writeText(tsv).then(function(){protoToast("Скопировано для реестра ("+rows.length+")");});}
function protoCopyList(){var rows=protoSelected();if(!rows.length){protoToast("Нет выбранных строк");return;}var txt=rows.map(function(r){return "- "+r[0]+(r[1]?" ("+r[1]+")":"")+(r[2]?", до "+r[2]:"");}).join("\n");navigator.clipboard.writeText(txt).then(function(){protoToast("Скопировано списком ("+rows.length+")");});}
function protoCsvCell(s){s=String(s==null?"":s);return /[",\n;]/.test(s)?"\""+s.replace(/"/g,"\"\"")+"\"":s;}
function protoCSV(){var rows=protoSelected();if(!rows.length){protoToast("Нет выбранных строк");return;}var head=["Задача","Исполнитель","Срок","Проект","Цитата"];var csv="﻿"+[head].concat(rows).map(function(r){return r.map(protoCsvCell).join(";");}).join("\r\n");var blob=new Blob([csv],{type:"text/csv;charset=utf-8"});var a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="Задачи_из_протокола_"+new Date().toISOString().slice(0,10)+".csv";a.click();URL.revokeObjectURL(a.href);protoToast("Файл выгружен");}
function protoSample(){document.getElementById("pr-text").value="Протокол оперативного совещания ОПУ от 07.07.2026.\n\n1. Обсудили статус интеграции с Битрикс24. Решено ускорить.\n2. Иванову подготовить ТЗ на доработку CRM-процесса Жалобы до пятницы.\n3. Петрова готовит отчёт по закрытым задачам за июнь, срок - к следующему совещанию.\n4. Сидоров К.П. согласует с финансовым отделом формулу расчёта до конца месяца.\n5. Общее замечание: качество данных в реестре надо повышать.";document.getElementById("pr-date").value="2026-07-07";protoToast("Пример вставлен - нажмите Извлечь задачи");}
'@) | Out-Null
$sb.AppendLine(@'
function mcEsc(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}
function mcStatus(t,cls){var el=document.getElementById("mc-status");if(el){el.textContent=t||"";el.className="mc-status"+(cls?" "+cls:"");}}
function mcToday(){var d=new Date();d.setHours(0,0,0,0);return d;}
function mcParse(s){var m=String(s||"").match(/^(\d{4})-(\d{2})-(\d{2})/);return m?new Date(+m[1],+m[2]-1,+m[3]):null;}
function mcIso(d){return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");}
function mcFmt(d){return String(d.getDate()).padStart(2,"0")+"."+String(d.getMonth()+1).padStart(2,"0")+"."+d.getFullYear();}
function mcWdOffset(d){if(!d)return 0;var t=mcToday();if(d<=t)return 0;var c=0,cur=new Date(t);while(cur<d){cur.setDate(cur.getDate()+1);var w=cur.getDay();if(w!==0&&w!==6)c++;}return c;}
function mcAddWd(n){n=Math.max(0,Math.ceil(n));var cur=new Date(mcToday()),a=0;while(a<n){cur.setDate(cur.getDate()+1);var w=cur.getDay();if(w!==0&&w!==6)a++;}return cur;}
function mcTri(o,m,p){if(p<=o)return m;var u=Math.random(),c=(m-o)/(p-o);return u<c?o+Math.sqrt(u*(p-o)*(m-o)):p-Math.sqrt((1-u)*(p-o)*(p-m));}
function mcSortNum(a){return a.slice().sort(function(x,y){return x-y;});}
function mcPct(sorted,q){if(!sorted.length)return 0;var i=Math.min(sorted.length-1,Math.floor(q*sorted.length));return sorted[i];}
function mcPctColor(p){return p>=0.8?"#16a34a":(p>=0.5?"#d97706":"#dc2626");}
function mcEndOfQuarter(){var d=new Date();var q=Math.floor(d.getMonth()/3);return new Date(d.getFullYear(),q*3+3,0);}
function mcInit(){var t=document.getElementById("mc-target");if(t&&!t.value)t.value=mcIso(mcEndOfQuarter());var sel=document.getElementById("mc-proj");if(sel&&sel.options.length<=1&&typeof SD!=="undefined"&&SD.projs){SD.projs.forEach(function(p){var o=document.createElement("option");o.value=p;o.textContent=p;sel.appendChild(o);});}}
function mcRunSim(){
  var all=(typeof SD!=="undefined"&&SD&&SD.tasks)?SD.tasks:[];
  var projSel=document.getElementById("mc-proj")?document.getElementById("mc-proj").value:"";
  var src=projSel?all.filter(function(t){return t.proj===projSel;}):all;
  if(!src.length){mcStatus(projSel?("В проекте «"+projSel+"» нет задач с исполнителем и сроком."):"Нет задач с исполнителем и сроком для прогноза.","err");document.getElementById("mc-result").style.display="none";return;}
  var opt=1-(parseFloat(document.getElementById("mc-opt").value)||0)/100;
  var pess=1+(parseFloat(document.getElementById("mc-pess").value)||0)/100;
  if(opt>1)opt=1;if(opt<0.1)opt=0.1;if(pess<1)pess=1;
  var def=parseFloat(document.getElementById("mc-def").value)||1;
  var runs=Math.max(500,Math.min(20000,parseInt(document.getElementById("mc-runs").value)||3000));
  var targetD=mcParse(document.getElementById("mc-target").value);var targetOff=mcWdOffset(targetD);
  var tasks=src.map(function(t){var hasEst=(parseFloat(t.tz)>0);return {id:t.id,name:t.name,exec:t.exec||"(без исполнителя)",est:hasEst?parseFloat(t.tz):def,dloff:mcWdOffset(mcParse(t.dl)),dl:t.dl,hits:0,noEst:!hasEst};});
  var missing=tasks.filter(function(t){return t.noEst;}).length;
  var byExec={};tasks.forEach(function(t){(byExec[t.exec]=byExec[t.exec]||[]).push(t);});
  var execNames=Object.keys(byExec);
  execNames.forEach(function(e){byExec[e].sort(function(a,b){return a.dloff-b.dloff;});});
  var portfolio=[],execFin={};execNames.forEach(function(e){execFin[e]=[];});
  var allMetCnt=0;
  for(var r=0;r<runs;r++){
    var maxFin=0,allMet=true;
    for(var ei=0;ei<execNames.length;ei++){
      var arr=byExec[execNames[ei]],cum=0;
      for(var i=0;i<arr.length;i++){var tk=arr[i];cum+=mcTri(tk.est*opt,tk.est,tk.est*pess);if(cum<=tk.dloff)tk.hits++;else allMet=false;}
      execFin[execNames[ei]].push(cum);if(cum>maxFin)maxFin=cum;
    }
    portfolio.push(maxFin);if(allMet)allMetCnt++;
  }
  portfolio=mcSortNum(portfolio);
  var pTarget=targetOff>0?(portfolio.filter(function(x){return x<=targetOff;}).length/runs):0;
  mcRenderResult({tasks:tasks,runs:runs,missing:missing,def:def,proj:projSel,pDeadlines:allMetCnt/runs,pTarget:pTarget,targetD:targetD,portfolio:portfolio,execFin:execFin,byExec:byExec});
  mcStatus((projSel?("Проект: "+projSel):"Весь портфель")+" · "+runs+" прогонов, задач "+tasks.length+(missing?(", без оценки "+missing+" (по "+def+" дн.)"):""));
}
function mcRenderResult(res){
  document.getElementById("mc-result").style.display="block";
  var expSlip=0;res.tasks.forEach(function(t){expSlip+=(1-t.hits/res.runs);});
  var total=res.tasks.length;var slipShare=total?expSlip/total:0;
  var sev=slipShare>0.4?"red":(slipShare>0.15?"amber":"green");
  var execEff=Object.keys(res.byExec).map(function(e){var eff=res.byExec[e].reduce(function(s,t){return s+t.est;},0);return {e:e,eff:eff,date:mcAddWd(mcPct(mcSortNum(res.execFin[e]),0.8))};}).sort(function(a,b){return b.eff-a.eff;});
  var top=execEff.slice(0,2);
  var scope=res.proj?("Проект «"+res.proj+"»"):"Портфель";
  var head=sev==="red"?"перегружено, сроки под угрозой":(sev==="amber"?"впритык, есть риски":"в целом реалистично");
  var line1="Из "+total+" задач ожидаемо не уложатся в срок примерно "+Math.round(expSlip)+" ("+Math.round(slipShare*100)+"%).";
  var line1b=res.targetD?(" К "+mcFmt(res.targetD)+" весь объём закрыт с вероятностью "+Math.round(res.pTarget*100)+"%."):"";
  var where=top.length?("Наибольшая нагрузка: "+top.map(function(x){return x.e+" (~"+Math.round(x.eff)+" дн. работы, готово к "+mcFmt(x.date)+")";}).join(", ")+"."):"";
  var todo=sev==="red"?"Что делать: перераспределить задачи перегруженных исполнителей или согласовать сдвиг сроков. Самые горящие задачи - в таблице «Задачи под наибольшим риском»."
    :sev==="amber"?"Что делать: держать на контроле перегруженных исполнителей и задачи из красной зоны; часть сроков, возможно, придётся подвинуть."
    :"Что делать: портфель выполним. Точечно следить за задачами из списка риска.";
  var icon=sev==="green"?"&#9989;":"&#9888;";
  document.getElementById("mc-verdict").innerHTML="<div class=\"mc-v mc-v-"+sev+"\"><button class=\"mc-v-copy\" onclick=\"mcCopyVerdict()\">Скопировать</button><div class=\"mc-v-head\">"+icon+" "+mcEsc(scope)+": "+head+"</div><div class=\"mc-v-body\">"+mcEsc(line1+line1b)+"</div>"+(where?"<div class=\"mc-v-body\">"+mcEsc(where)+"</div>":"")+"<div class=\"mc-v-todo\">"+mcEsc(todo)+"</div></div>";
  window.mcVerdictText=scope+" - прогноз (Монте-Карло):\nВердикт: "+head+".\n"+line1+line1b+"\n"+(where?where+"\n":"")+todo+"\n(Расчёт: "+res.runs+" прогонов"+(res.proj?", проект в изоляции":"")+".)";
  var pd=document.getElementById("mc-p-deadlines");pd.textContent=Math.round(res.pDeadlines*100)+"%";pd.style.color=mcPctColor(res.pDeadlines);
  document.getElementById("mc-target-lbl").textContent="портфель закрыт к "+(res.targetD?mcFmt(res.targetD):"—");
  var pt=document.getElementById("mc-p-target");pt.textContent=res.targetD?(Math.round(res.pTarget*100)+"%"):"—";if(res.targetD)pt.style.color=mcPctColor(res.pTarget);
  document.getElementById("mc-p80").textContent=mcFmt(mcAddWd(mcPct(res.portfolio,0.8)));
  mcDrawHist(res.portfolio,mcPct(res.portfolio,0.5),mcPct(res.portfolio,0.8));
  var rk=res.tasks.slice().sort(function(a,b){return a.hits-b.hits;});
  var rrows=rk.slice(0,15).map(function(t){var p=t.hits/res.runs;return "<tr><td>"+mcEsc(t.name)+"</td><td>"+mcEsc(t.exec)+"</td><td>"+mcEsc(t.dl)+"</td><td style=\"color:"+mcPctColor(p)+";font-weight:700;text-align:right\">"+Math.round(p*100)+"%</td></tr>";}).join("");
  document.getElementById("mc-risk").innerHTML="<thead><tr><th>Задача</th><th>Исп.</th><th>Срок</th><th style=\"text-align:right\">Успеть</th></tr></thead><tbody>"+rrows+"</tbody>";
  var execs=Object.keys(res.byExec).map(function(e){var fin=mcSortNum(res.execFin[e]);var p80=mcPct(fin,0.8);var minHit=Math.min.apply(null,res.byExec[e].map(function(t){return t.hits/res.runs;}));return {e:e,n:res.byExec[e].length,p80:p80,date:mcAddWd(p80),minHit:minHit};}).sort(function(a,b){return b.p80-a.p80;});
  var brows=execs.map(function(x){return "<tr><td>"+mcEsc(x.e)+"</td><td style=\"text-align:center\">"+x.n+"</td><td>"+mcFmt(x.date)+"</td><td style=\"color:"+mcPctColor(x.minHit)+";text-align:right;font-weight:700\">"+Math.round(x.minHit*100)+"%</td></tr>";}).join("");
  document.getElementById("mc-bottleneck").innerHTML="<thead><tr><th>Исполнитель</th><th style=\"text-align:center\">Задач</th><th>Готово к (P80)</th><th style=\"text-align:right\">Худшая</th></tr></thead><tbody>"+brows+"</tbody>";
  document.getElementById("mc-note").textContent="Как считается: длительность каждой задачи - случайная величина по трёхточечной оценке (лучший / введённая оценка трудозатрат / худший). Исполнитель делает задачи по очереди от ближайшего дедлайна, выходные исключены. «Успеть» - доля прогонов, где задача закрыта не позже своего срока. Это приближение: зависимости между задачами и параллельная работа не учитываются."+(res.proj?" Выбран один проект - расчёт в изоляции: занятость этих же людей в других проектах и заявках здесь не учитывается, поэтому реальные сроки могут быть позже.":"");
}
function mcCopyVerdict(){if(!window.mcVerdictText)return;navigator.clipboard.writeText(window.mcVerdictText).then(function(){var el=document.getElementById("toast");if(el){el.textContent="Вывод скопирован";el.style.display="block";setTimeout(function(){el.style.display="none";},2000);}});}
function mcDrawHist(portfolio,p50off,p80off){
  var cv=document.getElementById("mc-hist");if(!cv)return;var ctx=cv.getContext("2d");var W=cv.width,H=cv.height;ctx.clearRect(0,0,W,H);
  if(!portfolio.length)return;
  var pad={l:34,r:12,t:16,b:34};var minOff=portfolio[0],maxOff=portfolio[portfolio.length-1];if(maxOff<=minOff)maxOff=minOff+1;
  var bins=24,bw=(maxOff-minOff)/bins;var counts=new Array(bins).fill(0);
  portfolio.forEach(function(x){var b=Math.min(bins-1,Math.floor((x-minOff)/bw));counts[b]++;});
  var maxC=Math.max.apply(null,counts)||1;var plotW=W-pad.l-pad.r,plotH=H-pad.t-pad.b;
  ctx.fillStyle="#93c5fd";
  for(var i=0;i<bins;i++){var bh=counts[i]/maxC*plotH;ctx.fillRect(pad.l+i/bins*plotW,pad.t+plotH-bh,plotW/bins-1,bh);}
  ctx.strokeStyle="#cbd5e1";ctx.beginPath();ctx.moveTo(pad.l,pad.t+plotH);ctx.lineTo(W-pad.r,pad.t+plotH);ctx.stroke();
  ctx.fillStyle="#64748b";ctx.font="10px Segoe UI";ctx.textAlign="center";
  for(var k=0;k<=4;k++){var off=minOff+(maxOff-minOff)*k/4;var xx=pad.l+k/4*plotW;ctx.fillText(mcFmt(mcAddWd(off)).slice(0,5),xx,H-16);}
  function vline(off,color,label){if(off<minOff||off>maxOff)return;var xx=pad.l+(off-minOff)/(maxOff-minOff)*plotW;ctx.strokeStyle=color;ctx.setLineDash([4,3]);ctx.beginPath();ctx.moveTo(xx,pad.t);ctx.lineTo(xx,pad.t+plotH);ctx.stroke();ctx.setLineDash([]);ctx.fillStyle=color;ctx.textAlign="center";ctx.fillText(label,xx,pad.t-4);}
  vline(p50off,"#0ea5e9","P50");vline(p80off,"#d97706","P80");
}
'@) | Out-Null
$sb.AppendLine('function switchTab(t){["dash","rem","ai","plan","plnr","proto","mc","sver"].forEach(function(id){document.getElementById("tab-"+id).style.display=t===id?"block":"none";var btn=document.getElementById("tab-btn-"+id);if(btn)btn.classList.toggle("active",t===id);});if(t==="plan")planRender();if(t==="plnr"){var f=document.getElementById("plnr-frame");if(f&&!f.src&&f.getAttribute("data-src"))f.src=f.getAttribute("data-src");}if(t==="proto")protoInit();if(t==="mc"){mcInit();if(!window._mcRan){window._mcRan=true;mcRunSim();}}if(t==="sver")sverRender();}') | Out-Null
$sb.AppendLine('let planOffset=0;var planTaskMap={};') | Out-Null
$sb.AppendLine('function planWinStart(){const t=new Date();t.setHours(0,0,0,0);const dow=(t.getDay()+6)%7;const mon=new Date(t);mon.setDate(t.getDate()-dow);mon.setDate(mon.getDate()+planOffset*14);return mon;}') | Out-Null
$sb.AppendLine('function planShift(dir){if(dir===0)planOffset=0;else planOffset+=dir;planRender();}') | Out-Null
$sb.AppendLine('function planFmtD(d){return d.getDate()+"."+("0"+(d.getMonth()+1)).slice(-2);}') | Out-Null
$sb.AppendLine('function planEsc(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}') | Out-Null
$sb.AppendLine('function planRender(){const start=planWinStart();const days=[];for(let i=0;i<14;i++){const d=new Date(start);d.setDate(start.getDate()+i);days.push(d);}const today=new Date();today.setHours(0,0,0,0);const dn=["Пн","Вт","Ср","Чт","Пт","Сб","Вс"];document.getElementById("plan-range").textContent=planFmtD(days[0])+" - "+planFmtD(days[13]);const byExec={};SD.tasks.forEach(function(tk){if(!tk.exec)return;const dl=new Date(tk.dl);dl.setHours(0,0,0,0);if(dl<days[0]||dl>days[13])return;if(!byExec[tk.exec])byExec[tk.exec]=[];byExec[tk.exec].push(tk);});const execs=Object.keys(byExec);const grid=document.getElementById("plan-grid");if(execs.length===0){grid.innerHTML="<div class=\"plan-empty\">Нет задач с дедлайнами в этом периоде. Листайте кнопками Вперёд / Назад.</div>";return;}let avail=0;days.forEach(function(d){const wd=d.getDay();if(wd!==0&&wd!==6&&d>=today)avail++;});if(avail===0)avail=SD.avail;const load={};execs.forEach(function(nm){let tz=0,cnt=0,wtz=0;byExec[nm].forEach(function(tk){cnt++;if(tk.tz>0){tz+=tk.tz;wtz++;}});const cov=cnt>0?Math.round(wtz/cnt*100):0;load[nm]={tz:Math.round(tz*10)/10,cnt:cnt,cov:cov,coeff:avail>0&&tz>0?Math.round(tz/avail*100)/100:null};});execs.sort(function(a,b){return ((load[b].coeff||0)-(load[a].coeff||0))||(load[b].cnt-load[a].cnt);});let idx=0;planTaskMap={};let h="<table class=\"plan-grid\"><thead><tr><th class=\"plan-exec-col\">Исполнитель</th>";days.forEach(function(d){const wd=d.getDay();const isW=wd===0||wd===6;const isT=d.getTime()===today.getTime();const cls=isT?"today":(isW?"wknd":"");h+="<th class=\""+cls+"\">"+dn[(wd+6)%7]+"<span class=\"plan-day-num\">"+d.getDate()+"</span></th>";});h+="<th>Загрузка</th></tr></thead><tbody>";execs.forEach(function(nm){h+="<tr><td class=\"plan-exec-col\">"+planEsc(nm)+"</td>";days.forEach(function(d){const wd=d.getDay();const isW=wd===0||wd===6;const isT=d.getTime()===today.getTime();const cls="plan-cell"+(isT?" today":(isW?" wknd":""));h+="<td class=\""+cls+"\">";byExec[nm].forEach(function(tk){const dl=new Date(tk.dl);dl.setHours(0,0,0,0);if(dl.getTime()!==d.getTime())return;let bc=tk.src==="proj"?"proj":"req";if(dl<today)bc="over";else if((dl-today)/86400000<=2)bc="warn";const lbl=tk.name?tk.name:(tk.id||"задача");planTaskMap[idx]=tk;h+="<span class=\"plan-blk "+bc+"\" data-i=\""+idx+"\">"+planEsc(lbl)+"</span>";idx++;});h+="</td>";});const l=load[nm];const bcol=l.coeff===null?"#cbd5e1":(l.coeff>1?"#ef4444":(l.coeff>0.8?"#f59e0b":"#22c55e"));const pct=l.coeff===null?0:Math.min(100,Math.round(l.coeff*100));const ktxt=l.coeff===null?"нет данных":"К="+l.coeff;const covn=l.cov<100&&l.coeff!==null?" · данные "+l.cov+"%":"";h+="<td class=\"plan-load\"><div class=\"plan-load-txt\">"+l.cnt+" зад. · "+ktxt+covn+"</div><div class=\"plan-load-bar\"><div class=\"plan-load-fill\" style=\"width:"+pct+"%;background:"+bcol+"\"></div></div></td></tr>";});h+="</tbody></table>";grid.innerHTML=h;}') | Out-Null
$sb.AppendLine('function planTip(e){const tip=document.getElementById("plan-tip");const b=e.target&&e.target.closest?e.target.closest(".plan-blk"):null;if(!b){tip.style.display="none";tip.dataset.i="";return;}const tk=planTaskMap[b.getAttribute("data-i")];if(!tk){tip.style.display="none";return;}if(tip.dataset.i!==b.getAttribute("data-i")){tip.dataset.i=b.getAttribute("data-i");const src=tk.src==="proj"?"Проект":"Заявка";const dl=new Date(tk.dl);const dd=("0"+dl.getDate()).slice(-2)+"."+("0"+(dl.getMonth()+1)).slice(-2)+"."+dl.getFullYear();const tzt=tk.tz>0?tk.tz+" дн.":"не указаны";let c="<b>"+planEsc(tk.name||"Без названия")+"</b><br><span class=\"pt-meta\">"+src+(tk.id&&tk.id!=="НОВАЯ"?" ["+planEsc(tk.id)+"]":"")+"</span><br>";if(tk.proj)c+="<span class=\"pt-meta\">Проект: "+planEsc(tk.proj)+"</span><br>";c+="Исполнитель: "+planEsc(tk.exec)+"<br>Дедлайн: "+dd+"<br>Трудозатраты: "+tzt;tip.innerHTML=c;}tip.style.display="block";let x=e.clientX+14,y=e.clientY+14;const w=tip.offsetWidth,ht=tip.offsetHeight;if(x+w>window.innerWidth-10)x=e.clientX-w-14;if(y+ht>window.innerHeight-10)y=e.clientY-ht-14;tip.style.left=x+"px";tip.style.top=y+"px";}') | Out-Null
$sb.AppendLine('function planTipHide(){var t=document.getElementById("plan-tip");if(t){t.style.display="none";t.dataset.i="";}}') | Out-Null
$sb.AppendLine('function remCheck(cb){const row=cb.closest("tr");localStorage.setItem("rem_"+row.dataset.rkey,cb.checked?"1":"0");row.style.opacity=cb.checked?"0.3":"1";checkMissedDone();}') | Out-Null
$sb.AppendLine('function initRemChecks(){document.querySelectorAll(".rem-row").forEach(function(row){if(localStorage.getItem("rem_"+row.dataset.rkey)==="1"){row.querySelector(".rem-chk").checked=true;row.style.opacity="0.3";}});}') | Out-Null
$sb.AppendLine('function renderGantt(tasks){') | Out-Null
$sb.AppendLine('  const today=new Date();today.setHours(0,0,0,0);') | Out-Null
$sb.AppendLine('  const wDl=tasks.filter(t=>t.dl);') | Out-Null
$sb.AppendLine('  if(!wDl.length)return"<div class=''g-empty''>Нет задач с дедлайнами</div>";') | Out-Null
$sb.AppendLine('  const ms=wDl.map(t=>new Date(t.dl).getTime());') | Out-Null
$sb.AppendLine('  const tMin=new Date(Math.min(today.getTime()-7*864e5,Math.min(...ms)));') | Out-Null
$sb.AppendLine('  const tMax=new Date(Math.max(...ms)+14*864e5);') | Out-Null
$sb.AppendLine('  const span=tMax-tMin;') | Out-Null
$sb.AppendLine('  const pct=d=>Math.max(0,Math.min(100,(d-tMin)/span*100));') | Out-Null
$sb.AppendLine('  const todayP=pct(today);') | Out-Null
$sb.AppendLine('  const days=span/864e5;const labelEvery=days>90?14:days>60?7:days>30?3:1;') | Out-Null
$sb.AppendLine('  let lbls="";let grid="";let d=new Date(tMin);let dayIdx=0;') | Out-Null
$sb.AppendLine('  while(d<=tMax){const p=pct(d);const lft=`calc(${p.toFixed(1)}%+${(200-2*p).toFixed(1)}px)`;grid+=`<div style="position:absolute;top:0;bottom:0;left:${lft};width:1px;background:#f1f5f9;pointer-events:none"></div>`;if(dayIdx%labelEvery===0){const s=d.toLocaleDateString("ru-RU",{day:"numeric",month:"short"});lbls+=`<span class="g-lbl" style="left:${p}%">${s}</span>`;}else{lbls+=`<span class="g-tick" style="left:${p}%"></span>`;}d=new Date(d.getTime()+864e5);dayIdx++;}') | Out-Null
$sb.AppendLine('  const sorted=[...tasks].sort((a,b)=>{if(!a.dl&&!b.dl)return 0;if(!a.dl)return 1;if(!b.dl)return-1;return new Date(a.dl)-new Date(b.dl);});') | Out-Null
$sb.AppendLine('  let rows="";') | Out-Null
$sb.AppendLine('  sorted.forEach(t=>{') | Out-Null
$sb.AppendLine('    const nm=t.n.length>38?t.n.slice(0,36)+"…":t.n;') | Out-Null
$sb.AppendLine('    if(!t.dl){rows+=`<div class="g-row"><div class="g-nm" title="${t.n}">${nm}</div><div class="g-area"><span class="g-nodl">нет дедлайна</span></div></div>`;return;}') | Out-Null
$sb.AppendLine('    const dl=new Date(t.dl);const dp=pct(dl);') | Out-Null
$sb.AppendLine('    const isOver=!t.done&&dl<today;const isUrg=!t.done&&!isOver&&(dl-today)/864e5<=7;') | Out-Null
$sb.AppendLine('    const cls=t.done?"gd-done":isOver?"gd-over":isUrg?"gd-urg":"gd-ok";') | Out-Null
$sb.AppendLine('    const dlFmt=dl.toLocaleDateString("ru-RU",{day:"numeric",month:"short"});') | Out-Null
$sb.AppendLine('    const tip=`${t.n}\nДедлайн: ${dlFmt}${t.done?" (выполнено)":""}`;') | Out-Null
$sb.AppendLine('    rows+=`<div class="g-row" title="${tip}"><div class="g-nm">${nm}</div><div class="g-area"><div class="g-line ${cls}" style="width:${dp}%"></div><span class="g-dlbl ${cls}" style="left:${dp}%">${dlFmt}</span></div></div>`;') | Out-Null
$sb.AppendLine('  });') | Out-Null
$sb.AppendLine('  const tlLeft=`calc(${todayP.toFixed(1)}% + ${(200-2*todayP).toFixed(1)}px)`;') | Out-Null
$sb.AppendLine('  return`<div class="g-dates">${lbls}</div><div class="g-rows"><div style="position:absolute;top:0;left:0;right:0;bottom:0;pointer-events:none;overflow:hidden">${grid}</div><div class="g-today-line" style="left:${tlLeft}"><span class="g-today-lbl">сегодня</span></div>${rows}</div>`;') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function initPanels(){') | Out-Null
$sb.AppendLine('  const hidden=JSON.parse(localStorage.getItem("hp")||"[]");') | Out-Null
$sb.AppendLine('  document.querySelectorAll(".panel").forEach((p,i)=>{') | Out-Null
$sb.AppendLine('    p.dataset.pid="p"+i;p.dataset.pnm=(p.querySelector("h2")?.textContent?.trim()||"Панель "+(i+1));') | Out-Null
$sb.AppendLine('    const btn=document.createElement("button");btn.className="p-hide no-print";btn.title="Скрыть панель";btn.innerHTML="&times;";') | Out-Null
$sb.AppendLine('    btn.onclick=()=>hidePanel(p);p.appendChild(btn);') | Out-Null
$sb.AppendLine('    if(hidden.includes("p"+i))p.style.display="none";') | Out-Null
$sb.AppendLine('  });') | Out-Null
$sb.AppendLine('  updDock();') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function hidePanel(p){p.style.display="none";const h=JSON.parse(localStorage.getItem("hp")||"[]");if(!h.includes(p.dataset.pid))h.push(p.dataset.pid);localStorage.setItem("hp",JSON.stringify(h));updDock();}') | Out-Null
$sb.AppendLine('function restorePanel(pid){const p=document.querySelector("[data-pid=''"+pid+"'']");if(p)p.style.display="";const h=JSON.parse(localStorage.getItem("hp")||"[]").filter(x=>x!==pid);localStorage.setItem("hp",JSON.stringify(h));updDock();}') | Out-Null
$sb.AppendLine('function updDock(){') | Out-Null
$sb.AppendLine('  const dock=document.getElementById("restore-dock");') | Out-Null
$sb.AppendLine('  const list=document.getElementById("rd-list");') | Out-Null
$sb.AppendLine('  const h=JSON.parse(localStorage.getItem("hp")||"[]");') | Out-Null
$sb.AppendLine('  if(!h.length){dock.style.display="none";if(list)list.style.display="none";return;}') | Out-Null
$sb.AppendLine('  dock.style.display="block";') | Out-Null
$sb.AppendLine('  document.getElementById("rd-lbl").textContent="Скрытые ("+h.length+")";') | Out-Null
$sb.AppendLine('  list.innerHTML="";') | Out-Null
$sb.AppendLine('  h.forEach(pid=>{const p=document.querySelector("[data-pid=''"+pid+"'']");const nm=p?p.dataset.pnm:pid;const b=document.createElement("button");b.style.cssText="display:block;width:100%;text-align:left;padding:7px 10px;border:none;background:none;cursor:pointer;font-size:12.5px;color:#334155;border-radius:6px";b.textContent=nm;b.onmouseover=()=>{b.style.background="#f1f5f9";};b.onmouseout=()=>{b.style.background="none";};b.onclick=()=>{restorePanel(pid);};list.appendChild(b);});') | Out-Null
$sb.AppendLine('  if(h.length>1){const sep=document.createElement("div");sep.style.cssText="border-top:1px solid #e2e8f0;margin:4px 0";list.appendChild(sep);const all=document.createElement("button");all.style.cssText="display:block;width:100%;text-align:left;padding:7px 10px;border:none;background:none;cursor:pointer;font-size:12.5px;color:#2563eb;border-radius:6px;font-weight:600";all.textContent="Показать все";all.onmouseover=()=>{all.style.background="#eff6ff";};all.onmouseout=()=>{all.style.background="none";};all.onclick=()=>{h.slice().forEach(pid=>restorePanel(pid));list.style.display="none";};list.appendChild(all);}') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function toggleRestoreDock(){const l=document.getElementById("rd-list");if(l)l.style.display=l.style.display==="block"?"none":"block";}') | Out-Null
$sb.AppendLine('function copyGroup(btn){') | Out-Null
$sb.AppendLine('  const grp=btn.closest(".grp");') | Out-Null
$sb.AppendLine('  const title=grp.querySelector(".grp-hdr span")?.textContent?.trim()||"";') | Out-Null
$sb.AppendLine('  const lines=[];') | Out-Null
$sb.AppendLine('  grp.querySelectorAll(".task").forEach(t=>{') | Out-Null
$sb.AppendLine('    if(t.style.display==="none")return;') | Out-Null
$sb.AppendLine('    const id=t.dataset.id||"",nm=t.dataset.name||"",ex=t.dataset.exec||"",cm=(t.dataset.comm||"").replace(/\s+/g," ").trim();') | Out-Null
$sb.AppendLine('    lines.push(id+" - "+nm+(ex?" ("+ex+")":"")+(cm?" · "+cm:""));') | Out-Null
$sb.AppendLine('  });') | Out-Null
$sb.AppendLine('  if(!lines.length)return;') | Out-Null
$sb.AppendLine('  navigator.clipboard.writeText(title+":\n"+lines.join("\n")).then(()=>{') | Out-Null
$sb.AppendLine('    const t=document.getElementById("toast");') | Out-Null
$sb.AppendLine('    t.textContent="Список скопирован";t.style.display="block";') | Out-Null
$sb.AppendLine('    setTimeout(()=>{t.style.display="none";},2200);') | Out-Null
$sb.AppendLine('  });') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('let arTimer=null,arLeft=0;') | Out-Null
$sb.AppendLine('function initAR(){') | Out-Null
$sb.AppendLine('  const ar=new URLSearchParams(location.search).get("autorefresh");') | Out-Null
$sb.AppendLine('  const min=parseInt(ar||localStorage.getItem("ar_min")||"5");') | Out-Null
$sb.AppendLine('  if(ar||localStorage.getItem("ar_on")==="1")startAR(min);') | Out-Null
$sb.AppendLine('  else updateArBtn();') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function startAR(min){') | Out-Null
$sb.AppendLine('  if(arTimer)clearInterval(arTimer);') | Out-Null
$sb.AppendLine('  arLeft=min*60;localStorage.setItem("ar_on","1");localStorage.setItem("ar_min",min);') | Out-Null
$sb.AppendLine('  arTimer=setInterval(()=>{arLeft--;updateArBtn();if(arLeft<=0)hardReload();},1000);') | Out-Null
$sb.AppendLine('  updateArBtn();') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function stopAR(){clearInterval(arTimer);arTimer=null;localStorage.setItem("ar_on","0");updateArBtn();}') | Out-Null
$sb.AppendLine('function toggleAR(){if(arTimer)stopAR();else startAR(parseInt(localStorage.getItem("ar_min")||"5"));}') | Out-Null
$sb.AppendLine('function updateArBtn(){') | Out-Null
$sb.AppendLine('  const b=document.getElementById("btn-ar");if(!b)return;') | Out-Null
$sb.AppendLine('  const t=b.querySelector(".ar-t");') | Out-Null
$sb.AppendLine('  if(arTimer){') | Out-Null
$sb.AppendLine('    const m=Math.floor(arLeft/60),s=arLeft%60;') | Out-Null
$sb.AppendLine('    t.textContent=m+":"+(s<10?"0":"")+s;') | Out-Null
$sb.AppendLine('    b.classList.add("ar-active");b.title="Автообновление включено — нажмите для отключения";') | Out-Null
$sb.AppendLine('  }else{t.textContent="";b.classList.remove("ar-active");b.title="Автообновление каждые 5 мин — нажмите для включения";}') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function hardReload(){location.href=location.pathname+"?r="+Date.now();}') | Out-Null
$sb.AppendLine('document.addEventListener("keydown",function(e){') | Out-Null
$sb.AppendLine('  if(e.key==="Escape")closeGantt();') | Out-Null
$sb.AppendLine('  if(e.ctrlKey&&e.shiftKey&&e.key==="R"){e.preventDefault();hardReload();}') | Out-Null
$sb.AppendLine('});') | Out-Null
$sb.AppendLine('function globalSearch(){') | Out-Null
$sb.AppendLine('  const raw=document.getElementById("g-srch").value;') | Out-Null
$sb.AppendLine('  const q=raw.toLowerCase().trim();') | Out-Null
$sb.AppendLine('  document.getElementById("g-srch-clear").classList.toggle("on",raw.length>0);') | Out-Null
$sb.AppendLine('  // Синхронизируем с фильтром Этапов') | Out-Null
$sb.AppendLine('  document.getElementById("srch").value=raw; filterTasks();') | Out-Null
$sb.AppendLine('  // Прогресс проектов') | Out-Null
$sb.AppendLine('  document.querySelectorAll(".proj").forEach(function(proj){') | Out-Null
$sb.AppendLine('    if(!q){proj.style.display="";proj.querySelectorAll(".trow").forEach(function(r){r.style.display="";});return;}') | Out-Null
$sb.AppendLine('    const st=(proj.querySelector("summary")||{textContent:""}).textContent.toLowerCase();') | Out-Null
$sb.AppendLine('    let n=0; proj.querySelectorAll(".trow").forEach(function(r){const ok=r.textContent.toLowerCase().includes(q)||st.includes(q);r.style.display=ok?"":"none";if(ok)n++;});') | Out-Null
$sb.AppendLine('    proj.style.display=(n>0||st.includes(q))?"":"none";') | Out-Null
$sb.AppendLine('    if(n>0&&!proj.open)proj.setAttribute("open","");') | Out-Null
$sb.AppendLine('  });') | Out-Null
$sb.AppendLine('  // Таблицы просрочки и дедлайнов') | Out-Null
$sb.AppendLine('  ["ov-tbl","up-tbl"].forEach(function(id){') | Out-Null
$sb.AppendLine('    const tbl=document.getElementById(id);if(!tbl)return;') | Out-Null
$sb.AppendLine('    tbl.querySelectorAll("tbody tr").forEach(function(r){') | Out-Null
$sb.AppendLine('      if(r.querySelector("[colspan]"))return;') | Out-Null
$sb.AppendLine('      r.style.display=(!q||r.textContent.toLowerCase().includes(q))?"":"none";') | Out-Null
$sb.AppendLine('    });') | Out-Null
$sb.AppendLine('  });') | Out-Null
$sb.AppendLine('  // Счётчик результатов') | Out-Null
$sb.AppendLine('  const cnt=document.getElementById("g-srch-cnt");') | Out-Null
$sb.AppendLine('  if(!q){cnt.textContent="";return;}') | Out-Null
$sb.AppendLine('  let total=0;') | Out-Null
$sb.AppendLine('  document.querySelectorAll(".grp .task").forEach(function(t){if(t.style.display!=="none")total++;});') | Out-Null
$sb.AppendLine('  document.querySelectorAll(".proj .trow").forEach(function(t){if(t.style.display!=="none")total++;});') | Out-Null
$sb.AppendLine('  ["ov-tbl","up-tbl"].forEach(function(id){const tbl=document.getElementById(id);if(!tbl)return;tbl.querySelectorAll("tbody tr").forEach(function(r){if(r.style.display!=="none"&&!r.querySelector("[colspan]"))total++;});});') | Out-Null
$sb.AppendLine('  cnt.textContent=total?"Найдено: "+total+" совпадений":"Ничего не найдено";') | Out-Null
$sb.AppendLine('  cnt.style.color=total?"#64748b":"#dc2626";') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function clearGsrch(){document.getElementById("g-srch").value="";globalSearch();}') | Out-Null
$sb.AppendLine('function toggleWatch(btn){') | Out-Null
$sb.AppendLine('  const id=btn.dataset.id;') | Out-Null
$sb.AppendLine('  const key="watch_"+id;') | Out-Null
$sb.AppendLine('  if(localStorage.getItem(key)==="1"){') | Out-Null
$sb.AppendLine('    localStorage.removeItem(key);localStorage.removeItem(key+"_data");') | Out-Null
$sb.AppendLine('    btn.innerHTML="&#9734;";btn.classList.remove("watching");btn.title="Добавить в наблюдение";') | Out-Null
$sb.AppendLine('  } else {') | Out-Null
$sb.AppendLine('    const task=btn.closest(".task");') | Out-Null
$sb.AppendLine('    const nm=(task.querySelector(".t-name")||{textContent:""}).textContent;') | Out-Null
$sb.AppendLine('    const ex=(task.querySelector(".t-exec")||{textContent:""}).textContent;') | Out-Null
$sb.AppendLine('    const dl=(task.querySelector(".t-dl")||{textContent:"",className:""});') | Out-Null
$sb.AppendLine('    const dept=task.dataset.dept||"";') | Out-Null
$sb.AppendLine('    localStorage.setItem(key,"1");') | Out-Null
$sb.AppendLine('    localStorage.setItem(key+"_data",JSON.stringify({id:id,nm:nm,ex:ex,dl:dl.textContent,dlCls:dl.className,dept:dept}));') | Out-Null
$sb.AppendLine('    btn.innerHTML="&#9733;";btn.classList.add("watching");btn.title="Убрать из наблюдения";') | Out-Null
$sb.AppendLine('  }') | Out-Null
$sb.AppendLine('  renderWatchPanel();') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function removeWatch(id){') | Out-Null
$sb.AppendLine('  localStorage.removeItem("watch_"+id);localStorage.removeItem("watch_"+id+"_data");') | Out-Null
$sb.AppendLine('  document.querySelectorAll(".watch-btn").forEach(function(b){if(b.dataset.id===id){b.innerHTML="&#9734;";b.classList.remove("watching");b.title="Добавить в наблюдение";}});') | Out-Null
$sb.AppendLine('  renderWatchPanel();') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function renderWatchPanel(){') | Out-Null
$sb.AppendLine('  const panel=document.getElementById("watch-panel");') | Out-Null
$sb.AppendLine('  const list=document.getElementById("watch-list");') | Out-Null
$sb.AppendLine('  const items=[];') | Out-Null
$sb.AppendLine('  for(let i=0;i<localStorage.length;i++){') | Out-Null
$sb.AppendLine('    const k=localStorage.key(i);') | Out-Null
$sb.AppendLine('    if(k&&k.startsWith("watch_")&&!k.endsWith("_data")&&localStorage.getItem(k)==="1"){') | Out-Null
$sb.AppendLine('      const ds=localStorage.getItem(k+"_data");') | Out-Null
$sb.AppendLine('      if(ds){try{items.push(JSON.parse(ds));}catch(e){}}') | Out-Null
$sb.AppendLine('    }') | Out-Null
$sb.AppendLine('  }') | Out-Null
$sb.AppendLine('  if(!items.length){panel.style.display="none";return;}') | Out-Null
$sb.AppendLine('  panel.style.display="";') | Out-Null
$sb.AppendLine('  list.innerHTML=items.map(function(it){') | Out-Null
$sb.AppendLine('    const dlHtml=it.dl?"<span class=\"t-dl "+it.dlCls+"\">"+it.dl+"</span>":"";') | Out-Null
$sb.AppendLine('    const exHtml=it.ex?"<span class=\"t-exec\">"+it.ex+"</span>":"";') | Out-Null
$sb.AppendLine('    const dpHtml=it.dept?"<span class=\"t-dept\">"+it.dept+"</span>":"";') | Out-Null
$sb.AppendLine('    return "<div class=\"watch-item\"><span class=\"t-id\">"+it.id+"</span>"+dpHtml+"<span class=\"t-name\">"+it.nm+"</span>"+exHtml+dlHtml+"<button class=\"watch-rm\" onclick=\"removeWatch(''"+it.id+"'')\">&#10005;</button></div>";') | Out-Null
$sb.AppendLine('  }).join("");') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function initWatchPanel(){') | Out-Null
$sb.AppendLine('  document.querySelectorAll(".watch-btn").forEach(function(b){') | Out-Null
$sb.AppendLine('    if(localStorage.getItem("watch_"+b.dataset.id)==="1"){b.innerHTML="&#9733;";b.classList.add("watching");b.title="Убрать из наблюдения";}') | Out-Null
$sb.AppendLine('  });') | Out-Null
$sb.AppendLine('  renderWatchPanel();') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine('function checkMissedDone(){') | Out-Null
$sb.AppendLine('  document.querySelectorAll(".rem-grp.rem-missed").forEach(function(grp){') | Out-Null
$sb.AppendLine('    const chks=grp.querySelectorAll(".rem-chk");') | Out-Null
$sb.AppendLine('    if(!chks.length)return;') | Out-Null
$sb.AppendLine('    const allDone=Array.from(chks).every(function(cb){return cb.checked;});') | Out-Null
$sb.AppendLine('    const hdr=grp.querySelector(".rem-grp-hdr");') | Out-Null
$sb.AppendLine('    if(!hdr)return;') | Out-Null
$sb.AppendLine('    const lbl=hdr.querySelector(".rem-grp-lbl");') | Out-Null
$sb.AppendLine('    if(allDone){hdr.style.background="#dcfce7";hdr.style.color="#15803d";if(lbl)lbl.textContent="Всё готово";}') | Out-Null
$sb.AppendLine('    else{hdr.style.background="";hdr.style.color="";if(lbl)lbl.textContent="Пропущено";}') | Out-Null
$sb.AppendLine('  });') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine("const AI_CTX=$aiContextJson;") | Out-Null
$sb.AppendLine('let aiHistory=[];let aiWorking=false;let aiVerbose=false;') | Out-Null
$sb.AppendLine('const AI_QUICK={"Угрозы срыва":"Какие проекты или заявки под угрозой срыва сроков? Кратко объясни причины.","Кто перегружен?":"Кто из исполнителей перегружен? Покажи у кого больше всего активных задач и просрочек.","Статус-отчёт":"Сформируй краткий статус-отчёт по портфелю для руководства в 4-6 предложениях.","Что назначить сегодня?":"Какие задачи нужно назначить исполнителям сегодня? Перечисли приоритетные.","Просрочки по отделам":"Как распределены просрочки по отделам? Где ситуация наиболее критична?"};') | Out-Null
$sb.AppendLine('function initAI(){const k=localStorage.getItem("opu_ai_key");document.getElementById("ai-setup").style.display=k?"none":"";document.getElementById("ai-chat").style.display=k?"":"none";if(k&&aiHistory.length===0){const saved=loadAIHistory();if(saved.length>0){aiHistory=saved;const wrap=document.getElementById("ai-msgs");const sep=document.createElement("div");sep.className="ai-hist-sep";sep.textContent="— история предыдущей сессии —";wrap.appendChild(sep);saved.forEach(function(m){addAIMsg(m.role,m.role==="assistant"?renderAIMd(m.content):m.content);});}}}') | Out-Null
$sb.AppendLine('function saveAIKey(){const k=document.getElementById("ai-key-inp").value.trim();if(!k){return;}localStorage.setItem("opu_ai_key",k);initAI();}') | Out-Null
$sb.AppendLine('function resetAIKey(){localStorage.removeItem("opu_ai_key");localStorage.removeItem("opu_ai_hist");aiHistory=[];document.getElementById("ai-msgs").innerHTML="";initAI();}') | Out-Null
$sb.AppendLine('function clearAIChat(){aiHistory=[];document.getElementById("ai-msgs").innerHTML="";localStorage.removeItem("opu_ai_hist");}') | Out-Null
$sb.AppendLine('function aiQuick(btn){const q=AI_QUICK[btn.textContent]||btn.textContent;document.getElementById("ai-inp").value=q;sendAI();}') | Out-Null
$sb.AppendLine('function aiScenarioFill(){var inp=document.getElementById("ai-inp");inp.value="Что если снять ";inp.focus();}') | Out-Null
$sb.AppendLine('function addAIMsg(role,html,isErr){const wrap=document.getElementById("ai-msgs");const div=document.createElement("div");div.className="ai-msg ai-"+role+(isErr?" ai-err":"");div.innerHTML=html;if(role==="assistant"&&!isErr){const rawHtml=html;const cp=document.createElement("button");cp.className="ai-copy-btn";cp.innerHTML="&#128203;";cp.title="Копировать ответ";cp.onclick=function(){const tmp=document.createElement("div");tmp.innerHTML=rawHtml;navigator.clipboard.writeText(tmp.innerText.trim());cp.innerHTML="&#10003;";setTimeout(function(){cp.innerHTML="&#128203;";},1500);};div.appendChild(cp);}wrap.appendChild(div);wrap.scrollTop=wrap.scrollHeight;return div;}') | Out-Null
$sb.AppendLine('function saveAIHistory(){localStorage.setItem("opu_ai_hist",JSON.stringify(aiHistory.slice(-12)));}') | Out-Null
$sb.AppendLine('function loadAIHistory(){try{var h=JSON.parse(localStorage.getItem("opu_ai_hist")||"[]");return Array.isArray(h)?h:[];}catch(e){return[];}}') | Out-Null
$sb.AppendLine('function toggleAIMode(){aiVerbose=!aiVerbose;var btn=document.getElementById("ai-mode-btn");if(btn){btn.textContent=aiVerbose?"Подробно":"Кратко";btn.classList.toggle("active",aiVerbose);}}') | Out-Null
$sb.AppendLine('function askAIAbout(btn){var t=btn.closest(".task");var id=t?t.dataset.id:"";var name=t?t.dataset.name:"";switchTab("ai");document.getElementById("ai-inp").value="Расскажи о заявке ["+id+"] "+name+". Каков статус, кто исполнитель, каковы сроки? Что нужно предпринять?";document.getElementById("ai-inp").focus();}') | Out-Null
$sb.AppendLine('function renderAIMd(t){t=t.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");t=t.replace(/\*\*(.+?)\*\*/g,"<b>$1</b>");const lines=t.split("\n");const out=[];let inUl=false;for(const ln of lines){if(/^[\-\*•] /.test(ln)){if(!inUl){out.push("<ul>");inUl=true;}out.push("<li>"+ln.replace(/^[\-\*•] /,"")+"</li>");}else{if(inUl){out.push("</ul>");inUl=false;}out.push(ln?"<span>"+ln+"</span><br>":"<br>");}}if(inUl)out.push("</ul>");return out.join("");}') | Out-Null
$sb.AppendLine('function sdDetect(msg){return /что если|а если|как будет|что будет|если снять|если передать|если сдвинуть|если добавить|если уйд/i.test(msg);}') | Out-Null
$sb.AppendLine('function sdFindExec(hint){const execs=[...new Set(SD.tasks.map(function(t){return t.exec;}).filter(Boolean))];const h=hint.toLowerCase();const stem=h.substring(0,Math.min(5,h.length));let f=execs.find(function(e){return e.toLowerCase()===h;});if(f)return f;f=execs.find(function(e){return e.toLowerCase().startsWith(stem);});if(f)return f;f=execs.find(function(e){return h.startsWith(e.toLowerCase().substring(0,Math.min(5,e.length)));});return f||null;}') | Out-Null
$sb.AppendLine('function sdCalcLoad(tasks,availMod){const today=new Date();today.setHours(0,0,0,0);const pe=new Date(today);pe.setDate(today.getDate()+SD.planDays);const load={};tasks.forEach(function(t){if(!t.exec)return;const dl=new Date(t.dl);dl.setHours(0,0,0,0);if(dl<today||dl>pe)return;if(!load[t.exec])load[t.exec]={tz:0,count:0,withTZ:0};load[t.exec].count++;if(t.tz>0){load[t.exec].tz+=t.tz;load[t.exec].withTZ++;}});const res={};Object.keys(load).forEach(function(nm){const e=load[nm];const avail=(availMod&&availMod[nm]!==undefined)?availMod[nm]:SD.avail;const cov=e.count>0?Math.round(e.withTZ/e.count*100):0;res[nm]={tz:Math.round(e.tz*10)/10,count:e.count,withTZ:e.withTZ,cov:cov,avail:avail,coeff:avail>0&&e.tz>0?Math.round(e.tz/avail*100)/100:null};});return res;}') | Out-Null
$sb.AppendLine('function sdApply(tasks,msg){const h=msg.toLowerCase();const mt=tasks.map(function(t){return Object.assign({},t);});let desc="Сценарий не распознан";let availMod={};const shiftM=h.match(/на\s+(\d+)\s*(?:дней|дня|день)/i);const shiftDays=shiftM?parseInt(shiftM[1]):0;const shiftDir=h.match(/назад|раньше/)?-1:1;const projM=h.match(/(?:проекта?\s+|в\s+проект[ае]\s+)([а-яёa-z0-9][а-яёa-z0-9\s\-]+?)(?:\s+и\s|\s+все\s|\s*$)/i);const projF=projM?projM[1].trim().toLowerCase():null;const isAbsence=/отпуск|командировк|отсутств/i.test(h);const isNewTask=/новую?\s+заявку?|добавить\s+заявку?|взять\s+в\s+работу?|добавить\s+задачу?/i.test(h);const fromW=(h.match(/(?:снять|убрать|освободить)\s+([а-яёa-z]+)/i)||[])[1];const toW=(h.match(/(?:передать|отдать|назначить|поставить)\s+([а-яёa-z]+)/i)||[])[1];if(isAbsence){const execW=(h.match(/([а-яёa-z]+)\s+(?:уйд|уход|в\s+отпуск|будет\s+в|идёт|идет|выход)/i)||h.match(/(?:если|когда)\s+([а-яёa-z]+)\s+/i)||[])[1];const exec=execW?sdFindExec(execW):null;const dateM=h.match(/с\s+(\d{1,2})[.\s](\d{1,2}).*?по\s+(\d{1,2})[.\s](\d{1,2})/i);let absD=0;if(dateM){const today=new Date();today.setHours(0,0,0,0);const pe=new Date(today);pe.setDate(today.getDate()+SD.planDays);const yr=today.getFullYear();let d=new Date(yr,parseInt(dateM[2])-1,parseInt(dateM[1]));const eD=new Date(yr,parseInt(dateM[4])-1,parseInt(dateM[3]));let ds=new Date(Math.max(d.getTime(),today.getTime()));const ed=new Date(Math.min(eD.getTime(),pe.getTime()));while(ds<=ed){if(ds.getDay()!==0&&ds.getDay()!==6)absD++;ds.setDate(ds.getDate()+1);}}else if(shiftDays>0){absD=Math.min(shiftDays,SD.avail);}if(exec){availMod[exec]=Math.max(0,SD.avail-absD);desc=exec+(absD>0?" в отпуске/отсутствии: -"+absD+" раб. дн. из "+SD.avail+" (доступно "+availMod[exec]+")":" — укажите даты (с ДД.ММ по ДД.ММ) или количество дней");}else{desc="Укажите фамилию исполнителя для сценария отпуска";}}else if(isNewTask){const execW=(h.match(/(?:у|для|назначить)\s+([а-яёa-z]+)/i)||h.match(/([а-яёa-z]+)\s+(?:взять|получит|добавить)/i)||[])[1];const exec=execW?sdFindExec(execW):null;const daysM=h.match(/(\d+(?:[,.]\d+)?)\s*(?:дней|дня|день)/i);const taskDays=daysM?parseFloat(daysM[1].replace(",",".")):1;if(exec){const today=new Date();today.setHours(0,0,0,0);const dl=new Date(today);dl.setDate(today.getDate()+SD.planDays-1);mt.push({id:"НОВАЯ",name:"Новая заявка (гипотеза)",exec:exec,dl:dl.toISOString().substring(0,10),tz:taskDays,src:"new",proj:""});desc="Добавить "+exec+" новую заявку: "+taskDays+" дн. (дедлайн конец окна)";}else{desc="Укажите исполнителя: добавить заявку Иванову на N дней";}}else if(fromW){const fromE=sdFindExec(fromW)||fromW;const toE=toW?sdFindExec(toW):null;let cnt=0;for(var i=0;i<mt.length;i++){const t=mt[i];if(!t.exec)continue;const el=t.exec.toLowerCase();const fw=fromW.toLowerCase();const stem=fw.substring(0,Math.min(5,fw.length));if(!el.startsWith(stem)&&el!==fw)continue;if(projF&&!t.proj.toLowerCase().includes(projF)&&!t.name.toLowerCase().includes(projF))continue;mt[i]=Object.assign({},t,{exec:toE||""});cnt++;}desc="Снять "+fromE+" с "+cnt+" задач"+(projF?" ("+projF+")":"")+(toE?", передать "+toE:", оставить без исполнителя");}else if(shiftDays>0){let cnt=0;for(var j=0;j<mt.length;j++){const t=mt[j];if(projF&&!t.proj.toLowerCase().includes(projF)&&!t.name.toLowerCase().includes(projF)&&!t.id.toLowerCase().includes(projF))continue;const dl=new Date(t.dl);dl.setDate(dl.getDate()+shiftDays*shiftDir);mt[j]=Object.assign({},t,{dl:dl.toISOString().substring(0,10)});cnt++;}desc="Сдвиг дедлайна"+(projF?" ("+projF+")":"")+" "+(shiftDir>0?"+":"-")+shiftDays+" дн. (затронуто "+cnt+" задач)";}return{tasks:mt,desc:desc,availMod:availMod};}') | Out-Null
$sb.AppendLine('function sdFormatComp(b,a,availMod){const all=[...new Set([...Object.keys(b),...Object.keys(a)])];let s="ДО:\n";all.forEach(function(nm){const e=b[nm];s+="  "+nm+": "+(e?e.tz+" дн., К="+(e.coeff!==null?e.coeff:"н/д")+(e.cov<100?" [данные "+e.cov+"%]":"")+" "+(e.coeff&&e.coeff>1?"НЕ УСПЕВАЕТ":"ок"):"нет задач в окне")+"\n";});s+="\nПОСЛЕ:\n";all.forEach(function(nm){const e=a[nm];const modNote=availMod&&availMod[nm]!==undefined?" [доступно "+availMod[nm]+" р.д.]":"";s+="  "+nm+": "+(e?e.tz+" дн., К="+(e.coeff!==null?e.coeff:"н/д")+modNote+(e.cov<100?" [данные "+e.cov+"%]":"")+" "+(e.coeff&&e.coeff>1?"НЕ УСПЕВАЕТ":"ок"):"нет задач в окне")+"\n";});return s;}') | Out-Null
$sb.AppendLine('function sdBuildCompHtml(b,a,desc,availMod){const all=[...new Set([...Object.keys(b),...Object.keys(a)])];all.sort(function(x,y){return ((a[y]&&a[y].coeff)||0)-((a[x]&&a[x].coeff)||0);});let h="<div class=\"sd-comp\"><div class=\"sd-comp-lbl\">&#9889; "+desc+"</div>";h+="<table class=\"sd-comp-tbl\"><thead><tr><th>Исполнитель</th><th>До</th><th>После</th><th>&#916;</th></tr></thead><tbody>";all.forEach(function(nm){const bE=b[nm];const aE=a[nm];const bC=bE?bE.coeff:null;const aC=aE?aE.coeff:null;const diff=bC!==null&&aC!==null?Math.round((aC-bC)*100)/100:null;const bCls=bC!==null?(bC>1?"red":"green"):"";const aCls=aC!==null?(aC>1?"red":"green"):"";const dCls=diff!==null?(diff>0?"red":diff<0?"green":""):"";const bWarn=bE&&bE.cov<100?" <span class=\"sd-warn\">"+bE.cov+"%</span>":"";const aWarn=aE&&aE.cov<100?" <span class=\"sd-warn\">"+aE.cov+"%</span>":"";const aMod=availMod&&availMod[nm]!==undefined?" <span style=\"color:#64748b;font-size:10px\">("+availMod[nm]+"р.д.)</span>":"";const bTxt=bE?"К="+(bC!==null?bC:"н/д")+bWarn:"—";const aTxt=aE?"К="+(aC!==null?aC:"н/д")+aMod+aWarn:"—";const dTxt=diff!==null?(diff>0?"+"+diff:diff.toString()):"—";h+="<tr><td>"+nm+"</td><td class=\""+bCls+"\">"+bTxt+"</td><td class=\""+aCls+"\">"+aTxt+"</td><td class=\""+dCls+"\"><b>"+dTxt+"</b></td></tr>";});h+="</tbody></table><div style=\"font-size:10px;color:#94a3b8;margin-top:5px\">К=трудозатраты/раб.дни окна "+SD.planDays+"дн. ("+SD.avail+" раб.). К&gt;1.0 — не успевает.</div></div>";return h;}') | Out-Null
$sb.AppendLine('function sdFormInit(){const execs=SD.execs||[];const projs=SD.projs||[];["sd-from","sd-to"].forEach(function(id){const sel=document.getElementById(id);if(!sel)return;sel.innerHTML=id==="sd-to"?"<option value=\"\">— без исполнителя —</option>":"";execs.forEach(function(e){sel.innerHTML+="<option value=\""+e+"\">"+e+"</option>";});});const ps=document.getElementById("sd-proj");if(ps){ps.innerHTML="<option value=\"\">— все задачи —</option>";projs.forEach(function(p){ps.innerHTML+="<option value=\""+p+"\">"+p+"</option>";});}sdFormToggle();}') | Out-Null
$sb.AppendLine('function sdTogglePanel(){const p=document.getElementById("sd-panel");if(!p)return;if(p.style.display==="none"||!p.style.display){p.style.display="block";sdFormInit();}else p.style.display="none";}') | Out-Null
$sb.AppendLine('function sdFormToggle(){const act=document.getElementById("sd-action").value;document.getElementById("sd-to-row").style.display=act==="reassign"?"flex":"none";document.getElementById("sd-proj-row").style.display=(act==="reassign"||act==="shift")?"flex":"none";document.getElementById("sd-days-row").style.display=(act==="newtask"||act==="shift"||act==="absence")?"flex":"none";document.getElementById("sd-df-row").style.display=act==="absence"?"flex":"none";document.getElementById("sd-dt-row").style.display=act==="absence"?"flex":"none";}') | Out-Null
$sb.AppendLine('function runScenarioForm(){const act=document.getElementById("sd-action").value;const from=document.getElementById("sd-from").value;const to=document.getElementById("sd-to").value;const proj=document.getElementById("sd-proj").value;const days=document.getElementById("sd-days").value;const df=document.getElementById("sd-date-from").value;const dt=document.getElementById("sd-date-to").value;let q="";if(act==="reassign"){q="Что если снять "+from+(proj?" с проекта "+proj:"")+(to?" и передать "+to:"")+"?";}else if(act==="absence"){q="Что если "+from+" в отпуске"+(df&&dt?" с "+df+" по "+dt:"")+(days?" на "+days+" дней":"")+"?";}else if(act==="newtask"){q="Что если добавить заявку для "+from+" на "+days+" дней?";}else if(act==="shift"){q="Что если сдвинуть дедлайн"+(proj?" проекта "+proj:"")+" на "+days+" дней?";}if(q){document.getElementById("sd-panel").style.display="none";switchTab("ai");document.getElementById("ai-inp").value=q;sendAI();}}') | Out-Null
$sb.AppendLine('async function sendAI(){') | Out-Null
$sb.AppendLine('  if(aiWorking)return;') | Out-Null
$sb.AppendLine('  const inp=document.getElementById("ai-inp");') | Out-Null
$sb.AppendLine('  const msg=inp.value.trim(); if(!msg)return;') | Out-Null
$sb.AppendLine('  const key=localStorage.getItem("opu_ai_key"); if(!key){initAI();return;}') | Out-Null
$sb.AppendLine('  aiWorking=true;') | Out-Null
$sb.AppendLine('  const sendBtn=document.getElementById("ai-send-btn");sendBtn.disabled=true;') | Out-Null
$sb.AppendLine('  inp.value="";') | Out-Null
$sb.AppendLine('  addAIMsg("user",renderAIMd(msg));') | Out-Null
$sb.AppendLine('  if(aiHistory.length>10)aiHistory=aiHistory.slice(-10);') | Out-Null
$sb.AppendLine('  aiHistory.push({role:"user",content:msg});') | Out-Null
$sb.AppendLine('  let sdExtra="";let sdHtmlTable="";if(sdDetect(msg)){try{const sdR=sdApply(SD.tasks,msg);const sdB=sdCalcLoad(SD.tasks,{});const sdA=sdCalcLoad(sdR.tasks,sdR.availMod);sdHtmlTable=sdBuildCompHtml(sdB,sdA,sdR.desc,sdR.availMod);sdExtra="\n\nСЦЕНАРИЙ «ЧТО ЕСЛИ» (детерминированный расчёт JS, не придумывай числа!):\nГипотеза применена: "+sdR.desc+"\nОкно: "+SD.planDays+" дн. ("+SD.avail+" раб. дней). К>1.0 = объём задач превышает доступное время.\n\n"+sdFormatComp(sdB,sdA,sdR.availMod)+"\nОбъясни: кто оказался в риске, кто освободился? Какие задачи под угрозой? Конкретные рекомендации руководителю.";}catch(err){sdExtra="\n\n[Ошибка расчёта: "+err.message+"]";}}if(sdHtmlTable){const tblDiv=document.createElement("div");tblDiv.className="ai-msg ai-assistant";tblDiv.innerHTML=sdHtmlTable;document.getElementById("ai-msgs").appendChild(tblDiv);document.getElementById("ai-msgs").scrollTop=9999;}') | Out-Null
$sb.AppendLine('  const typingDiv=addAIMsg("assistant","<span class=''ai-typing''><span></span><span></span><span></span></span>");') | Out-Null
$sb.AppendLine('  const timerEl=document.createElement("span");timerEl.className="ai-timer";timerEl.textContent="0с";typingDiv.appendChild(timerEl);') | Out-Null
$sb.AppendLine('  let elapsed=0;const timerInterval=setInterval(function(){elapsed++;timerEl.textContent=elapsed+"с";},1000);') | Out-Null
$sb.AppendLine('  const days=["воскресенье","понедельник","вторник","среда","четверг","пятница","суббота"];') | Out-Null
$sb.AppendLine('  const modeInstr=aiVerbose?" Отвечай развёрнуто, с деталями. Структурируй по разделам.":" Отвечай кратко — до 7 предложений или маркированный список.";') | Out-Null
$sb.AppendLine('') | Out-Null
$sb.AppendLine('  const sys="Ты ИИ-ассистент по управлению проектным портфелем ОПУ. Отвечай на русском языке."+modeInstr+" Указывай ID заявок в скобках [ID]. Исполнителей называй по фамилии. Предлагай конкретные следующие шаги. Сегодня: "+new Date().toLocaleDateString("ru-RU")+" ("+days[new Date().getDay()]+").\n\nДАННЫЕ ПОРТФЕЛЯ:\n"+AI_CTX+sdExtra;') | Out-Null
$sb.AppendLine('  try {') | Out-Null
$sb.AppendLine('    const resp=await fetch("https://api.groq.com/openai/v1/chat/completions",{method:"POST",headers:{"Content-Type":"application/json","Authorization":"Bearer "+key},body:JSON.stringify({model:"llama-3.1-8b-instant",max_tokens:800,messages:[{role:"system",content:sys}].concat(aiHistory)})});') | Out-Null
$sb.AppendLine('    const data=await resp.json();') | Out-Null
$sb.AppendLine('    if(!resp.ok)throw new Error(data.error?.message||"HTTP "+resp.status);') | Out-Null
$sb.AppendLine('    const answer=data.choices[0].message.content;') | Out-Null
$sb.AppendLine('    aiHistory.push({role:"assistant",content:answer});saveAIHistory();') | Out-Null
$sb.AppendLine('    typingDiv.remove();clearInterval(timerInterval);addAIMsg("assistant",renderAIMd(answer));') | Out-Null
$sb.AppendLine('  } catch(e){typingDiv.remove();clearInterval(timerInterval);addAIMsg("assistant","Ошибка: "+e.message,true);}') | Out-Null
$sb.AppendLine('  aiWorking=false;sendBtn.disabled=false;inp.focus();') | Out-Null
$sb.AppendLine('}') | Out-Null
$sb.AppendLine("const GEN_TS=new Date('$genTime').getTime();") | Out-Null
$sb.AppendLine('function checkStale(){if(Date.now()-GEN_TS>12*3600*1000){document.getElementById("stale-banner").classList.add("on");}}') | Out-Null
$sb.AppendLine('function toggleUrgent(){const btn=document.getElementById("btn-urgent");const on=btn.classList.toggle("active");document.querySelectorAll(".proj").forEach(function(proj){if(!on){proj.style.display="";proj.querySelectorAll(".trow").forEach(function(r){r.style.display="";});return;}var hasUrgent=false;proj.querySelectorAll(".trow").forEach(function(r){var u=r.querySelector(".dl-red,.dl-orange");r.style.display=u?"":"none";if(u)hasUrgent=true;});proj.style.display=hasUrgent?"":"none";});}') | Out-Null
$sb.AppendLine('window.addEventListener("DOMContentLoaded",function(){') | Out-Null
$sb.AppendLine('  const p=new URLSearchParams(window.location.search);') | Out-Null
$sb.AppendLine('  const q=p.get("q"),ex=p.get("exec"),dp=p.get("dept");') | Out-Null
$sb.AppendLine('  if(q)document.getElementById("srch").value=q;') | Out-Null
$sb.AppendLine('  if(ex)document.getElementById("f-ex").value=ex;') | Out-Null
$sb.AppendLine('  if(dp)document.getElementById("f-dp").value=dp;') | Out-Null
$sb.AppendLine('  if(q||ex||dp)filterTasks();') | Out-Null
$sb.AppendLine('  initAR();') | Out-Null
$sb.AppendLine('  initPanels();') | Out-Null
$sb.AppendLine('  initRemChecks();') | Out-Null
$sb.AppendLine('  checkMissedDone();') | Out-Null
$sb.AppendLine('  initWatchPanel();') | Out-Null
$sb.AppendLine('  initAI();') | Out-Null
$sb.AppendLine('  checkStale();') | Out-Null
$sb.AppendLine('  document.getElementById("gantt-modal").addEventListener("click",function(e){if(e.target===this)closeGantt();});') | Out-Null
$sb.AppendLine('  document.addEventListener("click",function(e){const dock=document.getElementById("restore-dock");if(dock&&!dock.contains(e.target)){const l=document.getElementById("rd-list");if(l)l.style.display="none";}});') | Out-Null
$sb.AppendLine('});') | Out-Null
$sb.AppendLine('</script>') | Out-Null
# ── Вкладка «Сверка с Битриксом»: данные + разметка + рендер ──
$sb.AppendLine('<script>') | Out-Null
$sb.Append('window.SVERKA=') | Out-Null
$sb.Append($sverkaJson) | Out-Null
$sb.AppendLine(';') | Out-Null
$sb.Append('window.SVERKA_MSG=') | Out-Null
$sb.Append(($sverkaMsg | ConvertTo-Json)) | Out-Null
$sb.AppendLine(';</script>') | Out-Null
$sb.AppendLine(@'
<div id="tab-sver" style="display:none">
<div class="panel">
<h2>Сверка с Битриксом</h2>
<p class="sver-intro">Сравнение «Реестра заявок» с задачами Битрикс24 (привязка по ссылке на задачу в столбце ID). Подсвечены расхождения дедлайна и исполнителя. Данные - снимок на момент сборки дашборда, не в реальном времени.</p>
<div id="sver-empty" class="sver-empty" style="display:none"></div>
<div id="sver-wrap">
  <div class="sver-cards">
    <div class="sver-card"><div class="sver-num" id="sver-n-total">0</div><div class="sver-lbl">задач сверено</div></div>
    <div class="sver-card"><div class="sver-num sver-red" id="sver-n-diff">0</div><div class="sver-lbl">с расхождениями</div></div>
    <div class="sver-card"><div class="sver-num" id="sver-n-dl">0</div><div class="sver-lbl">по дедлайну</div></div>
    <div class="sver-card"><div class="sver-num" id="sver-n-ex">0</div><div class="sver-lbl">по исполнителю</div></div>
  </div>
  <label class="sver-only"><input type="checkbox" id="sver-only-diff" onchange="sverRender()"> показывать только расхождения</label>
  <div class="sver-tbl-wrap">
  <table class="sver-tbl">
    <thead><tr><th>ID</th><th>Задача</th><th>Дедлайн<br><span>реестр / Битрикс</span></th><th>Исполнитель<br><span>реестр / Битрикс</span></th></tr></thead>
    <tbody id="sver-body"></tbody>
  </table>
  </div>
</div>
<style>
#tab-sver .sver-intro{color:#64748b;font-size:13px;margin:-4px 0 14px;max-width:900px;line-height:1.5}
#tab-sver .sver-empty{background:#fffbeb;border:1px solid #fde68a;color:#92400e;border-radius:10px;padding:16px;font-size:13px}
#tab-sver .sver-cards{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:14px}
#tab-sver .sver-card{flex:1;min-width:130px;background:#fff;border:1px solid #eef2f7;border-radius:12px;padding:14px 16px;box-shadow:0 1px 3px rgba(0,0,0,.05)}
#tab-sver .sver-num{font-size:30px;font-weight:700;line-height:1;color:#1e3a5f}
#tab-sver .sver-num.sver-red{color:#dc2626}
#tab-sver .sver-lbl{font-size:12px;color:#64748b;margin-top:6px}
#tab-sver .sver-only{display:inline-flex;align-items:center;gap:7px;font-size:13px;color:#475569;margin-bottom:12px;cursor:pointer;user-select:none}
#tab-sver .sver-tbl-wrap{overflow-x:auto}
#tab-sver .sver-tbl{width:100%;border-collapse:collapse;font-size:13px}
#tab-sver .sver-tbl th{text-align:left;padding:9px 12px;border-bottom:2px solid #e2e8f0;color:#475569;font-weight:600;white-space:nowrap;vertical-align:top}
#tab-sver .sver-tbl th span{font-weight:400;font-size:11px;color:#94a3b8}
#tab-sver .sver-tbl td{padding:9px 12px;border-bottom:1px solid #f1f5f9;vertical-align:top}
#tab-sver .sver-tbl tbody tr:hover,#tab-sver .sver-tbl tbody tr:hover{background:#f8fafc}
#tab-sver .sver-tbl a{color:#1e3a5f;font-weight:700;text-decoration:none}
#tab-sver .sver-tbl a:hover{text-decoration:underline}
#tab-sver .sver-diff{background:#fef2f2;color:#b91c1c;font-weight:600;border-radius:5px}
#tab-sver .sver-arr{color:#cbd5e1}
#tab-sver .sver-miss{color:#b45309;font-style:italic}
</style>
<script>
function sverEsc(s){return String(s==null?"":s).replace(/[&<>"]/g,function(c){return{"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c];});}
function sverRender(){
  var data=window.SVERKA||[];
  var wrap=document.getElementById("sver-wrap"),empty=document.getElementById("sver-empty");
  if(!data.length){wrap.style.display="none";empty.style.display="block";empty.textContent=window.SVERKA_MSG||"Нет заявок со ссылкой на задачу Битрикса - сверять нечего.";return;}
  wrap.style.display="block";empty.style.display="none";
  var onlyDiff=document.getElementById("sver-only-diff").checked;
  var nDiff=0,nDl=0,nEx=0;
  data.forEach(function(d){if(d.dlDiff||d.execDiff||!d.found)nDiff++;if(d.dlDiff)nDl++;if(d.execDiff)nEx++;});
  document.getElementById("sver-n-total").textContent=data.length;
  document.getElementById("sver-n-diff").textContent=nDiff;
  document.getElementById("sver-n-dl").textContent=nDl;
  document.getElementById("sver-n-ex").textContent=nEx;
  var rows="";
  data.forEach(function(d){
    var any=d.dlDiff||d.execDiff||!d.found;
    if(onlyDiff&&!any)return;
    var idTxt=sverEsc(d.reg||d.bxId);
    var idCell=d.url?("<a href=\""+sverEsc(d.url)+"\" target=\"_blank\">"+idTxt+"</a>"):idTxt;
    if(!d.found){rows+="<tr><td>"+idCell+"</td><td>"+sverEsc(d.name)+"</td><td colspan=\"2\" class=\"sver-miss\">задача "+sverEsc(d.bxId)+" не найдена в Битриксе</td></tr>";return;}
    rows+="<tr><td>"+idCell+"</td><td>"+sverEsc(d.name)+"</td>"
      +"<td class=\""+(d.dlDiff?"sver-diff":"")+"\">"+sverEsc(d.regDl||"—")+" <span class=\"sver-arr\">/</span> "+sverEsc(d.bxDl||"—")+"</td>"
      +"<td class=\""+(d.execDiff?"sver-diff":"")+"\">"+sverEsc(d.regExec||"—")+" <span class=\"sver-arr\">/</span> "+sverEsc(d.bxExec||"—")+"</td></tr>";
  });
  document.getElementById("sver-body").innerHTML=rows||"<tr><td colspan=\"4\" style=\"text-align:center;color:#94a3b8;padding:18px\">Нет строк для показа</td></tr>";
}
</script>
</div>
</div>
'@) | Out-Null
$sb.AppendLine('</body></html>') | Out-Null

# ── AES-CBC шифрование дашборда ──────────────────────────────────────────────
function Encrypt-Dashboard {
    param([string]$plaintext, [string]$password)
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.KeySize = 256; $aes.BlockSize = 128
    $aes.Mode    = [System.Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
    $rng  = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $salt = New-Object byte[] 16; $rng.GetBytes($salt)
    $iv   = New-Object byte[] 16; $rng.GetBytes($iv)
    $pbk  = New-Object System.Security.Cryptography.Rfc2898DeriveBytes(
                $password, $salt, 100000,
                [System.Security.Cryptography.HashAlgorithmName]::SHA256)
    $key  = $pbk.GetBytes(32)
    $aes.Key = $key; $aes.IV = $iv
    $enc    = $aes.CreateEncryptor()
    $plain  = [System.Text.Encoding]::UTF8.GetBytes($plaintext)
    $cipher = $enc.TransformFinalBlock($plain, 0, $plain.Length)
    [System.Convert]::ToBase64String($salt + $iv + $cipher)
}

$_encPwd  = $dashPwd
$_encData = Encrypt-Dashboard $sb.ToString() $_encPwd

$_loader = @"
<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Дашборд ОПУ</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0f172a;display:flex;align-items:center;justify-content:center;min-height:100vh;font-family:'Segoe UI',system-ui,sans-serif}
.lk{background:#1e293b;border-radius:16px;padding:40px 36px;width:360px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.5)}
.lk h1{color:#f1f5f9;font-size:20px;margin-bottom:8px}
.lk p{color:#94a3b8;font-size:13px;margin-bottom:24px}
.lk input{width:100%;padding:12px 16px;border-radius:8px;border:1.5px solid #334155;background:#0f172a;color:#f1f5f9;font-size:15px;margin-bottom:12px;outline:none}
.lk input:focus{border-color:#3b82f6}
.lk button{width:100%;padding:12px;border-radius:8px;background:#1e3a5f;color:#fff;border:none;font-size:15px;font-weight:600;cursor:pointer}
.lk button:hover{background:#1d4ed8}
.err{color:#f87171;font-size:13px;margin-top:10px;min-height:18px}
.spin{display:none;margin-top:14px;color:#64748b;font-size:13px}
</style>
</head>
<body>
<div class="lk">
<h1>Дашборд ОПУ</h1>
<p>Введите пароль для доступа</p>
<input type="password" id="pwd" placeholder="Пароль" autofocus>
<button onclick="unlock()">Войти</button>
<div class="err" id="err"></div>
<div class="spin" id="spin">Расшифровка...</div>
</div>
<script>
const D='$_encData';
async function dec(b64,pwd){
  const d=Uint8Array.from(atob(b64),c=>c.charCodeAt(0));
  const salt=d.slice(0,16),iv=d.slice(16,32),ct=d.slice(32);
  const km=await crypto.subtle.importKey('raw',new TextEncoder().encode(pwd),'PBKDF2',false,['deriveKey']);
  const k=await crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:100000,hash:'SHA-256'},km,{name:'AES-CBC',length:256},false,['decrypt']);
  const pt=await crypto.subtle.decrypt({name:'AES-CBC',iv},k,ct);
  return new TextDecoder().decode(pt);
}
async function unlock(){
  const pwd=document.getElementById('pwd').value;
  if(!pwd){document.getElementById('err').textContent='Введите пароль';return;}
  document.getElementById('err').textContent='';
  document.getElementById('spin').style.display='block';
  try{
    const html=await dec(D,pwd);
    sessionStorage.setItem('opu_pwd',pwd);
    localStorage.setItem('opu_auth','1');
    document.open();document.write(html);document.close();
  }catch(e){
    document.getElementById('spin').style.display='none';
    document.getElementById('err').textContent='Неверный пароль';
  }
}
document.getElementById('pwd').addEventListener('keypress',function(e){if(e.key==='Enter')unlock();});
(async function(){
  const p=sessionStorage.getItem('opu_pwd');
  if(p){try{const h=await dec(D,p);localStorage.setItem('opu_auth','1');document.open();document.write(h);document.close();}catch(e){sessionStorage.removeItem('opu_pwd');}}
})();
</script>
</body>
</html>
"@

[System.IO.File]::WriteAllText($outputFile, $_loader, $utf8bom)
Write-Host "Готово: $outputFile (ссылок привязано: $($linkMap.Count))"
