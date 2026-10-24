# Loaded inside Workflow.psm1 to use its authorization and snapshot helpers.
function Get-WStatisticsSummary {
    param([object[]]$Cases=@(),[object[]]$Accounts=@())
    $participants=@($Accounts | Where-Object {$_.participating})
    $complete=@($Cases | Where-Object {$_.evaluationStatus -eq '평가 완료'}).Count
    $delivered=@($Cases | Where-Object {$_.deliveryStatus -eq '입력 완료'}).Count
    [pscustomobject]@{
        accountCount=$Accounts.Count;participantCount=$participants.Count
        unselectedCount=@($participants | Where-Object {$_.selectionStatus -in @('미선정','임시저장')}).Count
        caseCount=$Cases.Count;deliveredCount=$delivered
        waitingCount=@($Cases | Where-Object {$_.deliveryStatus -eq '대기 중'}).Count
        completeCount=$complete;partialCount=@($Cases | Where-Object {$_.evaluationStatus -eq '평가 중'}).Count
        unreviewedCount=@($Cases | Where-Object {$_.evaluationStatus -eq '미평가'}).Count
        staleCount=@($Cases | Where-Object {$_.evaluationStatus -eq '재확인 필요'}).Count
        deliveryRate=$(if ($Cases.Count -and @($Cases | Where-Object {$_.aiTotal -gt 0}).Count) {'{0:0.#}%' -f (100.0*$delivered/$Cases.Count)} else {'—'})
        evaluationRate=$(if ($Cases.Count) {'{0:0.#}%' -f (100.0*$complete/$Cases.Count)} else {'—'})
    }
}

function Get-AdminStatistics {
    param([Parameter(Mandatory)]$Context,[Parameter(Mandatory)][string]$Period)
    $null=Assert-WAdmin $Context
    # A shared month lock fixes its participants/sources while each case is read
    # under its own short lock. Different cases can still be saved concurrently.
    Invoke-WPeriodReadLock $Context $Period {
        $settings=Get-Period $Context $Period
        $users=@(Get-Users $Context);$groups=@(Get-Groups $Context);$sources=@(Get-Sources $Context)
        $userMap=@{};foreach ($user in $users) {$userMap[$user.userId]=$user}
        $groupMap=@{};foreach ($group in $groups) {$groupMap[$group.groupId]=$group}
        $sourceMap=@{};foreach ($source in $sources) {$sourceMap[$source.sourceId]=$source}
        $caseRows=New-Object 'System.Collections.Generic.List[object]'
        $sourceRows=New-Object 'System.Collections.Generic.List[object]'
        $cancelledCount=0
        foreach ($case in @(Get-Cases $Context $Period All)) {
            $captured=Invoke-FileLock -Path (Get-WCaseLock $Context $case) -Action {
                $actual=Get-WCase $Context $case
                if ($actual.status -eq 'Cancelled') {return [pscustomobject]@{cancelled=$true}}
                $evaluation=Read-Revision -Path (Get-WEvaluationPath $Context $actual)
                $published=@{};$latest=@{};$stale=$false
                foreach ($sourceId in $settings.sourceIds) {
                    $published[$sourceId]=Get-WPublished $Context $actual $sourceId
                    $latest[$sourceId]=Read-Revision -Path (Get-WDeliveryPath $Context $actual $sourceId)
                    $version=if ($published[$sourceId]) {[int]$published[$sourceId].version} else {0}
                    if ($evaluation -and $version -ne [int](Get-WValue $evaluation.data.deliveryVersions $sourceId 0)) {$stale=$true}
                }
                [pscustomobject]@{cancelled=$false;case=$actual;evaluation=$evaluation;published=$published;latest=$latest;stale=$stale}
            }
            if ($captured.cancelled) {$cancelledCount++;continue}
            $actual=$captured.case;$evaluation=$captured.evaluation
            $aiTotal=0;$aiCompleted=0;$documents=0
            foreach ($sourceId in $settings.sourceIds) {
                $source=$sourceMap[$sourceId];$delivery=$captured.published[$sourceId];$edit=$captured.latest[$sourceId]
                $isAi=(Get-WValue $source 'type') -eq 'AI'
                if ($isAi) {$aiTotal++;if ($delivery) {$aiCompleted++}}
                $docs=@(if ($delivery) {Get-WValue $delivery.data 'documents' @()})
                $documents+=$docs.Count
                $matched=$evaluation -and $delivery -and [int](Get-WValue $evaluation.data.deliveryVersions $sourceId 0) -eq [int]$delivery.version
                $scoreSum=0;$scoreCount=0;$citedCount=0;$citationCount=0
                if ($matched) {
                    foreach ($rating in @(Get-WValue $evaluation.data 'documentRatings' @())) {
                        if ($rating.sourceId -ne $sourceId -or $rating.documentId -notin @($docs | ForEach-Object {$_.documentId})) {continue}
                        if ($null -ne $rating.score) {$scoreSum+=[int]$rating.score;$scoreCount++}
                        if ($null -ne $rating.cited) {$citationCount++;if ($rating.cited) {$citedCount++}}
                    }
                }
                $rank=$null
                if ($evaluation -and -not $captured.stale) {
                    $ranking=@($evaluation.data.rankings | Where-Object {$_.sourceId -eq $sourceId -and $_.status -eq 'Ranked'}) | Select-Object -First 1
                    if ($ranking) {$rank=$ranking.rank}
                }
                $assignee=@($settings.assignments | Where-Object {$_.groupId -eq $actual.groupId -and $_.sourceId -eq $sourceId}) | Select-Object -First 1
                $sourceRows.Add([pscustomobject]@{
                    caseId=$actual.caseId;userId=$actual.userId;groupId=$actual.groupId;sourceId=$sourceId
                    sourceName=(Get-WValue $source 'name' $sourceId);isAi=$isAi;assigneeId=(Get-WValue $assignee 'userId')
                    published=[bool]$delivery;draft=($edit -and $edit.data.status -eq 'Draft');documentCount=$docs.Count
                    scoreSum=$scoreSum;scoreCount=$scoreCount;citedCount=$citedCount;citationCount=$citationCount;rank=$rank
                })
            }
            $evaluationStatus=if (-not $evaluation) {'미평가'} elseif ($captured.stale) {'재확인 필요'} elseif ($evaluation.data.complete) {'평가 완료'} else {'평가 중'}
            $caseRows.Add([pscustomobject]@{
                caseId=$actual.caseId;userId=$actual.userId;groupId=$actual.groupId
                groupName=(Get-WValue $groupMap[$actual.groupId] 'name' $actual.groupId)
                examiner=$actual.examinerName;division=(Get-WValue $actual 'examinerDivision');applicationNumber=$actual.applicationNumber;cpc=(Get-WValue $actual 'cpc')
                aiTotal=$aiTotal;aiCompleted=$aiCompleted;aiProgress="$aiCompleted / $aiTotal";documentCount=$documents
                deliveryStatus=$(if (-not $aiTotal) {'AI 대상 없음'} elseif ($aiCompleted -eq $aiTotal) {'입력 완료'} else {'대기 중'})
                evaluationStatus=$evaluationStatus
            })
        }
        $accountRows=New-Object 'System.Collections.Generic.List[object]'
        foreach ($user in $users) {
            $own=@($caseRows | Where-Object {$_.userId -eq $user.userId})
            $member=@($settings.members | Where-Object {$_.userId -eq $user.userId}) | Select-Object -First 1
            $assignments=@($settings.assignments | Where-Object {$_.userId -eq $user.userId})
            $groupIds=@(@($own | ForEach-Object {$_.groupId})+@($assignments | ForEach-Object {$_.groupId})+@($(if ($member) {$member.groupId})) | Where-Object {$_} | Sort-Object -Unique)
            $selection=if ($member) {Read-Revision -Path (Get-WSelectionPath $Context $Period $user.userId)} else {$null}
            $selectionStatus=if ($own.Count) {'확정'} elseif ($selection -and $selection.data.status -eq 'Confirmed') {'확정'} elseif ($selection) {'임시저장'} elseif ($member) {'미선정'} else {'미참여'}
            $accountRows.Add([pscustomobject]@{
                userId=$user.userId;name=$user.name;bureau=(Get-WValue $user 'bureau');division=(Get-WValue $user 'division')
                active=[bool]$user.active;activeLabel=$(if ($user.active) {'활성'} else {'비활성'})
                roles=(@($user.roles | ForEach-Object {switch ($_) {'admin' {'관리자'};'examiner' {'심사관'};'ai' {'AI 담당자'}}}) -join ', ')
                groupIds=$groupIds;groupName=(@($groupIds | ForEach-Object {Get-WValue $groupMap[$_] 'name' $_}) -join ', ')
                memberGroupId=(Get-WValue $member 'groupId');participating=[bool]$member;selectionStatus=$selectionStatus
                caseCount=$own.Count;completeCount=@($own | Where-Object {$_.evaluationStatus -eq '평가 완료'}).Count
                pendingCount=@($own | Where-Object {$_.evaluationStatus -ne '평가 완료'}).Count
                assignments=(@($assignments | ForEach-Object {(Get-WValue $groupMap[$_.groupId] 'name' $_.groupId)+' / '+(Get-WValue $sourceMap[$_.sourceId] 'name' $_.sourceId)}) -join ', ')
            })
        }
        [pscustomobject]@{period=$Period;periodStatus=$settings.status;capturedAt=[datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss');cancelledCount=$cancelledCount
            cases=@($caseRows.ToArray());accounts=@($accountRows.ToArray());sourceCases=@($sourceRows.ToArray())
            groups=@($groups | Select-Object groupId,name);sources=@($sources | Where-Object {$_.sourceId -in $settings.sourceIds} | Select-Object sourceId,name)
        }
    }
}

function Select-AdminStatistics {
    param([Parameter(Mandatory)]$Snapshot,[string]$GroupId='',[string]$UserId='')
    # Only project the already authorized snapshot; filtering does not reread files.
    $cases=@($Snapshot.cases | Where-Object {(-not $GroupId -or $_.groupId -eq $GroupId) -and (-not $UserId -or $_.userId -eq $UserId)})
    $accounts=@($Snapshot.accounts | Where-Object {(-not $GroupId -or $GroupId -in $_.groupIds) -and (-not $UserId -or $_.userId -eq $UserId)})
    $sourceCases=@($Snapshot.sourceCases | Where-Object {(-not $GroupId -or $_.groupId -eq $GroupId) -and (-not $UserId -or $_.userId -eq $UserId)})
    $accountRows=@(foreach ($account in $accounts) {
        $own=@($cases | Where-Object {$_.userId -eq $account.userId})
        $tasks=@($Snapshot.sourceCases | Where-Object {$_.isAi -and $_.assigneeId -eq $account.userId -and (-not $GroupId -or $_.groupId -eq $GroupId)})
        $summary=Get-WStatisticsSummary -Cases $own
        $row=$account | Select-Object *
        if ($GroupId -and $account.memberGroupId -ne $GroupId) {Set-WValue $row 'participating' $false;if (-not $own.Count) {Set-WValue $row 'selectionStatus' '미참여'}}
        foreach ($entry in @{caseCount=$own.Count;completeCount=$summary.completeCount;pendingCount=($own.Count-$summary.completeCount);waitingCount=$summary.waitingCount;staleCount=$summary.staleCount
            aiAssigned=$tasks.Count;aiPublished=@($tasks | Where-Object {$_.published}).Count;aiPending=@($tasks | Where-Object {-not $_.published}).Count}.GetEnumerator()) {Set-WValue $row $entry.Key $entry.Value}
        Set-WValue $row 'aiWork' $(if ($tasks.Count) {"$($row.aiPublished) / $($row.aiAssigned)"} else {'—'})
        $row
    })
    $groupRows=@(foreach ($group in @($Snapshot.groups | Where-Object {-not $GroupId -or $_.groupId -eq $GroupId})) {
        $members=@($accounts | Where-Object {$_.memberGroupId -eq $group.groupId})
        $groupCases=@($cases | Where-Object {$_.groupId -eq $group.groupId})
        $summary=Get-WStatisticsSummary -Cases $groupCases -Accounts $members
        Set-WValue $summary 'groupId' $group.groupId;Set-WValue $summary 'groupName' $group.name
        $summary
    })
    $sourceRows=@(foreach ($source in $Snapshot.sources) {
        $rows=@($sourceCases | Where-Object {$_.sourceId -eq $source.sourceId})
        $scoreCount=0;$scoreSum=0;$citationCount=0;$citedCount=0;$documentCount=0
        foreach ($row in $rows) {$scoreCount+=$row.scoreCount;$scoreSum+=$row.scoreSum;$citationCount+=$row.citationCount;$citedCount+=$row.citedCount;$documentCount+=$row.documentCount}
        [pscustomobject]@{sourceName=$source.name;sourceId=$source.sourceId;expectedCount=$rows.Count;publishedCount=@($rows | Where-Object {$_.published}).Count
            pendingCount=@($rows | Where-Object {-not $_.published}).Count;documentCount=$documentCount
            scoreCount=[int]$scoreCount;averageScore=$(if ($scoreCount) {'{0:0.00}' -f ($scoreSum/$scoreCount)} else {'—'})
            citedCount=[int]$citedCount;citationCount=[int]$citationCount;citationRate=$(if ($citationCount) {'{0:0.#}%' -f (100.0*$citedCount/$citationCount)} else {'—'})
            firstCount=@($rows | Where-Object {$_.rank -eq 1}).Count
        }
    })
    [pscustomobject]@{summary=(Get-WStatisticsSummary -Cases $cases -Accounts $accountRows);groups=$groupRows;accounts=$accountRows;cases=$cases;sources=$sourceRows}
}
