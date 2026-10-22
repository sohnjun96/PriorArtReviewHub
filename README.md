# AI 선행기술조사 · 평가 관리 0.1.2

Windows PowerShell 5.1 또는 Windows용 PowerShell 7의 WPF 프로그램입니다. 공유폴더의 JSON이 원본이며 Excel 설치 없이 `.xlsx`와 첨부 ZIP을 만듭니다. AI 검색을 직접 실행하는 도구가 아니라, 제출·조사 결과 납품·평가를 관리하는 도구입니다.

핵심 업무를 실행할 수 있는 시제품입니다. 사용자 등록 시 심사국·심사과를, 출원 등록 시 출원번호와 CPC를 입력할 수 있습니다. 관리자 문헌 일괄 등록, 진행 중인 건의 그룹 이관, 조사 방식 간 동일 문헌·인용 불일치 자동 안내는 아직 포함하지 않습니다. 사용자 정보는 입력 화면으로 편집하며, 그룹·조사 방식·월 운영 설정은 JSON 입력 방식입니다.

최신 배포본: [AIReview-0.1.2.zip](dist/AIReview-0.1.2.zip). 0.1.0에서 실행 정책 오류가 났다면 이 버전의 `실행.cmd`를 사용하세요.

## 로컬 시연

`실행.cmd`를 실행하거나 다음 명령을 사용합니다. 현재 Windows 계정에 연결된 시연 관리자와 가상 참여자가 `DemoData`에 생성됩니다. 실제 출원정보는 시연 저장소에 넣지 마세요.

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\App\Start.ps1 -Demo
```

## 첫 운영 설정과 배포

전산 담당자가 승인한 **비어 있는 저장소**를 지정합니다. 다음은 설명용 경로이며 스크립트가 서버나 ACL을 자동 설정하지 않습니다. 초기화는 현재 계정 SID에 연결된 `U001` 한 명, `G01`, 조사 방식 세 개, 현재 운영월만 만듭니다. 별도 관리자의 SID를 쓰려면 `-AdminSid 'S-1-5-21-…'`를 추가합니다. 기존 파일이 있으면 중단하며, 시연 표식은 만들지 않습니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\Tools\Initialize-Store.ps1 -Root '\\FILESERVER\AIReview\Store' -AdminName '관리자 성명'
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\Tools\Publish-Release.ps1 -Destination '\\FILESERVER\AIReview\Program' -DataRoot '\\FILESERVER\AIReview\Store'
```

사용자는 배포 폴더의 `실행.cmd`를 실행합니다. `Launch.ps1`이 정확히 `App\releases\0.1.2`를 가리킵니다. 직접 실행할 때는 다음과 같습니다.

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\App\Start.ps1 -DataRoot '\\FILESERVER\AIReview\Store'
```

실행 파일은 해당 PowerShell 프로세스에만 `RemoteSigned`를 적용합니다. 컴퓨터·사용자 계정의 영구 실행 정책을 변경하지 않으며, `Bypass`는 사용하지 않습니다. 기관의 `MachinePolicy`·`UserPolicy`와 스크립트 서명 기준은 계속 우선합니다. 배포 스크립트는 기존 버전과 실행 진입점을 덮어쓰지 않으므로 업데이트 때는 새 배포 폴더에서 검증한 뒤 사용자의 바로가기를 전환합니다. 초기화가 중간 실패하면 일부 기록이 남을 수 있습니다. 아직 사용하지 않은 저장소인지 확인하고 별도 보존한 후 새 빈 폴더에 초기화하세요.

## 스크립트를 실행할 수 없다는 오류

0.1.1부터 `실행.cmd`가 이 실행에만 `RemoteSigned`를 지정하므로, 기본 `Restricted` 정책인 PC에서도 로컬 스크립트를 실행할 수 있습니다. 기존 파일을 쓰는 경우 프로젝트 폴더에서 위의 직접 실행 명령을 사용하세요. [Microsoft 실행 정책 설명](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1)

GitHub ZIP을 내려받았다면 인터넷에서 받은 파일이라는 표시 때문에 여전히 차단될 수 있습니다. ZIP 파일의 **속성 → 차단 해제 → 적용** 후 새 폴더에 압축을 푸세요. 이미 압축을 풀었다면, 내용을 신뢰하는 이 프로젝트 폴더에서 아래 명령으로 프로그램 파일만 차단 해제할 수 있습니다. 원본 데이터나 다른 폴더는 변경하지 않습니다.

```powershell
Get-ChildItem -LiteralPath .\App,.\Tools,.\Tests -Recurse -File |
    Where-Object { $_.Extension -in '.ps1', '.psm1' } |
    Unblock-File
```

배포 폴더의 `Launch.ps1`에도 다운로드 표시가 있다면 해당 파일에 `Unblock-File -LiteralPath .\Launch.ps1`을 적용합니다. 이 명령은 파일의 다운로드 표시만 제거하며 실행 정책을 변경하지 않습니다. [Microsoft Unblock-File 설명](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/unblock-file?view=powershell-5.1)

계속 차단되면 다음 명령으로 **Windows PowerShell 5.1** 정책을 확인하세요. `MachinePolicy` 또는 `UserPolicy`가 `Restricted`·`AllSigned`이면 전산 담당자가 허용·서명한 배포본이 필요합니다. 일부 공유폴더 경로가 인터넷 영역으로 분류될 때도 같은 담당자의 확인이 필요합니다.

```powershell
powershell.exe -NoProfile -Command "Get-ExecutionPolicy -List"
```

## 사용자·그룹·월 운영 관리

첫 관리자로 실행한 뒤 **운영 관리 → 사용자 선택 → ID 입력 → 불러오기**를 누릅니다. 신규 ID도 먼저 불러온 뒤 사용자 입력 화면에서 **성명·Windows SID·심사국·심사과·역할·활성 여부**를 입력하고 **설정 저장**을 누릅니다. 심사국과 심사과는 비워 둘 수 있습니다.

그룹·조사 방식·월 운영 설정은 **운영 관리 → 종류와 ID 입력 → 불러오기 → JSON 수정 → 설정 저장** 순서로 등록합니다. `version`과 `path`는 직접 입력하지 않습니다. 사용자부터, 그룹·조사 방식, 월 운영 순으로 등록합니다. 새 운영월은 다시 실행하면 목록에 나타납니다.

각 사용자의 PC에서 `whoami /user`로 SID를 확인합니다. 이름 선택으로 로그인하지 않으며, 하나의 SID는 활성 사용자 한 명에만 연결합니다. 역할은 소문자 `admin`, `examiner`, `ai`를 배열로 중복 부여합니다. 전출자는 삭제 대신 `active: false`로 바꾸고 파일서버 접근권한도 해제합니다.

```json
{"userId":"U023","sid":"S-1-5-21-111-222-333-1234","name":"심사관 성명","bureau":"심사국 이름","division":"심사과 이름","roles":["examiner","ai"],"active":true}
```

그룹은 `{"groupId":"G02","name":"그룹 이름","active":true}`, 조사 방식은 `{"sourceId":"S04","name":"추가 AI","type":"AI","active":true}` 형태입니다. 기본 `S01`은 `General`, `S02`·`S03`은 `AI`입니다. 일반 조사 결과는 해당 출원의 심사관 또는 관리자가 입력합니다.

월 운영 예시입니다. `members`는 제출 당시 그룹을 정하고, `assignments`는 그룹·AI별 담당자 한 명을 지정합니다. 월마다 배열을 별도로 보존합니다.

```json
{
  "period":"2026-10","status":"Open",
  "members":[{"userId":"U001","groupId":"G01"},{"userId":"U023","groupId":"G01"}],
  "sourceIds":["S01","S02","S03"],
  "assignments":[
    {"userId":"U001","groupId":"G01","sourceId":"S02"},
    {"userId":"U023","groupId":"G01","sourceId":"S03"}
  ],
  "submissionDeadline":"","deliveryDeadline":"","evaluationDeadline":"",
  "submissionLimit":0
}
```

기한 공란은 제한 없음이며, 지정할 때는 `2026-10-31T18:00:00+09:00` 같은 시각을 사용합니다. `submissionLimit: 0`은 건수 제한 없음입니다. 월 마감·재개방은 사유를 입력하고 전용 버튼을 사용합니다. 이미 진행한 월의 참여자·조사 방식 변경은 기존 건과의 연결을 확인한 뒤 진행하세요.

## 일상 업무

1. **내 제출**에서 **출원번호·CPC**를 입력하고 **목록에 추가**를 누릅니다. 목록의 값을 수정한 뒤 **출원 등록** 버튼으로 등록합니다. 한 건은 목록에 추가하지 않고 바로 등록할 수도 있습니다. 여러 건은 `출원번호<TAB>CPC` 형식으로 한 줄에 하나씩 붙여 넣을 수 있으며, 출원번호만 입력해도 됩니다. CPC는 선택 항목입니다. 같은 심사관·같은 달의 중복은 거절하며, 등록에 실패한 행은 입력 목록에 남습니다.
2. AI 담당자는 담당 그룹 접수를 선택하고 **결과 입력 → 조사 방식 선택 → 편집 자료 불러오기**로 작업합니다. 방식별 문헌 최대 7건이며, 임시저장과 납품은 별개입니다. 결과 없음은 사유를 입력합니다.
3. 심사관은 **납품 조회**에서 공개된 판을 보고 **내 평가**에서 문헌 점수·O/X·비고와 조사 방식 순위를 저장합니다. 공란과 0점·X는 구분됩니다. 납품 판이 바뀌면 평가를 다시 확인해야 합니다.
4. 왼쪽 **선택 출원 엑셀·첨부 다운로드**는 선택한 한 건을, 월별 다운로드 버튼은 선택한 운영월과 조회 범위의 전체 자료를 문서 폴더 `AI검색평가`에 저장합니다. Excel은 출원별 현황·선행기술 목록·품질 순위·문헌별 평가 네 시트이며, 각 시트에 심사국·심사과·CPC를 함께 표시합니다. ZIP에는 Excel과 첨부문헌이 포함됩니다.

출원 목록에서도 심사국·심사과·CPC를 확인할 수 있습니다. 제출 건의 `cpc`와 제출 당시 소속인 `examinerBureau`·`examinerDivision`을 함께 보관하므로, 이후 사용자 소속을 바꿔도 기존 출원의 소속은 바뀌지 않습니다. 이전 버전에서 만든 사용자·출원 데이터도 그대로 읽으며, 새 항목이 없으면 공란으로 표시합니다.

저장 충돌이 나면 입력내용을 보존한 채 최신 판을 확인하세요. 저장 직후 연결이 끊긴 경우 같은 입력으로 재시도하면 저장 요청 ID를 확인합니다. 같은 ID에 다른 내용을 보내면 거절합니다. 확정 JSON이나 `.lock` 파일을 사용자가 직접 편집·삭제하지 마세요.

## 폴더와 접근권한

| 위치 | 내용 |
|---|---|
| `App/Start.ps1`, `App/Views/Main.xaml` | 실행 화면과 이벤트 처리 |
| `App/Modules/Storage.psm1` | 배타 잠금, 버전 확인, UTF-8 JSON 확정 |
| `App/Modules/Workflow.psm1` | 사용자·월별 배정·제출·납품·평가 규칙 |
| `App/Modules/Export.psm1` | Excel·ZIP 생성 |
| `Master/{Users,Groups,Sources}/ID` | 기준정보 |
| `Periods/YYYY-MM` | 해당 월 참여자·담당자·기한·마감 상태 |
| `Data/월/그룹/Submissions/사용자/건` | 제출정보 |
| `Data/월/그룹/Deliveries/조사방식/사용자/건` | 납품자료 |
| `Data/월/그룹/Evaluations/사용자/건` | 심사관 평가 |
| `Data/월/그룹/Attachments/조사방식/사용자/건` | 첨부문헌 |
| `Locks/월` 및 각 기록의 `write.lock` | 저장·등록·월 상태 변경 조정 |

각 기록은 `revisions/000001.json`부터 전체 스냅샷으로 추가됩니다. 확정 전 임시파일은 조회하지 않습니다. 관리자는 업무별 권한 경계 폴더를 미리 만들고 공유 권한과 NTFS ACL을 함께 적용해야 합니다. GUI의 역할 설정만으로 파일 권한이 바뀌지 않습니다.

| 영역 | 필요한 권한 원칙 |
|---|---|
| 배포 프로그램 | 사용자는 읽기·실행, 배포 담당자만 쓰기 |
| 공유 루트와 상위 업무 폴더 | 필요한 경로 통과·목록 조회; 개별 하위 데이터는 별도 ACL, 필요 시 접근 기반 열거 적용 |
| 사용자·그룹·조사 방식·월 설정 | 관리자가 쓰기, 프로그램 운영에 필요한 읽기는 참여자에게 허용 |
| 제출정보 | 본인 쓰기, 본인·담당 그룹 AI 담당자·관리자 읽기 |
| 일반 조사 납품 | 해당 심사관·관리자 쓰기, 해당 심사관·관리자 읽기 |
| AI 납품·첨부 | 지정 담당자·관리자 쓰기, 지정 담당자·대상 심사관·관리자 읽기 |
| 평가 | 해당 심사관 쓰기, 해당 심사관·관리자 읽기 |
| 공유 잠금 | 해당 업무 참여자가 잠금 파일 생성·읽기·쓰기·열기 가능; 출원 잠금은 심사관과 모든 담당 AI가 공유 |
| 월 상태 잠금 `Locks/월/lifecycle.lock` | 그 월의 모든 작성자가 공유 읽기 열기, 관리자 상태 변경은 배타 열기; 최초 생성 권한도 필요 |
| 백업 | 별도 백업 계정 쓰기, 복구 담당자 읽기 |

엔티티별 `write.lock`에도 해당 작성자의 열기·쓰기 권한이 필요합니다. 잠금 파일은 존재 여부로 잠금을 판정하지 않으며, 프로세스가 가진 파일 핸들이 해제되면 다시 사용할 수 있습니다. 관리자 역할과 서버 ACL 관리 권한은 별개입니다. 작성자가 원본 폴더에 쓰기 권한을 가지는 구조이므로 직접 파일 변조, 작성자도 바꿀 수 없는 감사 이력, 강제 마감 통제를 보장하지 않습니다. 그런 요구에는 내부 API와 서버 저장소가 필요합니다.

## 점검과 백업

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\Tests\Run-All.ps1
pwsh.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\Tests\Run-All.ps1
# 실제 창을 열어 화면 동작까지 검사하려면 추가:
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\Tests\Run-All.ps1 -IncludeGui
```

로컬 테스트는 별도 프로세스의 동시 저장·오래된 버전 거절·동일 요청 재시도·중단된 프로세스의 잠금 해제·JSON 손상·WPF 구성·업무 규칙·내보내기·운영 초기화와 배포를 검사합니다. 기본 검증은 창을 표시하지 않으며 `-IncludeGui`가 실제 창 실행 검증을 추가합니다. **실제 SMB/NAS에서 두 PC 동시 사용, 실제 사용자 ACL, 네트워크 단절 중 저장·복구는 아직 검증하지 않았습니다.** 운영 전 해당 환경에서 확인하세요.

버전 이력과 별도로 데이터·첨부·월 설정·ACL을 함께 백업하고, 월 마감본을 별도 보관하세요. 파일서버 스냅샷과 외부 백업을 사용하고 출원 한 건의 실제 복구를 시험합니다. JSON 버전 이력은 백업을 대체하지 않습니다.
