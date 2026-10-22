# Prior Art Review Hub 0.1.4

Windows PowerShell 5.1 또는 Windows용 PowerShell 7의 WPF 프로그램입니다. 공유폴더의 JSON이 원본이며 Excel 설치 없이 `.xlsx`와 첨부 ZIP을 만듭니다. AI 검색을 직접 실행하는 도구가 아니라, 대상건 선정·AI 검색 결과 입력·평가를 관리하는 도구입니다.

최신 배포본: [AIReview-0.1.4.zip](dist/AIReview-0.1.4.zip).

## 로컬 시연

용도에 맞는 CMD 파일을 실행합니다. 시연 관리자와 가상 참여자가 `DemoData`에 생성됩니다.

| 실행 파일 | 로그인 방식 |
|---|---|
| `관리자용.cmd` | 사용자 선택이나 비밀번호 없이 관리자 화면으로 바로 진입 |
| `일반사용자용.cmd` | 첫 화면의 **심사과 - 이름** 목록에서 본인을 선택 |
| `실행.cmd` | `일반사용자용.cmd`와 동일 |

소스 폴더의 CMD 파일은 로컬 시연 저장소인 `DemoData`를 사용합니다. 실제 출원정보는 시연 저장소에 넣지 마세요. 직접 실행할 때는 다음과 같습니다. 관리자 모드는 명령 끝에 `-AdminMode`를 추가합니다.

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\App\Start.ps1 -Demo
```

## 첫 운영 설정과 배포

공용 데이터용 **비어 있는 저장소**를 지정합니다. 초기화는 자동 ID가 부여된 관리자 한 명, `G01`, 조사 방식 세 개와 현재 운영월을 만듭니다. Windows SID를 확인하거나 입력할 필요가 없습니다. 기존 파일이 있으면 중단하며, 시연 표식은 만들지 않습니다. 다음 경로는 예시이고 스크립트가 서버나 공유 권한을 설정하지는 않습니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\Tools\Initialize-Store.ps1 -Root '\\FILESERVER\AIReview\Store' -AdminName '관리자 성명'
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\Tools\Publish-Release.ps1 -Destination '\\FILESERVER\AIReview\Program' -DataRoot '\\FILESERVER\AIReview\Store'
```

배포 폴더에는 `관리자용.cmd`, `일반사용자용.cmd`, `실행.cmd`가 함께 생성됩니다. 관리자는 `관리자용.cmd`, 심사관·AI 담당자는 `일반사용자용.cmd`를 실행합니다. 두 실행 방식은 `Launch.ps1`을 통해 `App\releases\0.1.4`의 같은 프로그램과 위에서 지정한 **같은 공용 저장소**를 사용합니다. `실행.cmd`는 일반사용자용 실행 진입점입니다. 관리자 실행은 Windows의 **관리자 권한으로 실행**이나 UAC 승격을 요구하지 않습니다.

초기화에서 만든 관리자 계정으로 관리자 화면에 바로 진입하며, 활성 관리자 계정이 없으면 안내 메시지를 표시하고 종료합니다. 새 저장소에서는 위의 초기화 명령으로 첫 관리자를 생성합니다. 직접 실행할 때는 다음과 같습니다.

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\App\Start.ps1 -DataRoot '\\FILESERVER\AIReview\Store'
# 관리자 화면으로 바로 실행:
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\App\Start.ps1 -DataRoot '\\FILESERVER\AIReview\Store' -AdminMode
```

실행 파일은 해당 PowerShell 프로세스에만 `RemoteSigned`를 적용하며 PC의 영구 실행 정책을 변경하지 않습니다. 기관의 `MachinePolicy`·`UserPolicy`와 스크립트 서명 기준은 계속 우선합니다. 배포 스크립트는 기존 버전과 실행 진입점을 덮어쓰지 않으므로 업데이트 때는 새 배포 폴더에서 검증한 뒤 바로가기를 전환합니다. 초기화가 중간 실패하면 기록을 보존하고 새 빈 폴더에 초기화하세요.

## 스크립트를 실행할 수 없다는 오류

0.1.1부터 `실행.cmd`가 이 실행에만 `RemoteSigned`를 지정합니다. 기존 파일을 쓰는 경우 프로젝트 폴더에서 위의 직접 실행 명령을 사용하세요. [Microsoft 실행 정책 설명](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1)

GitHub ZIP을 내려받았다면 ZIP 파일의 **속성 → 차단 해제 → 적용** 후 새 폴더에 압축을 푸세요. 이미 압축을 풀었다면, 내용을 신뢰하는 이 프로젝트 폴더에서 아래 명령으로 프로그램 파일만 차단 해제할 수 있습니다.

```powershell
Get-ChildItem -LiteralPath .\App,.\Tools,.\Tests -Recurse -File |
    Where-Object { $_.Extension -in '.ps1', '.psm1' } |
    Unblock-File
```

배포 폴더의 `Launch.ps1`에도 다운로드 표시가 있다면 `Unblock-File -LiteralPath .\Launch.ps1`을 적용합니다. 이 명령은 파일의 다운로드 표시만 제거합니다. [Microsoft Unblock-File 설명](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/unblock-file?view=powershell-5.1)

계속 차단되면 다음 명령으로 정책을 확인하세요. `MachinePolicy` 또는 `UserPolicy`가 `Restricted`·`AllSigned`이면 전산 담당자가 허용·서명한 배포본이 필요합니다. 일부 공유폴더 경로가 인터넷 영역으로 분류되는 경우도 확인해야 합니다.

```powershell
powershell.exe -NoProfile -Command "Get-ExecutionPolicy -List"
```

## 사용자 등록과 로그인

관리자로 로그인한 뒤 **운영 관리 → 사용자 → 신규 사용자**에서 **성명·심사국·심사과·역할·활성 여부**를 입력하고 **설정 저장**을 누릅니다. 사용자 ID와 내부 식별값은 자동으로 부여됩니다. 기존 사용자는 사용자 목록에서 선택해 불러온 뒤 수정합니다. 심사국과 심사과는 비워 둘 수 있습니다.

관리자는 **관리자용.cmd**를 실행하면 사용자 선택이나 비밀번호 없이 관리자 화면에 바로 진입합니다. 심사관·AI 담당자는 **일반사용자용.cmd** 또는 **실행.cmd**를 실행한 뒤 첫 화면의 **심사과 - 이름** 목록에서 본인을 선택합니다. 같은 PC에서도 서로 다른 사용자를 선택할 수 있습니다. 이 선택은 업무 작성자를 정하는 간편 로그인이며 Windows 계정 인증과 연결하지 않습니다.

일반사용자용 로그인에는 활성 심사관·AI 담당자를 표시하고, 메뉴와 업무 범위는 선택한 사용자의 역할·월별 배정에 따라 정합니다. 관리자와 심사관 역할을 함께 가진 사용자를 선택해도 일반사용자용 실행에서는 관리자 메뉴를 제공하지 않습니다. 관리자 역할만 가진 사용자는 일반사용자용 로그인 목록에 표시하지 않습니다.

역할은 `admin`, `examiner`, `ai`를 중복 부여할 수 있습니다. 전출자는 삭제 대신 비활성화하여 과거 기록을 보존합니다. 이전 버전의 사용자 ID와 SID도 그대로 유지하므로 기존 사용자·납품·평가 연결을 다시 만들 필요가 없습니다.

## 그룹·조사 방식·월 운영 관리

그룹·조사 방식·월 운영 설정은 **운영 관리 → 종류와 ID 입력 → 불러오기 → JSON 수정 → 설정 저장** 순서로 등록합니다. `version`과 `path`는 직접 입력하지 않습니다. 사용자부터, 그룹·조사 방식, 월 운영 순으로 등록합니다. 새 운영월은 다시 실행하면 목록에 나타납니다.

그룹은 `{"groupId":"G02","name":"그룹 이름","active":true}`, 조사 방식은 `{"sourceId":"S04","name":"추가 AI","type":"AI","active":true}` 형태입니다. 기본 `S01`은 `General`, `S02`·`S03`은 `AI`입니다. 일반 조사 결과는 해당 출원의 심사관 또는 관리자가 입력합니다.

월 설정의 `members`에는 사용자별 자동 `userId`와 그룹을, `assignments`에는 그룹·AI별 담당자 한 명을 지정합니다. 초기 관리자는 첫 운영월의 `G01` 참여자 및 두 AI 담당자로 자동 등록됩니다. 이후 사용자를 추가했다면 해당 월에도 참여자·담당자를 추가해야 합니다.

```json
{
  "period":"2026-10","status":"Open",
  "members":[{"userId":"U자동부여된사용자ID","groupId":"G01"}],
  "sourceIds":["S01","S02","S03"],
  "assignments":[
    {"userId":"U자동부여된사용자ID","groupId":"G01","sourceId":"S02"},
    {"userId":"U자동부여된사용자ID","groupId":"G01","sourceId":"S03"}
  ],
  "submissionDeadline":"","deliveryDeadline":"","evaluationDeadline":"",
  "submissionLimit":0
}
```

예시의 `U자동부여된사용자ID`는 저장된 사용자 정보의 실제 `userId`로 바꿉니다. 월마다 참여자·배정 배열을 별도로 보존합니다. 기한 공란은 제한 없음이며, 지정할 때는 `2026-10-31T18:00:00+09:00` 같은 시각을 사용합니다. `submissionLimit: 0`은 건수 제한 없음입니다. 월 마감·재개방은 사유를 입력하고 전용 버튼을 사용합니다.

## 일상 업무

1. **대상건 선정**에서 **출원번호·CPC**를 입력하고 **목록에 추가 → 출원 등록**을 누릅니다. 한 건은 목록에 추가하지 않고 바로 등록할 수도 있습니다. 여러 건은 `출원번호<TAB>CPC` 형식으로 한 줄에 하나씩 붙여 넣습니다. CPC는 선택 항목입니다. 등록 목록에서 출원을 선택하면 같은 메뉴 안에서 조사 방식별 납품자료도 조회할 수 있습니다.
2. AI 담당자는 접수 목록에서 대상 출원을 선택한 뒤 **AI 검색 결과 입력**에서 조사 방식을 고르고 편집 자료를 불러옵니다. 결과 문헌은 **출원번호만** 입력하거나 한 줄에 하나씩 붙여 넣습니다. 문헌명·유형 입력은 필요하지 않으며 방식별 최대 7건입니다. 임시저장과 납품은 별개이고, 결과 없음은 사유와 함께 저장합니다.
3. 심사관은 **AI 검색 결과 평가**에서 **조사 방식·문헌·점수(0~5점)·인용여부(O/X)·비고**를 확인·입력합니다. 문헌 명칭 열은 표시하지 않습니다. 공란과 0점·X는 구분됩니다. 조사 방식별 품질 순위는 별도로 입력하며, 납품 판이 바뀌면 평가를 다시 확인해야 합니다.
4. **선택 출원 엑셀·첨부 다운로드**는 선택한 한 건을, 월별 다운로드는 선택한 운영월과 조회 범위의 전체 자료를 문서 폴더 `AI검색평가`에 저장합니다. Excel은 출원별 현황·선행기술 목록·품질 순위·문헌별 평가 네 시트이며 심사국·심사과·CPC도 표시합니다. ZIP에는 Excel과 첨부문헌이 포함됩니다.

제출 당시 심사국·심사과와 CPC를 보관하므로 이후 사용자 소속을 바꿔도 기존 출원의 소속은 바뀌지 않습니다. 이전 데이터에 새 항목이 없으면 공란으로 표시합니다. 기존 납품 문헌의 명칭 등 보관 정보도 유지합니다.

저장 충돌이 나면 입력내용을 보존한 채 최신 판을 확인하세요. 저장 직후 연결이 끊긴 경우 같은 입력으로 재시도하면 저장 요청 ID를 확인합니다. 같은 ID에 다른 내용을 보내면 거절합니다. 확정 JSON이나 `.lock` 파일은 프로그램이 관리합니다.

## 폴더와 공유 설정

| 위치 | 내용 |
|---|---|
| `App/Start.ps1`, `App/Views/Main.xaml` | 로그인·실행 화면과 이벤트 처리 |
| `App/Modules/Storage.psm1` | 배타 잠금, 버전 확인, UTF-8 JSON 확정 |
| `App/Modules/Workflow.psm1` | 사용자·월별 배정·제출·납품·평가 규칙 |
| `App/Modules/Export.psm1` | Excel·ZIP 생성 |
| `Master/{Users,Groups,Sources}/ID` | 기준정보 |
| `Periods/YYYY-MM` | 참여자·담당자·기한·마감 상태 |
| `Data/월/그룹/Submissions/사용자/건` | 제출정보 |
| `Data/월/그룹/Deliveries/조사방식/사용자/건` | 납품자료 |
| `Data/월/그룹/Evaluations/사용자/건` | 심사관 평가 |
| `Data/월/그룹/Attachments/조사방식/사용자/건` | 첨부문헌 |
| `Locks/월` 및 각 기록의 `write.lock` | 저장·등록·월 상태 변경 조정 |

각 기록은 `revisions/000001.json`부터 전체 스냅샷으로 추가됩니다. 확정 전 임시파일은 조회하지 않습니다. 잠금은 파일의 존재 여부가 아니라 열려 있는 파일 핸들로 판정합니다.

배포 프로그램은 참여자에게 읽기·실행 권한을, 공유 저장소는 참여자의 실제 Windows 계정에 프로그램 운영에 필요한 읽기·쓰기 권한을 부여합니다. 사용자 선택은 프로그램 안의 작성자와 업무 범위를 정하고, 파일서버 접근은 접속 중인 Windows 계정의 권한을 사용합니다. 선택한 사용자에 맞춰 NTFS 권한이 바뀌지는 않습니다. 별도 백업은 백업 운영 계정으로 관리합니다.

## 점검과 백업

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\Tests\Run-All.ps1
pwsh.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\Tests\Run-All.ps1
# 실제 창을 열어 화면 동작까지 검사하려면 추가:
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\Tests\Run-All.ps1 -IncludeGui
```

로컬 테스트는 별도 프로세스의 동시 저장·오래된 버전 거절·동일 요청 재시도·잠금 해제·JSON 손상·로그인 선택·업무 규칙·내보내기·운영 초기화와 배포를 검사합니다. 기본 검증은 창을 표시하지 않으며 `-IncludeGui`가 실제 창 실행 검증을 추가합니다. 실제 SMB/NAS에서 두 PC 동시 사용과 네트워크 단절 중 저장·복구는 해당 환경에서 확인해야 합니다.

버전 이력과 별도로 데이터·첨부·월 설정을 함께 백업하고, 월 마감본을 보관하세요. 파일서버 스냅샷과 별도 백업을 사용하고 출원 한 건의 실제 복구를 시험합니다.

관리자 문헌 일괄 등록, 진행 중인 건의 그룹 이관, 조사 방식 간 동일 문헌·인용 불일치 자동 안내는 아직 포함하지 않습니다.
