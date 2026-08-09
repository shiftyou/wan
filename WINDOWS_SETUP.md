# Windows 설치 및 자동 실행 설정

이 문서는 이 프로젝트(`qr_generator.py`, `qr_classifier.py`, `list_updator.py`)를
Windows PC(네트워크 드라이브 `Z:` 기준)에서 설치하고, `qr_classifier.py`가
로그인할 때마다 자동으로 실행되게 설정하는 방법을 설명합니다.

`qr_generator.py`, `qr_classifier.py`, `list_updator.py` 모두 화면(GUI)이 있는
프로그램입니다. Python 설치 없이 쓰고 싶다면 exe로 빌드된 버전을 그대로
복사해서 실행해도 됩니다 — 이 경우 1~3단계(Python 설치, 가상환경)는
건너뛰어도 됩니다. 권장 배포 형태는 `dist/PatientTools` 폴더(exe 3개가
`_internal` 폴더 하나를 공유, 빌드 방법은 7번 참고)이며, 이 폴더 전체를
그대로 옮기면 됩니다.

## 0. 폴더 구조

이 프로젝트는 아래 3개 폴더를 기준으로 동작하도록 기본값이 설정되어 있습니다
(전부 `Z:` 드라이브 바로 아래).

| 폴더 | 역할 |
| --- | --- |
| `Z:\99_사진작업폴더` | 파이썬 스크립트(`qr_generator.py`, `qr_classifier.py`, `list_updator.py`) 또는 exe와 `.venv`, `current_session.json`, `qr_generator_config.json`, `qr_classifier_config.json`, `records.csv` 등 작업 파일이 위치. 그 아래 `photo` 하위폴더(`Z:\99_사진작업폴더\photo`)가 EOS Utility의 사진 저장 위치(감시 폴더 기본값) |
| `Z:\01_환자이름별사진` | QR을 읽어 `홍보여부_이름_수술명_수술날짜` 세션 폴더가 만들어지고, 그 아래 `홍보여부_이름_#경과일` 폴더에 사진이 분류되어 저장됨(분류 대상 폴더 기본값). 홍보 환자 엑셀(`완성형_홍보환자_리스트.xlsx`)도 이 폴더에 위치 |
| `Z:\02_날짜별사진` | 분류 전 원본 사진을 촬영일 기준 `년/년월/년월일` 구조로 그대로 백업(원본 백업 폴더 기본값) |

이 세 폴더(감시 폴더/분류 대상 폴더/원본 백업 폴더)는 `qr_classifier.py` 화면에서
직접 입력하거나 "찾아보기"로 바꿀 수 있습니다. 마지막으로 바꾼 값은
`qr_classifier_config.json`에 저장되어 다음 실행 때도 그대로 불러옵니다(환경변수는
더 이상 쓰지 않습니다). `qr_generator.py`의 저장 폴더도 같은 방식으로
`qr_generator_config.json`에 저장됩니다.

## 1. Python 설치

1. https://www.python.org 에서 Python 3.11 이상 설치 파일을 받습니다.
2. 설치 화면 첫 페이지에서 **"Add python.exe to PATH"** 체크박스를 반드시 켭니다.
3. 설치 후 명령 프롬프트(cmd)에서 `python --version`으로 정상 설치를 확인합니다.

## 2. 프로젝트 폴더 준비

1. 이 프로젝트 폴더 전체를 `Z:\99_사진작업폴더`에 복사합니다.
2. `Z:\01_환자이름별사진`, `Z:\02_날짜별사진` 폴더를 미리 만들어 둡니다
   (없어도 실행 시 자동 생성되지만, 네트워크 드라이브 권한을 미리 확인하는
   차원에서 먼저 만들어 두는 것을 권장합니다).
3. `완성형_홍보환자_리스트.xlsx` 파일(엑셀 서식이 이미 들어있는 관리대장)을
   `Z:\01_환자이름별사진` 안에 넣습니다.
4. EOS Utility의 사진 저장 폴더를 `Z:\99_사진작업폴더\photo`로 지정합니다.

기존 macOS용 `.venv` 폴더는 그대로 복사해도 Windows에서 쓸 수 없습니다
(가상환경은 운영체제/경로에 종속적입니다). 복사했더라도 삭제하고 아래처럼
새로 만듭니다.

## 3. 가상환경 생성 및 패키지 설치

`Z:\99_사진작업폴더`에서 명령 프롬프트를 열고:

```bat
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

## 4. 각 스크립트 실행

macOS의 `qr.sh`/`qr_classifier.sh`/`xlsx.sh` 대신 Windows용 `.bat` 파일을 씁니다.
`Z:\99_사진작업폴더`에서 더블클릭하거나 명령 프롬프트에서 실행합니다.
(exe 버전을 쓴다면 `.bat` 대신 `QR_Generator.exe`/`QR_Classifier.exe`/
`List_Updator.exe`를 바로 더블클릭하면 됩니다.)

- `qr.bat` — QR 코드 생성기 화면 실행 (환자 정보 입력, QR 생성/찾기)
- `qr_classifier.bat` — 사진 자동 분류기 화면 실행. 창을 열면 바로 감시가 시작되고,
  진행 상황이 창 안 로그 영역에 실시간으로 표시됩니다("감시 중지" 버튼으로
  멈출 수 있습니다).
- `xlsx.bat` — 홍보 환자 엑셀 백필 화면 실행 (기존 폴더 일괄 백필용, 평소엔
  `qr_classifier.py`가 세션 시작 시점에 자동으로 등록하므로 수동 실행할 필요는
  거의 없음). 대상 폴더/엑셀 파일을 지정하고 "실행" 버튼을 누르면 됩니다.

## 5. 한글 폰트

QR 이미지 하단에 페이로드를 한글로 표시하기 위해 Windows 기본 폰트인
`C:\Windows\Fonts\malgun.ttf`(맑은 고딕)을 자동으로 사용하도록 이미 설정되어
있습니다. 별도 설치가 필요 없습니다.

## 6. 로그인 시 자동 실행

`qr_classifier.py`(또는 `QR_Classifier.exe`)는 창을 열면 화면 안 체크박스
"실행하면 자동으로 감시 시작"이 켜져 있을 때만 사람이 버튼을 누르지 않아도
곧바로 감시를 시작합니다. **기본값은 꺼짐**이라 처음 설치했을 때는 창만 뜨고
감시는 시작되지 않습니다 — 폴더 설정을 먼저 확인한 뒤 이 체크박스를 한 번
켜 두세요. 이 체크박스 상태와 폴더 설정은 `qr_classifier_config.json`에
저장되어 다음 실행 때도 그대로 유지됩니다(한 번 켜 두면 그다음부터는 계속
자동 시작됩니다).

로그인할 때 이 프로그램이 자동으로 뜨게 하려면 일반적인 Windows 시작프로그램
바로가기만 등록하면 됩니다(예전처럼 콘솔/pythonw 트릭이 필요 없습니다).

1. 먼저 `qr_classifier.py`(또는 exe)를 한 번 실행해서 폴더 설정을 확인하고
   "실행하면 자동으로 감시 시작" 체크박스를 켭니다.
2. `Win + R` → `shell:startup` 입력 후 엔터를 누르면 시작프로그램 폴더가
   열립니다.
3. 아래 둘 중 하나의 바로가기를 그 폴더에 만듭니다.
   - exe 버전: `QR_Classifier.exe`
   - 소스 버전: `qr_classifier.bat` (내부적으로 `pythonw.exe`로 실행되어 콘솔 창이
     뜨지 않습니다)
4. 다음 로그인부터는 이 바로가기가 자동 실행되어 `qr_classifier.py`
   화면이 뜨고, 체크박스를 켜 뒀으므로 바로 감시를 시작합니다.

지금 바로 테스트해보려면 시작프로그램 폴더에 넣은 바로가기를 더블클릭하면
됩니다 (재부팅하지 않아도 즉시 실행됩니다).

### 정상 동작 확인 / 문제 해결

- 창이 뜨고 "감시 중" 상태가 보이면 정상입니다. 진행 상황(`새로운 환자 시작`,
  `사진 저장` 등)은 창 안 "실행 로그" 영역에서 스크롤하며 확인할 수 있습니다
  (더 이상 로그 파일을 따로 열어볼 필요가 없습니다).
- 자동 실행을 끄고 싶다면 시작프로그램 폴더(`shell:startup`)에서
  바로가기만 지우면 됩니다. 창은 그대로 두고 감시만 잠시 끄고 싶다면
  창 안의 "감시 중지" 버튼을 누르면 됩니다.
- `Z:` 드라이브가 로그인 직후 아직 연결되지 않았을 수 있어서,
  감시 폴더/분류 대상 폴더에 접근할 수 있을 때까지 자동으로 재시도합니다.
  로그에 "접근 대기 중..." 메시지가 반복되면 `Z:` 드라이브 매핑이 안 된
  것이니 네트워크 연결/드라이브 매핑을 먼저 확인하세요.
- 로그인 전(부팅 직후)부터 실행되게 하고 싶다면 Windows 작업 스케줄러에서
  트리거를 "시스템 시작 시"로, "사용자 로그온 여부와 관계없이 실행"으로
  등록하는 방법도 있지만, 관리자 권한과 계정 비밀번호 저장이 필요해서
  기본으로는 위 시작프로그램 방식을 권장합니다.

## 7. exe 빌드 방법 (개발자용)

exe는 Windows에서만 빌드할 수 있습니다(PyInstaller는 크로스 컴파일을 지원하지
않음 — Mac/Linux에서 빌드하면 그 OS용 실행파일만 나옵니다). `dist/` 폴더에
결과물이 생기며, 이 폴더는 git에 올라가지 않으니 배포할 땐 여기서 직접
복사해서 씁니다.

0. 가상환경에 PyInstaller를 한 번 설치합니다.

   ```bat
   .venv\Scripts\pip install pyinstaller
   ```

### 권장 방법: 세 도구를 한 폴더로 묶어서 빌드 (`patient_tools.spec`)

`--onefile`은 실행할 때마다 임시 폴더에 압축을 풀기 때문에 창이 뜨는 데
시간이 걸립니다. `--onedir`(폴더형)은 이 과정이 없어 훨씬 빠르지만, 세
도구를 각각 onedir로 만들면 opencv/numpy 같은 무거운 의존성이 도구마다
중복 저장되어 용량이 커집니다(따로 만들면 합쳐서 약 350MB, 아래 방법은
약 225MB).

저장소에 있는 `patient_tools.spec`은 PyInstaller의 "Multi-Package Bundle"
기능(`MERGE()`)으로 세 스크립트를 하나의 `_internal` 폴더를 공유하도록
묶어 빌드합니다. 결과물은 `dist/PatientTools/` 안에 `QR_Generator.exe`,
`QR_Classifier.exe`, `List_Updator.exe`가 같은 폴더에, `_internal/`을
공유하는 형태로 생깁니다.

`patient_tools.spec` 안의 `PYZBAR_DIR` 값(pyzbar 네이티브 DLL 위치)은
빌드하는 컴퓨터의 실제 경로로 맞춰야 합니다. 확인 방법:

```bat
.venv\Scripts\python -c "import pyzbar, os; print(os.path.dirname(pyzbar.__file__))"
```

경로를 확인해 `.spec` 파일의 `PYZBAR_DIR` 줄을 수정한 뒤 빌드합니다.

```bat
.venv\Scripts\pyinstaller patient_tools.spec
```

빌드 후에는 회전되었거나 흐린 QR 사진으로 `QR_Classifier.exe`를 한 번
테스트해 보세요 — pyzbar DLL이 제대로 안 들어가면 크래시 없이 인식률만
조용히 떨어져서 눈치채기 어렵습니다.

배포할 땐 `dist/PatientTools` **폴더 전체**를 옮겨야 합니다(`_internal`
없이 exe만 옮기면 실행이 안 됩니다).

### 개별 빌드 (도구 하나만 필요할 때)

```bat
.venv\Scripts\pyinstaller --onefile --windowed --name QR_Generator qr_generator.py
.venv\Scripts\pyinstaller --onefile --windowed --name List_Updator list_updator.py
.venv\Scripts\pyinstaller --onefile --windowed --name QR_Classifier ^
  --add-binary "<PYZBAR_DIR>\libzbar-64.dll;pyzbar" ^
  --add-binary "<PYZBAR_DIR>\libiconv.dll;pyzbar" ^
  qr_classifier.py
```

`--onefile`은 exe 파일 하나만 옮기면 되는 대신 실행할 때마다 압축을 풀어서
느립니다. 빠른 실행이 필요하면 위 "권장 방법"을 쓰세요.

소스 코드를 고칠 때마다 해당 exe(또는 `patient_tools.spec` 빌드 전체)를
다시 빌드해야 반영됩니다. `build/` 폴더는 캐시 용도라 지워도 됩니다.
