# QR 촬영 세션 운영 방법

1. EOS Utility의 저장 폴더를 이 프로젝트의 `photo` 폴더(또는 NAS에 마운트한 수신 폴더)로 지정합니다.
2. 가상환경에서 의존성을 한 번 설치합니다.

   ```bash
   .venv/bin/python -m pip install -r requirements.txt
   ```

3. 환자 이름으로 촬영 시작 QR을 만듭니다.

   ```bash
   .venv/bin/python qr_session_generator.py
   ```

4. `qr_codes/환자명.png`을 태블릿 또는 인쇄물로 보여 주고 첫 장으로 촬영합니다.
5. `ocr_classifier.py`를 실행하고, 이어서 해당 환자의 사진을 촬영합니다.

   ```bash
   .venv/bin/python ocr_classifier.py
   ```

6. 다음 환자가 오면 그 환자의 QR을 먼저 촬영합니다. 새 QR을 읽는 즉시 이전 세션은 자동 종료되고, 이후 사진은 새 환자 폴더에 저장됩니다.

## 결과 폴더

- `photo/환자명/`: 해당 세션의 사진
- `photo/_세션마커/환자명/`: 해당 환자의 시작 QR이 촬영된 사진
- `photo/_미분류/`: 시작 QR 전에 촬영했거나 QR 오류가 난 사진

사진 파일이 `STABLE_SECONDS`(기본 3초) 동안 변하지 않아야 처리합니다. EOS Utility의 Wi-Fi 전송이 끝나기 전에 파일을 읽지 않도록 하기 위한 설정입니다.

## 개인정보 주의

QR과 생성되는 폴더명에 환자 이름이 포함됩니다. NAS 공유 권한을 진료 인력으로 제한하고, 외부 공유 링크를 사용하지 마세요. NAS 백업도 암호화하는 것이 좋습니다.
