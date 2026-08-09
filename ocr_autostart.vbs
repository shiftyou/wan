' ocr_classifier.py를 콘솔 창 없이(숨김) 백그라운드로 실행한다.
' 아래 PROJECT_DIR을 실제 설치 경로로 바꾼 뒤, 이 파일의 바로가기를
' Windows 시작프로그램 폴더(Win+R -> shell:startup)에 넣으면
' 로그인할 때마다 자동으로 조용히 실행된다.
' 실행 로그/오류는 ocr_classifier.log 에서 확인한다.

Dim projectDir, quote, command, objShell

projectDir = "Z:\99_사진작업폴더"
quote = Chr(34)
command = "cmd /c " & quote & ".venv\Scripts\pythonw.exe" & quote & _
          " ocr_classifier.py >> ocr_classifier.log 2>&1"

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = projectDir
objShell.Run command, 0, False
