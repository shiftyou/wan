# -*- mode: python ; coding: utf-8 -*-
"""세 도구(QR 생성기/분류기/엑셀 백필)를 하나의 _internal 폴더를 공유하는
형태로 묶어서 빌드한다 (PyInstaller의 Multi-Package Bundle 기능).

빌드: pyinstaller patient_tools.spec
결과물: dist/PatientTools/ 안에 exe 3개 + 공유 _internal/ 폴더 하나
"""

import os

PYZBAR_DIR = r"C:\Users\shiftyou\AppData\Local\Python\pythoncore-3.14-64\Lib\site-packages\pyzbar"

a_generator = Analysis(
    ["qr_generator.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

a_classifier = Analysis(
    ["qr_classifier.py"],
    pathex=[],
    binaries=[
        (os.path.join(PYZBAR_DIR, "libzbar-64.dll"), "pyzbar"),
        (os.path.join(PYZBAR_DIR, "libiconv.dll"), "pyzbar"),
    ],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

a_updator = Analysis(
    ["list_updator.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

MERGE(
    (a_generator, "qr_generator", "QR_Generator"),
    (a_classifier, "qr_classifier", "QR_Classifier"),
    (a_updator, "list_updator", "List_Updator"),
)

pyz_generator = PYZ(a_generator.pure, a_generator.zipped_data)
pyz_classifier = PYZ(a_classifier.pure, a_classifier.zipped_data)
pyz_updator = PYZ(a_updator.pure, a_updator.zipped_data)

exe_generator = EXE(
    pyz_generator, a_generator.scripts, [], exclude_binaries=True,
    name="QR_Generator", console=False,
)
exe_classifier = EXE(
    pyz_classifier, a_classifier.scripts, [], exclude_binaries=True,
    name="QR_Classifier", console=False,
)
exe_updator = EXE(
    pyz_updator, a_updator.scripts, [], exclude_binaries=True,
    name="List_Updator", console=False,
)

coll = COLLECT(
    exe_generator, a_generator.binaries, a_generator.zipfiles, a_generator.datas,
    exe_classifier, a_classifier.binaries, a_classifier.zipfiles, a_classifier.datas,
    exe_updator, a_updator.binaries, a_updator.zipfiles, a_updator.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="PatientTools",
)
