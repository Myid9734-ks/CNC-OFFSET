@echo off
chcp 65001 > nul
echo =====================================
echo FanucFocasTutorial 배포 스크립트
echo =====================================
echo.

set SOURCE="%~dp0FanucFocasTutorial\bin\Release"

REM 배포할 위치를 입력받습니다
echo 배포할 폴더 경로를 입력하세요 (예: C:\Users\사용자명\Desktop\Release)
echo 또는 Enter를 누르면 기본 경로를 사용합니다: %USERPROFILE%\Desktop\Release
echo.
set /p DEST="배포 경로: "
if "%DEST%"=="" set DEST="%USERPROFILE%\Desktop\Release"

echo.
echo 소스 폴더: %SOURCE%
echo 대상 폴더: %DEST%
echo.

REM 대상 폴더가 없으면 생성
if not exist %DEST% (
    echo 대상 폴더를 생성합니다...
    mkdir %DEST%
)

REM x86 폴더 생성
if not exist %DEST%\x86 (
    mkdir %DEST%\x86
)

echo 파일을 복사하는 중...
echo.

REM 실행 파일과 DLL 복사
copy /Y %SOURCE%\FanucFocasTutorial.exe %DEST%\
copy /Y %SOURCE%\FanucFocasTutorial.exe.config %DEST%\
copy /Y %SOURCE%\EntityFramework.dll %DEST%\
copy /Y %SOURCE%\EntityFramework.SqlServer.dll %DEST%\
copy /Y %SOURCE%\System.Data.SQLite.dll %DEST%\
copy /Y %SOURCE%\System.Data.SQLite.EF6.dll %DEST%\
copy /Y %SOURCE%\System.Data.SQLite.Linq.dll %DEST%\
copy /Y %SOURCE%\Fwlib32.dll %DEST%\

REM x86 네이티브 라이브러리 복사
copy /Y %SOURCE%\x86\SQLite.Interop.dll %DEST%\x86\

echo.
echo =====================================
echo 배포 완료!
echo =====================================
echo.
echo 배포 위치: %DEST%
echo.
pause
