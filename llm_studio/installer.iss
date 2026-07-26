; LocalLLM Studio 인스톨러 (Inno Setup 6)
; 컴파일: Inno Setup Compiler에서 이 파일을 열고 Build → Compile
;
; 하는 일:
;   1. dist\LocalLLMStudio\* 를 Program Files에 설치
;   2. C:\ProgramData\LocalLLMStudio 데이터 폴더를 만들고
;      Users 그룹에 수정(Modify) 권한을 부여 — 관리자가 아닌 계정으로
;      실행해도 모델/설정/대화기록을 쓸 수 있게 하는 핵심 부분
;   3. 바탕화면/시작메뉴 바로가기 생성

#define AppName "LocalLLM Studio"
#define AppVersion "1.0.0"
#define AppExeName "LocalLLMStudio.exe"
#define DataDirName "LocalLLMStudio"

[Setup]
AppId={{7E1B8A52-4C0D-4B4E-9F3A-2D8E51C0A7B4}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={autopf}\{#DataDirName}
DefaultGroupName={#AppName}
OutputBaseFilename=LocalLLMStudio-Setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
; ProgramData 권한 부여에 관리자 권한이 필요하다
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Files]
Source: "dist\LocalLLMStudio\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Dirs]
; 데이터 폴더 생성 + Users 그룹에 수정 권한 부여 (하위 폴더로 상속됨)
Name: "{commonappdata}\{#DataDirName}"; Permissions: users-modify
Name: "{commonappdata}\{#DataDirName}\models"
Name: "{commonappdata}\{#DataDirName}\conversations"
Name: "{commonappdata}\{#DataDirName}\uploads"
Name: "{commonappdata}\{#DataDirName}\logs"

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\모델 폴더 열기"; Filename: "{commonappdata}\{#DataDirName}\models"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "바탕화면 바로가기 만들기"; GroupDescription: "추가 작업:"

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{#AppName} 지금 실행"; \
  Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 프로그램만 지우고 데이터(모델/대화기록)는 남긴다.
; 데이터까지 지우려면 아래 주석을 해제:
; Type: filesandordirs; Name: "{commonappdata}\{#DataDirName}"
