@echo off
setlocal enabledelayedexpansion

:: アクセスリストのファイル名
set "listfile=アクセスリスト.txt"

:: 現在選択中のカテゴリ（初期は0:直接アクセス）
set "current_category=0"

:: カテゴリ名称を初期化
set "category_name_0="
set "category_name_1="
set "category_name_2="
set "category_name_3="
set "category_name_4="
set "category_name_5="

:: アクセスリストの存在チェック
if not exist "%listfile%" (
    echo アクセスリスト "%listfile%" が見つかりません。
    pause
    exit /b
)

:: カテゴリ名称をアクセスリストから読み込む
for /f "usebackq tokens=1,2,* delims= " %%A in ("%listfile%") do (
    if "%%A"=="#" (
        if "%%B"=="CATEGORY" (
            for /f "tokens=1,* delims= " %%X in ("%%C") do (
                set "category_name_%%X=%%Y"
            )
        )
    )
)

:main_menu
cls
echo ========================================
if "!current_category!"=="all" (
    echo OpenFolder - 全項目表示
) else (
    echo OpenFolder - !category_name_%current_category%!
)
echo ========================================
echo.

:: 現在のカテゴリの項目を表示
echo 利用可能な項目:
set "item_count=0"

if "!current_category!"=="all" (
    :: 全カテゴリ表示時はカテゴリ名も表示
    for /f "usebackq tokens=1,2,* delims= " %%A in ("%listfile%") do (
        if not "%%A"=="#" (
            call :get_category_name %%A
            echo   [!temp_category_name!] %%B : %%C
            set /a item_count+=1
        )
    )
) else (
    :: 特定カテゴリ表示
    for /f "usebackq tokens=1,2,* delims= " %%A in ("%listfile%") do (
        if not "%%A"=="#" (
            if "%%A"=="!current_category!" (
                echo   %%B : %%C
                set /a item_count+=1
            )
        )
    )
)

:: 項目が見つからない場合
if !item_count! equ 0 (
    echo   このカテゴリには項目がありません。
)

echo.

:: カテゴリ選択メニューを表示（常に表示）
echo カテゴリ選択:
if not "!current_category!"=="0" (
    echo   0 : !category_name_0!
)
if not "!current_category!"=="1" (
    echo   1 : !category_name_1!
)
if not "!current_category!"=="2" (
    echo   2 : !category_name_2!
)
if not "!current_category!"=="3" (
    echo   3 : !category_name_3!
)
if not "!current_category!"=="4" (
    echo   4 : !category_name_4!
)
if not "!current_category!"=="5" (
    echo   5 : !category_name_5!
)
if not "!current_category!"=="all" (
    echo   a : 全項目表示
)

echo.
set /p key=選択してください: 

:: カテゴリ切り替えの処理
if "!key!"=="0" (
    set "current_category=0"
    goto main_menu
)
if "!key!"=="1" (
    set "current_category=1"
    goto main_menu
)
if "!key!"=="2" (
    set "current_category=2"
    goto main_menu
)
if "!key!"=="3" (
    set "current_category=3"
    goto main_menu
)
if "!key!"=="4" (
    set "current_category=4"
    goto main_menu
)
if "!key!"=="5" (
    set "current_category=5"
    goto main_menu
)
if /i "!key!"=="a" (
    set "current_category=all"
    goto main_menu
)

:: 入力されたキーに対応するパスを検索
set "target="
set "found_category="

if "!current_category!"=="all" (
    :: 全カテゴリから検索
    for /f "usebackq tokens=1,2,* delims= " %%A in ("%listfile%") do (
        if not "%%A"=="#" (
            if /i "%%B"=="!key!" (
                set "target=%%C"
                set "found_category=%%A"
            )
        )
    )
) else (
    :: まず現在のカテゴリから検索
    for /f "usebackq tokens=1,2,* delims= " %%A in ("%listfile%") do (
        if not "%%A"=="#" (
            if "%%A"=="!current_category!" (
                if /i "%%B"=="!key!" (
                    set "target=%%C"
                    set "found_category=%%A"
                )
            )
        )
    )
    
    :: 現在のカテゴリで見つからない場合は全カテゴリから検索
    if "!target!"=="" (
        for /f "usebackq tokens=1,2,* delims= " %%A in ("%listfile%") do (
            if not "%%A"=="#" (
                if /i "%%B"=="!key!" (
                    set "target=%%C"
                    set "found_category=%%A"
                )
            )
        )
    )
)

:: 該当項目が見つからない場合
if "!target!"=="" (
    echo.
    echo 入力されたキー "!key!" に対応する項目が見つかりません。
    echo.
    pause
    goto main_menu
)

:: ファイル・フォルダを開く処理
goto open_target

:open_target
:: 拡張子とディレクトリ区分
for %%F in ("!target!") do (
    set "ext=%%~xF"
    set "basefolder=%%~dpF"
)

echo.
:: 実行内容を表示
call :get_category_name !found_category!
set "category_name=!temp_category_name!"

:: 現在のカテゴリと異なる場合は明示
if not "!found_category!"=="!current_category!" (
    if not "!current_category!"=="all" (
        echo [!category_name!から実行]
    ) else (
        echo [!category_name!]
    )
) else (
    echo [!category_name!]
) 

:: 拡張子が .ps1 の場合：そのフォルダにcdしてから実行
if /i "!ext!"==".ps1" (
    echo PowerShell スクリプトを実行します...
    echo ファイル: !target!
    echo.
    pushd "!basefolder!"
    powershell.exe -ExecutionPolicy Bypass -NoProfile -File "!target!"
    popd
) else (
    echo エクスプローラーで開いています...
    echo パス: !target!
    echo.
    start "" "!target!"
)

echo.
echo 処理が完了しました。

goto main_menu

:get_category_name
:: カテゴリ番号から名称を取得するサブルーチン
set "temp_category_name=!category_name_%1!"
goto :eof