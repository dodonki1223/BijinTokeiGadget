echo off

Rem 注意注意注意注意注意注意注意注意注意注意注意注意注意注意
Rem ・改行コードが「LF」のみになっていると正しく動作しないため、
Rem   改行コードを「CR+LF」にすること
Rem   githubからダウンロードしてくるとLFになってしまう模様
Rem 注意注意注意注意注意注意注意注意注意注意注意注意注意注意


Rem ----------------------- 設定 -----------------------
Rem コンパイル対象ソースフォルダ設定 
Rem ※コンパイル対象ソースのデフォルトの設定です
Set CompileSourceFolder=%~dp0
Rem slnファイル名                    
Rem ※ソリューションファイル名のデフォルトの設定です
Set SlnFileName=BijinTokeiGadget.sln
Rem .NetFramework Version指定        
Rem ※.NetFrameworkのバージョン指定のデフォルト設定です 「%windir%\Microsoft.NET\Framework\」配下のフォルダ名を指定して下さい
Set FrameworkVersion=v4.0.30319
Rem ----------------------- 設定 -----------------------

Rem タイトル表示
:DisplayTitle

    echo;
    echo  ********************************************************
    echo    美人時計ガジェットコンパイル用バッチ
    echo      美人時計ガジェットのリビルドを行いexeファイルを作成します
    echo *********************************************************

Rem フォルダ・ファイル指定処理
:SpecifyCompileFolderAndSlnFile

    echo;
    echo  ********************************************************
    echo  コンパイルフォルダ・ソリューションファイル指定処理
    echo  ********************************************************

    echo;
    echo ★コンパイル対象ソースフォルダを指定してください
    echo ※何も指定しない場合は：「%CompileSourceFolder%」フォルダになります
    echo;
    Set /p CompileSourceFolder="フォルダの入力　＞　"

    echo;
    echo ★ソリューションファイル名を指定してください
    echo ※何も指定しない場合は：「%SlnFileName%」ファイルになります
    echo;
    Set /p SlnFileName="ファイル名の入力　＞　"

Rem 設定の確認
:DisplayConfiguration

    echo;
    echo  ********************************************************
    echo  コンパイルフォルダ　　　：%CompileSourceFolder%
    echo  ソリューションファイル名：%SlnFileName%
    echo  ********************************************************
    echo;
    
    echo ★下記メッセージは(「Y」又は「y」)以外はキャンセルされます
    Set /p RunContinueResult="上記情報を使用して処理を実行しますか？(y/n)　＞　"

    Rem 大文字/小文字変換(Y以外は全てキャンセル扱い) 
    Set RunContinueResult=%RunContinueResult:y=Y%%
    
    Rem Y以外の入力の時はコンパイルフォルダ・ソリューションファイル指定処理へ
    If /i Not %RunContinueResult%==Y Goto SpecifyCompileFolderAndSlnFile

Rem ソースのコンパイル処理
:CompileProcess

    echo;
    echo  ********************************************************
    echo  コンパイル処理
    echo  ********************************************************
    
    echo;
    echo  ★カレントフォルダを変更します……
    echo;
    cd %CompileSourceFolder%

    echo;
    echo  ★コンパイルを実行します……
    echo;
    %windir%\Microsoft.NET\Framework\%FrameworkVersion%\MSBuild.exe %SlnFileName% /t:Rebuild /p:Configuration=Release

    echo;
    echo  ★処理が完了しました！！
    echo;

    pause

Rem 終了処理
:EndProcess

    Rem処理の終了
    exit