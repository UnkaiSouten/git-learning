#Persistent
#SingleInstance, Force
#NoEnv
#UseHook
#InstallKeybdHook
#InstallMouseHook
#HotkeyInterval, 2000
#MaxHotkeysPerInterval, 200
Process, Priority,, Realtime
SendMode, Input
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2

;==Ｆ6でChromeを開くがあふ====
;=あふF6をAHKが奪わないようにする設定
;Chromeを開く
#IfWinNotActive ahk_exe AFXW.EXE
    ; --- AFXW 以外でのみ F6 を有効化 ---
    F6::
        Run, "C:\Program Files\Google\Chrome\Application\chrome.exe"
    return
#If

;==V1 OR V2====
;MsgBox, %A_AhkVersion%

; ===== テンキー入力の場合は日本語入力でも半角であることを利用
; ======あふ123をAHKが奪わないようにする設定
#IfWinNotActive ahk_exe AFXW.EXE
    1::Send, {Numpad1}
    2::Send, {Numpad2}
    3::Send, {Numpad3}
    4::Send, {Numpad4}
    5::Send, {Numpad5}
    6::Send, {Numpad6}
    7::Send, {Numpad7}
    8::Send, {Numpad8}
    9::Send, {Numpad9}
    0::Send, {Numpad0}

#IfWinNotActive

; ===== AHK v1 （chromeとEdge 閲覧情報表示ctrl+shift+a →alt+q）=====

; Chrome 用
#IfWinActive ahk_exe chrome.exe
    !q::
        Sleep, 20
        SendInput, {Ctrl Down}{Shift Down}a{Shift Up}{Ctrl Up}
    return
#IfWinActive

; Edge 用
#IfWinActive ahk_exe msedge.exe
    !q::
        Sleep, 20
        SendInput, {Ctrl Down}{Shift Down}a{Shift Up}{Ctrl Up}
    return
#IfWinActive

; Ctrl + Shift + z で実行
; Googleブックマークのファイルパス取得
^+z::
    ; アクティブウィンドウ（Chrome）のタイトルを取得
    WinGetTitle, title, A

    ; 「 のインデックス」より前を抽出
    RegExMatch(title, "^(.*?)(?= のインデックス)", m)

    ; 抽出結果がある場合のみ処理
    if (m1 != "")
    {
        path := m1

        ; 末尾に \ が無ければ付与（安全性向上）
        if (SubStr(path, 0) != "\")
            path .= "\"

        ; クリップボードへ格納
        Clipboard := path

        ; エクスプローラーを開く
        Run, explorer.exe "%path%"
    }
    else
    {
        Clipboard := ""  ; 取得失敗時は空にする
    }

return

;windowキー+Shift+H  はNotionとスプレッドシートでダブル
;解決方法として、#IfWinActiveでアプリごとに分ける
#IfWinActive, Notion
    #+h::Send ^+!h  ; windowキー+Shift+H → Ctrl+Shift+Alt+H（Notion用）
#IfWinActive

#IfWinActive, Google スプレッドシート
    ;スプレッドシート（エクセル）ショートカットキー
    ;最終行に記載すると実行されないので、最前列に記載してみた！
    ; windowキー+ shift +hを Ctrl + Shift + PageUp にマッピング 左シート移動
    #+h::
        Send, ^+{PgUp}
    return

    ; windowキー+  shift + lを Ctrl + Shift + PageDown にマッピング 右シート移動
    #+l::
        Send, ^+{PgDn}
    return
    ; Ctrl + jをAlt +shift+ k にマッピング シート一覧を表示
    ^j::
        Send, !+k
    return
#IfWinActive

;［エクセル］ショートカットキー※
; windowキー + hを Ctrl  + PageUp にマッピング 左シート移動
;#h::
;    Send, ^{PgUp}
;return
; WIｎ + lを Ctrl + PageDown にマッピング 右シート移動
#l::
    Send, ^{PgDn}
return

;一番右のシートに移動
#Persistent
^F12::
    ; Excelを取得
    ExcelApp := ComObjActive("Excel.Application")
    if (ExcelApp) {
        ; 最後のシートを選択
        ExcelApp.Sheets(ExcelApp.Sheets.Count).Select
    } else {
        MsgBox, Excelが開いていません。
    }
return

;一番左のシートに移動
#Persistent
^F11::
    ; Excelを取得
    ExcelApp := ComObjActive("Excel.Application")
    if (ExcelApp) {
        ; 最初のシートを選択
        ExcelApp.Sheets(1).Select
    } else {
        MsgBox, Excelが開いていません。
    }
return

;F12::
;  Send, F12が入力されました
;Return

;共通
;リネーム(r)
vk1D & r::Send,{Blind}{F2}

;エンター
vk1D & f::Send,{Blind}{Enter}

;戻る/進む
vk1D & c::Send,{Blind}!{Left}
vk1D & v::Send,{Blind}!{Right}

; ホットキーを設定（例えばCtrl + 9）
;コピーしたテキストをメモ帳に貼り付け
;空白行を削除

^9::
    ClipWait, 2
    if ErrorLevel
        return

    text := Clipboard

    ; 空白行を削除
    text := RegExReplace(text, "m)^\s*\r?\n", "")

    ; ハイフンだけの行を削除
    text := RegExReplace(text, "m)^\s*-\s*\r?\n", "")

    ; ===== No. + 数字（行末が空白のみ）の前に罫線を追加 =====
    text := RegExReplace(text, "m)^No\.\s*\d+\s*$", "--------------------------------------------------------------------------------------------------------------------------------`n$0")

    ; 罫線を反映した内容をクリップボードへ
    Clipboard := text

    Run, "C:\Program Files (x86)\sakura\sakura.exe"
    WinWaitActive, ahk_exe sakura.exe

    SendInput, ^v
return

;===============
^8::
    ClipWait, 2
    if ErrorLevel
        return

    text := Clipboard

    ; 空白行を削除
    text := RegExReplace(text, "m)^\s*\r?\n", "")
    ; ハイフンだけの行を削除
    text := RegExReplace(text, "m)^\s*-\s*\r?\n", "")

    ; ===== No. + 数字（行末が空白のみ）の前に罫線を追加（テキスト側） =====
    line := "--------------------------------------------------------------------------------------------------------------------------------"
    text := RegExReplace(text, "m)^No\.\s*\d+\s*$", line . "`n$0")

    ; ===== HTML エスケープ処理 =====
    esc := text
    esc := StrReplace(esc, "&", "&amp;")
    esc := StrReplace(esc, "<", "&lt;")
    esc := StrReplace(esc, ">", "&gt;")

    ; ===== HTML側で罫線を太字に変換 =====
    esc := StrReplace(esc, line, "<b>" . line . "</b>")

    Clipboard := text   ; Sakura に貼る用（テキスト）

    ; ===== HTMLテンプレート生成（Shift-JIS / CSS付き） =====
    html =
    (
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<style>
body {
    font-family: "Meiryo", sans-serif;
    margin: 20px;
    background: #f7f7f7;
}
.container {
    background: #fff;
    padding: 20px;
    border-radius: 8px;
    border: 1px solid #ccc;
}
pre {
    white-space: pre-wrap;
    word-wrap: break-word;
    font-size: 14px;
    line-height: 1.6;
}
</style>
</head>
<body>
<div class="container">
<pre>
%esc%
</pre>
</div>
</body>
</html>
    )

    ; ===== 一時ファイルとして保存してブラウザで確認 =====
    tempFile := A_Temp . "\preview_temp.html"
    FileDelete, %tempFile%
    FileAppend, %html%, %tempFile%

    Run, %tempFile%

    ; ===== 保存するかどうか選択 =====
    MsgBox, 4, HTML保存, このHTMLを保存しますか？`n`nYes → 保存する`nNo → 保存しない
    IfMsgBox, Yes
    {
        saveDir := A_Desktop
        FormatTime, now, , yyyyMMdd_HHmmss
        filePath := saveDir . "\cleaned_" . now . ".html"

        FileAppend, %html%, %filePath%
        if ErrorLevel {
            MsgBox, 16, Error, HTMLファイルの保存に失敗しました。`n`nパス:`n%filePath%
        }
    }

    ; ===== Sakura起動 =====
    Run, "C:\Program Files (x86)\sakura\sakura.exe"
    WinWaitActive, ahk_exe sakura.exe
    SendInput, ^v
return

;Windows単語登録Ctrl + F7 →（例えばCtrl + 7）
^7::
    Send ^{F7}
return

;カーソル位置から単語の後ろまで選択(単語単位でコピーしたい場合)
;半角スペースと全角スペースの両方を考慮して、空白の位置を見つけ、その空白より後の文字を除去しています
; ホットキーを設定（例えばCtrl + C）
^+c::           ; Ctrl + C
    {
        ; カーソルの位置から単語の後ろまで選択
        Send, ^+{Right}  ; Ctrl + Shift + Right Arrow
        Sleep, 100

        ; クリップボードにコピー
        Send, ^c  ; Ctrl + C
        Sleep, 100

        ; クリップボードの内容を取得
        ClipSaved := Clipboard
        Clipboard := ""

        ; 空白（半角スペースと全角スペース）の位置を探す
        Loop, Parse, ClipSaved
        {
            if (A_LoopField = " " || A_LoopField = "　")
            {
                ; 空白の位置までの文字列を取得
                StringLeft, ClipSaved, ClipSaved, A_Index - 1
                break
            }
        }

        ; クリップボードに設定
        Clipboard := ClipSaved

        ; クリップボードの内容を表示（デバッグ用）
        MsgBox, %Clipboard%

        return
    }

*Space::Send,{Blind}{Space}

;Alt + F4※- Space は通常の入力キーであり修飾キーで使うと挙動が不安定
!1:: ; Alt + 1
    Send !{F4}
return

;行追加（i,I,a,Aは意味がないので定義しない）
Space & o::
    GetKeyState,state,Shift
    if state=U
    {
        Send,{End}
        Send,{Enter}
    }
    else
    {
        Send,{Home}
        Send,{Enter}
        Send,{Up}
        Send,{End}
    }
return

Space & x::
    GetKeyState,state,Shift
    if state=U
        Send,{BS}
    else
        Send,{Delete}
return

;Space & d::Send,^x
Space & d::
    GetKeyState,state,Shift
    if state=U
        Send,^x
    else
        Send,{Delete}
return

Space & p::
    GetKeyState,state,Shift
    if state=U
        Send,^v
    else
    {
        Send,!e
        Send,s
        Send,v
        Send,{Enter}
    }
return

;無変換キー+変換キーでIME切り替え
;無変換+無変換　→かな入力変換
vk1D & vk1C:: Send, {vkF3sc029}
;無変換+ @　→ALT+ F4
vk1D & @:: Send, !{F4}
;無変換+u →　Ctrl+Backspace（単語削除）
vk1D & u:: Send, ^{BS}
;無変換+y → Backspace を送る
vk1D & y:: Send, { BS }

;エンター
vk1D & Space:: Send, {Blind}{Enter}

;上下左右
;カーソル
;無変換キー + ijkl -> カーソルキー
#IfWinNotActive ahk_exe AFXW.EXE
    Space & j::Send,{Blind}{Left}
    Space & l::Send,{Blind}{Right}
    Space & i::Send,{Blind}{Up}
    Space & k::Send,{Blind}{Down}
    ;Home / End
    Space & h::Send,{Blind}{Home}
    Space & vkBB::Send,{Blind}{end}

    vk1D & l::Send,{Blind}{Right}
    vk1D & i::Send,{Blind}{Up}
    vk1D & j::Send,{Blind}{Left}
    vk1D & k::Send,{Blind}{Down}
    ;Home / End
    vk1D & h::Send,{Blind}{Home}
    vk1D & vkBB::Send,{Blind}{end}
#If
; 無変換 + n で行削除
vk1D & n::
    Send, {Home}+{End}{Del}{Del}
return

vk1D & z::
    {
        Send, {RWin down}
        Sleep, 50
        Send, {Down}

    }

;vk1C & z:: Send, {Blind}{PgUp}
;vk1C & x:: Send, {Blind}{PgDn}
;vk1D は無変換（左側）キー
;vk1C は変換（右側）キー
vk1C & 1::Send, {RWinDown}{ShiftDown}{left}{ShiftUp}{RWinUp}
vk1C & 2::Send, {RWinDown}{ShiftDown}{right}{ShiftUp}{RWinUp}

;Windowsキー＋下矢印：ウィンドウを最小化したい場合
;Windowsー＋上矢印：ウィンドウを最大化したい場合
;vk1D & >:: Send, {Blind}#{Up}
;vk1D & <:: Send, {Blind}#{Down}
; - +記号入力
;vk1D & n:: Send, {+}
;vk1D & m:: Send, {-}

;vk1D & v:: Send, *
;vk1D & f:: Send, /

;vk1C 変換（右側）キー
;カーソル位置の行を削除
;vk1C & z::Send,{Blind}{End}+{Home}{BS}
;カーソル行をコピー
;vk1C & c::Send,{Blind}{End}+{Home}^{c}
;vk1d & Down::Send,{Blind}{End}+{Home}^{c}
;カーソル行を切り取り
vk1d & x::Send,{Blind}{End}+{Home}^{x}
;vk1d & Up::Send,{Blind}{End}+{Home}^{x}
;カーソル位置から前を選択
;vk1C & h::Send,{Blind}+{home}
;vk1d & Left::Send,{Blind}+{home}
;カーソル位置から後を選択
;vk1C & b::Send,{Blind}+{End}
;vk1d & Right::Send,{Blind}+{End}
;カーソル位置から改行 リターン
;vk1C & r::Send,{Blind}{End}{Enter}

;WebのTabの移動関係
vk1d & w:: Send, {Blind}^w
vk1d & e:: Send, {Blind}^{Tab}
vk1d & t:: Send, {Blind}^t
vk1d & q:: Send, {Blind}^+{Tab}
;vk1d & space::Esc無変換+スペースはすぐに押下してしまうので変更！
;vk1D & a::Esc
;Sleep, 2
;Return
;========Alt===========
;Alt & l::AltTab
;Alt & j::ShiftAltTab
;Alt & i::Up
;Alt & k::Down

;Runコマンドで一発設定 Windowキー+n：テキスト Windowキー+n：Wev接続

#q::Run "\\Ft-cs020\サービスEG部_NW設備部門_NSOC7\オペレーション担当\99_個人フォルダ\02_ネットワークG\masamitsu.fukuda\◆◇◆◇◆◇◆◇\01_■日勤業務\DTP\18_Mail"

;デスクトップフォルダを開く　＃はWindowsキー
#e::Run, %A_Desktop%
;OneNote
#o::Run "OneNote"
;タブラカスエクスプローラー
#t::Run "C:\Users\a2020\OneDrive\デスクトップ\20240824_onedribe\te250806\TE64.exe"
;スケジューラ
#s::Run "https://web.drm.ddreams.jp/drm/api/schedule"
;マインドマイスター
#m::Run "www.mindmeister.com/app/map/"
#k::Run "https://keep.google.com/u/0/"

; AutoHotkeyスクリプト
; ブラウザを開いてログインページに移動
#c::Run, chrome.exe https://tosho.nishi.or.jp/search/
WinWaitActive, 図書館ホームページ - Google Chrome ; ウィンドウのタイトルを指定

; ユーザー名を入力
Send, 7025711
Send, {Tab} ; 次の入力フィールドに移動

; パスワードを入力
Send, Kaosaku0602
Send, {Enter} ; ログインボタンを押す
Sleep, 3000

#9::
    FormatTime,Now,,yyyyMMdd
    Send,%Now%
Return
#0::
    FormatTime,Now,,MM/dd
    Send,%Now%
Return

;グーグルを開きたい　Ctrl+Shift+Alt+G　※1行の場合、Returnは省略可能
^+!G::Run, https://www.google.com/

;vk1Cは無変換キー
;vk1D & 8::
;backup:=ClipboardAll
;ClipWait, 2
;Clipboard:="「」"
;ClipWait, 2
;Send, ^v
;Sleep,60
;Clipboard:=backup
;Send, {Left}
;Return

; AltTab
IsAltTabMenu := false
!Tab::
    Send !^{Tab}
    IsAltTabMenu := true
return
; カタカナひらがなローマ字キー2連打でAltTabMenuキーとして割当
vkF2::
    If (A_PriorHotKey == A_ThisHotKey and A_TimeSincePriorHotkey < 1000){
        Send !^{Tab}
        IsAltTabMenu := true
    }
return
; 変換キー2連打でAltTabMenuキーとして割当
vk1C::
    If (A_PriorHotKey == A_ThisHotKey and A_TimeSincePriorHotkey < 1000){
        Send !^{Tab}
        IsAltTabMenu := true
    }
return

; ↓↓↓ ホットストリングの設定↓↓↓
#Hotstring O C ?
; O　終了文字を入力しない
; C 大文字小文字を区別する
; ? 単語の途中であっても発動
#Hotstring EndChars `t
;ホットストリングの変換をタブキーを押した時のみ動作させる

::ahk::AutoHotkey
::ps::Kaosaku0602
::tl::08012345678
::ml::yamada.tarou@gmail.com
::al::alert();{Left}{Left}
::objc::Objective-C
::br::<br>
::red::<font color=red></font>
::blue::<font color=blue></font>
::pow::powershell -ExecutionPolicy RemoteSigned -File "C:\Users\6341666\OneDrive - WEST.NTT.CO.JP\SecuredPC\Desktop\kidou.ps1"
::ori::<details><summary>あれ</summary>これ</details>
::150%::<span style="font-size: 135%">${TM_SELECTED_TEXT}</span>
::ipkvm::ipkvm6341666 Sawayaka

; 右クリック押しながらホイール上下でPageUp,PageDown(右クリックも効いちゃうけど)
~RButton & WheelUp::Send, {Pgup}
~RButton & WheelDown::Send, {PgDn}

/*
ひらがな/カタカナキーにEnterを割り当て
https://note.com/noratetsu/n/nc9322eac8d31
vkF2::Enter ;ひらがな/カタカナキーにEnterを割り当て
*/
;vk1d & space::Enter
;Sleep, 2
;Return
;vk1C & space::Enter
;Sleep, 2
;Return

;========m@@ でメアドを入力（ただしコピペで一気に貼り付ける）
#Hotstring *
#Hotstring O
::m@@::
    Clipboard=masamitsu.fukuda.vp@west.ntt.co.jp
    Send,^v
return
#Persistent
#SingleInstance, Force
    .
    .
    .
SetTitleMatchMode, 2
Return

; ホットキー設定
;リンク化できていない場合、1．パス選択　2．ショートカットで開く！
;本ページの最終行にもっていくと動作しない、ここだと動作するよ！
; Ctrl + Shift + 1（イチ）
^+1::
    clipboard := ""
    Send, ^c
    ClipWait, 2  ; 2秒待機してクリップボードの内容を取得
    clipboard := Trim(clipboard)  ; 余分な空白を削除

    if (clipboard = "") {
        MsgBox, クリップボードが空です。コピーする内容を選択してください。
        return
    }

    if !FileExist(clipboard) {
        MsgBox, クリップボードの内容が有効なパスではありません。
        return
    }

    ; フォルダかファイルかを判別して適切な処理を実行
    if InStr(FileExist(clipboard), "D") {  ; D = Directory (フォルダ)
        Run, explorer.exe %clipboard%
    } else {
        Run, %clipboard%  ; ファイルなら既定のプログラムで開く
    }

return

;ツールチップ設定
my_tooltip_function(str, delay) {
    ToolTip, %str%
    SetTimer, remove_tooltip, -%delay%
}
remove_tooltip:
    ToolTip
Return
remove_tooltip_all:
    SetTimer, remove_tooltip, Off
    Loop, 20
        ToolTip, , , , % A_Index
Return
OnClipboardChange:
    my_tooltip_function("コピー", 300)
Return
^s::
    Send, ^s
    my_tooltip_function("上書き保存", 300)
Return

;toggle設定はここから
;ツールチップで　;コピーしたとき、；上書き保存したときが
;カーソル横に表示される
hotkeys_define(keys, label, OnOff) {
    Loop, PARSE, keys, `,
        Hotkey, %A_LoopField%, %label%, %OnOff%
    Return
}
; キー無効化用ラベル(Hotkeyコマンドでラベルとして指定する)
disable_keys:
Return

; セカンダリキー入力待ちにし、タイムアウトをSetTimerする
toggle_activation(toggle) {
    time_limitation := 2000
    SetTimer, toggle_deactivation, -%time_limitation%
    my_tooltip_function("vk1C + " . toggle . " -> ?", time_limitation)
}

toggle_deactivation:
    toggle := false
    my_tooltip_function("toggle == false", 1000)
    hotkeys_define(keys_all, "disable_keys", "Off")
    SetTimer, toggle_deactivation, Off
    SetTimer, watch_hotkey_done, Off
Return

; セカンダリキーの入力があった場合、toggleをfalseにし、SetTimerしたタイムアウト設定を解除する
watch_hotkey_done:
    new_ThisHotkey := A_ThisHotkey

    ; toggleにはプライマリキー送信時のA_ThisHotkeyが格納されている
    ; 何らかのホットキー(つまりセカンダリキー)が実行されたとき、A_ThisHotkeyが書き換わることを利用する
    If (new_ThisHotkey != toggle)
        Goto, toggle_deactivation
Return

#If, GetKeyState("vk1C", "P") == true
    1::
    2::
    3::
    4::
    5::
    6::
    7::
    8::
    9::
    a::
    b::
    c::
    d::
    e::
    f::
    g::
    h::
    i::
    j::
    k::
    l::
    m::
    n::
    o::
    p::
    q::
    r::
    s::
    t::
    u::
    v::
    w::
    x::
    y::
    z::

        toggle := A_ThisHotkey
        toggle_activation(toggle)
        SetTimer, watch_hotkey_done, 50

        hotkeys_define(keys_all, "disable_keys", "On")
    Return
#If

;変換+A→DならAHKの公式ドキュメントをブラウザで開く

#If, toggle == "a"
    t::Run, "C:\Users\a2020\OneDrive\デスクトップ\AHK\simple 1シンプル♪.ahk"
    Return

#If, toggle == "f"

    ; オペ統合
    o::Run, http://10.79.52.9/nsoc/html/login.php
    ;NSOC7 ［07］故障事例
    ; 1::Run "\\Ft-cs020\サービスEG部_NW設備部門_NSOC7\NSOC7_SG\07共有\共有\02_ネットワークG\【02】NGN監視業務\【07】故障事例"
    ;Outlook
    m::Run "Outlook"
    ;Vscode
    v::Run "C:\Users\a2020\AppData\Local\Programs\Microsoft VS Code\Code.exe"

    ;Teams
    t::Run "https://teams.microsoft.com/v2/"
    ;さくらエディタ
    s::Run "C:\Program Files (x86)\sakura\sakura.exe"
    ;KEYAKI
    k::Run "http://10.79.52.5:3000/"
    ; Autohotkeyリロード
    r::Reload
    ; 実行ファイルのフォルダを開く
    d::Run, %A_ScriptDir%
    q::Run "C:\Users\6341666\OneDrive - WEST.NTT.CO.JP\SecuredPC\Desktop\??????\???t???[?\?t?g\Kokomite\Kokomite.exe"
    ;WinMerge
    w::Run "C:\Program Files\WinMerge\WinMergeU.exe"
    ;IPKVM
    i::Run "C:\Program Files\RC0001\WinClient.exe"

    Return

#If, toggle == "n"

    ;NSOC7［07］故障事例
    k::Run "\\Ft-cs020\サービスEG部_NW設備部門_NSOC7\NSOC7_SG\07共有\共有\02_ネットワークG\【02】NGN監視業務\【07】故障事例"
    ;dall データ連携一覧表示ツール
    d::Run "http://10.79.52.2/~nwguser/tool/daltool/viewer/dal_filelist.php"
    ;RMS解除後の保存先
    6::Run "\\rms-sp\west001\6341666"
    ;RMS解除後
    r::Run "https://rms-to.ntt-west.local/RmsComGui/Servlet"
    ;NWGミーティング
    m::Run "\\ft-cs020\サービスEG部_NW設備部門_NSOC7\NSOC7_SG\07共有\共有\02_ネットワークG\【01】常日勤業務（勉強会、業務整理等）\【01】ミーティング関連\【00】週一ミーティング資料\2024年度"
    Return

#If, toggle == "k"

    ;フレッツ企画担当
    k::Run "\\Ft-cs020\サービスEG部_NW設備部門_NSOC1\NSOC1_SG\32_オペ企フレッツ管理\"

    Return

#If, toggle == "9"

    ;統制　移転WG
    1::Run "\\Ft-cs020\サービスEG部_NW設備部門_NSOC7\NSOC7_SG\07共有\共有\00_全チーム共通\【502】日本橋移転\打合せ"
    ;統制　BCP検討
    2::Run "\\Ft-cs020\サービスEG部_NW設備部門_NSOC7\NSOC7_SG\07共有\共有\00_全チーム共通\【99-2】危機対応(オペ支援班)\BCP検討\2024年度_BCP検討"
    Return

#If, toggle == "g"
    c::Run "C:\Program Files\Google\Chrome\Application\chrome.exe" "https://calendar.google.com/calendar/u/0/r"
    d::Run "C:\Program Files\Google\Chrome\Application\chrome.exe" "https://drive.google.com/drive/my-drive"
    k::Run "C:\Program Files\Google\Chrome\Application\chrome.exe" "https://keep.google.com/#label/00_ちょこっとメモ"
    g::
        Gosub, toggle_deactivation

        ; 多重起動防止
        If (WinExist("ahk_class AutoHotkeyGUI")) {
            Return
        }

        stash := ClipboardAll
        Clipboard :=
        Send, ^c
        ClipWait, 0.05
        clip := Clipboard
        Clipboard := stash
        clip := rm_crlf(clip)

        Gui, Add, Edit, v_str_google w380, %clip%
        Gui, Add, Button, Default, Search
        Gui, Show, Center w400, Google
        Send, {vkF2}
        clip := ""
    Return

    ButtonSearch:
        Gui, Submit
        ; ★ Chrome で検索を開く
        Run, "C:\Program Files\Google\Chrome\Application\chrome.exe" "https://www.google.co.jp/search?q=%_str_google%"
    Return

    2GuiEscape:
    2GuiClose:
        Gui, Destroy
    Return

    ;----変換→Gとしてから、やっぱりやめたい

    ;改行コード除去
    rm_crlf(str) {
        str := RegExReplace(str, "\n", "")
        str := RegExReplace(str, "\r", "")
        Return str
    }

    ;カレントディレクトリ取得
    get_current_dir() {
        explorerHwnd := WinActive("ahk_class CabinetWClass")
        If (explorerHwnd) {
            for window in ComObjCreate("Shell.Application").Windows {
                If (window.hwnd==explorerHwnd)
                    Return window.Document.Folder.Self.Path
            }
        }
    }
