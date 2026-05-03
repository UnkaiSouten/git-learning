#SingleInstance Force

^0::
    ClipWait, 2
    if ErrorLevel
        return

    text := Clipboard

    ; ===== 空白行を削除 =====
    text := RegExReplace(text, "m)^\s*\r?\n", "")

    ; ===== ハイフンだけの行を削除 =====
    text := RegExReplace(text, "m)^\s*-\s*\r?\n", "")

    ; ===== No. + 数字 の前に罫線を追加 =====
    line := "--------------------------------------------------------------------------------------------------------------------------------"
    text := RegExReplace(text, "m)^No\.\s*\d+\s*$", line . "`n$0")

    ; ===== VSCode 用にテキスト保存 =====
    tempTxt := A_Temp . "\preview_temp.txt"
    FileDelete, %tempTxt%
    FileAppend, %text%, %tempTxt%

    ; ===== VSCode 起動 =====
    vscode := "C:\Users\" . A_UserName . "\AppData\Local\Programs\Microsoft VS Code\Code.exe"
    Run, "%vscode%" "%tempTxt%"

    ; ===== HTMLエスケープ =====
    esc := text
    esc := StrReplace(esc, "&", "&amp;")
    esc := StrReplace(esc, "<", "&lt;")
    esc := StrReplace(esc, ">", "&gt;")

    ; ===== HTML側で罫線を太字に変換 =====
    esc := StrReplace(esc, line, "<b>" . line . "</b>")

    ; ===== 本文中の日付・時刻（太字＋色） =====
    esc := RegExReplace(esc, "(\d{4}[/-]\d{1,2}[/-]\d{1,2})", "<span class=""date-inline"">$1</span>")
    esc := RegExReplace(esc, "\b(\d{1,2}[/-]\d{1,2})\b", "<span class=""date-inline"">$1</span>")
    esc := RegExReplace(esc, "(\b\d{1,2}:\d{2}(?::\d{2})?\b)", "<span class=""time-inline"">$1</span>")

    ; ===== 識別番号 =====
    esc := RegExReplace(esc, "\b0\d{1,3}-\d{2,4}-\d{3,4}\b", "<span class=""tel-inline"">$0</span>")
    esc := RegExReplace(esc, "\bT\d{6}\b", "<span class=""t-inline"">$0</span>")
    esc := RegExReplace(esc, "\bNo\.\d+\b", "<span class=""no-inline"">$0</span>")
    esc := RegExReplace(esc, "\bFM\d+\b", "<span class=""fm-inline"">$0</span>")
    esc := RegExReplace(esc, "\bWM\d+\b", "<span class=""wm-inline"">$0</span>")

    ; ===== 上部の日付・時刻 =====
    FormatTime, dateDisplay, , yyyy/MM/dd
    FormatTime, timeDisplay, , HH:mm:ss

    ; ===== HTMLテンプレート（継続セクション） =====
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
    background: #e8f5e9; /* ← 薄い緑 */
}
.container {
    background: #e8f5e9; /* ← 白から薄い緑に変更 */
    padding: 20px;
    border-radius: 8px;
    border: 1px solid #ccc;
}
.timestamp-wrapper { margin-bottom: 12px; }
.date-tag {
    display: inline-block;
    background: #e6f0ff;
    color: #004a99;
    padding: 4px 8px;
    border-radius: 4px;
    font-weight: bold;
    margin-right: 6px;
}
.time-tag {
    display: inline-block;
    background: #ffe9cc;
    color: #b85c00;
    padding: 4px 8px;
    border-radius: 4px;
    font-weight: bold;
}
.date-inline { font-weight: bold; color: #ff66aa; }
.time-inline { font-weight: bold; color: #0066ff; }
.tel-inline { background:#e0ffe0; color:#008000; padding:2px 4px; border-radius:3px; font-weight:bold; }
.t-inline   { background:#f0e0ff; color:#6600cc; padding:2px 4px; border-radius:3px; font-weight:bold; }
.no-inline  { background:#ffe0e0; color:#cc0000; padding:2px 4px; border-radius:3px; font-weight:bold; }
.fm-inline  { background:#fff0e0; color:#a65e00; padding:2px 4px; border-radius:3px; font-weight:bold; }
.wm-inline  { background:#ffe0f5; color:#cc0077; padding:2px 4px; border-radius:3px; font-weight:bold; }
.text-body {
    white-space: pre-wrap;
    word-wrap: break-word;
    font-size: 14px;
    line-height: 1.6;
}
</style>
</head>
<body>
<div class="container">

<div class="timestamp-wrapper">
    <span class="date-tag">%dateDisplay%</span>
    <span class="time-tag">%timeDisplay%</span>
</div>

<div class="text-body">
%esc%
</div>

</div>
</body>
</html>
)

    ; ===== HTML保存 =====
    tempHtml := A_Temp . "\preview_temp.html"
    FileDelete, %tempHtml%
    FileAppend, %html%, %tempHtml%

    Run, %tempHtml%

    ; ===== Sakura に貼り付け =====
    Clipboard := text
    Run, "C:\Program Files (x86)\sakura\sakura.exe"
    WinWaitActive, ahk_exe sakura.exe
    SendInput, ^v
return
