# 退社時のPC終了方法
# LOCK:     画面ロック
# LOGOFF:   ログオフ
# SHUTDOWN: PC電源オフ
set LEAVING_METHOD "LOCK" -option constant

$PYTHON_SCRIPT_PATH = $PSScriptRoot + "\TimeRecorder.py"

function LockPC() {
    rundll32 user32.dll, LockWorkStation
}

function Logoff() {
    shutdown.exe /l
}

function Shutdown() {
    shutdown.exe /s
}

function AskClockOut($message) {
    Add-Type -Assembly System.Windows.Forms
    $result = [System.Windows.Forms.MessageBox]::Show($message, "退社します", "YesNo", "Question", "button1")
    return $result -eq "Yes"
}

function ClockOut() {
    Start-Process -FilePath $PYTHON_SCRIPT_PATH -ArgumentList "OUT" -Wait
}

switch ($LEAVING_METHOD) {
    "LOCK" {
        if (AskClockOut("画面ロックする前に打刻しますか?")) { ClockOut }
        LockPC
    }
    "LOGOFF" {
        if (AskClockOut("ログオフ前に打刻しますか?")) { ClockOut }
        LockPC
    }
    "SHUTDOWN" {
        if (AskClockOut("シャットダウン前に打刻しますか?")) { ClockOut }
        LockPC
    }
}