# �ގЎ���PC�I�����@
# LOCK:     ��ʃ��b�N
# LOGOFF:   ���O�I�t
# SHUTDOWN: PC�d���I�t
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
    $result = [System.Windows.Forms.MessageBox]::Show($message, "�ގЂ��܂�", "YesNo", "Question", "button1")
    return $result -eq "Yes"
}

function ClockOut() {
    Start-Process -FilePath $PYTHON_SCRIPT_PATH -ArgumentList "OUT" -Wait
}

switch ($LEAVING_METHOD) {
    "LOCK" {
        if (AskClockOut("��ʃ��b�N����O�ɑō����܂���?")) { ClockOut }
        LockPC
    }
    "LOGOFF" {
        if (AskClockOut("���O�I�t�O�ɑō����܂���?")) { ClockOut }
        LockPC
    }
    "SHUTDOWN" {
        if (AskClockOut("�V���b�g�_�E���O�ɑō����܂���?")) { ClockOut }
        LockPC
    }
}