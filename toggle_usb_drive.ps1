# =============================================================================
# USB�h���C�u�̊Ǘ��ƈ��S�Ȏ��O���x��
#
# �ړI:
# 1. ���[�U�[���ݒ肵��USB�h���C�u�̐ڑ�(�L����)�Ɛؒf(������)���g�O�����삷��B
# 2. �l�b�g���[�N�o�R�̋��L�Z�b�V�����⃍�[�J���ŊJ����Ă���t�@�C���n���h����������A
#    USB�h���C�u�����S�Ɏ��O�����Ԃɂ��邱�Ƃ��x������B
#
# =============================================================================

# --- ���O���� ----------------------------------------------------------------
# 1. �Ǘ��Ҍ���:
#    ���̃X�N���v�g�͊Ǘ��Ҍ����Ŏ��s����K�v������܂��B
#    �X�N���v�g�t�@�C�����E�N���b�N���A�u�Ǘ��҂Ƃ��Ď��s�v��I�����Ă��������B
#
# 2. Handle �c�[���̔z�u:
#    Microsoft Sysinternals �� "Handle" �c�[�����K�v�ł��B
#    - �_�E�����[�h: https://learn.microsoft.com/ja-jp/sysinternals/downloads/handle
#    - ����: �_�E�����[�h�����t�@�C������ `handle64.exe` ���A
#            ���̃X�N���v�g�t�@�C���Ɓu�����t�H���_�v�ɔz�u���Ă��������B
# -----------------------------------------------------------------------------


# ��������������������������������������������������������������������������������������������������������������������������������������
# ���[�U�[�ݒ�: ���Ȃ���USB�f�o�C�X�ɍ��킹�Ĉȉ���3�̒l��ύX���Ă�������
# ��������������������������������������������������������������������������������������������������������������������������������������

# --- �ݒ���@ ---
# 1. �u�f�o�C�X �}�l�[�W���[�v���J���܂��B
# 2. �Y����USB�h���C�u���E�N���b�N���A�u�v���p�e�B�v��I�����܂��B
# 3. �u�ڍׁv�^�u�Ɉړ����܂��B
# 4. �u�v���p�e�B�v�̃h���b�v�_�E�����X�g����ȉ��̍��ڂ�I�����A�u�l�v���R�s�[���܂��B
#
#    - �u�n�[�h�E�F�A ID�v��I��:
#      ��: USB\VID_067B&PID_2775...
#      ���̏ꍇ�AVID�� "067B"�APID�� "2775" �ł��B
#      $targetVid �� $targetPid �ɂ��̒l��ݒ肵�܂��B
#
#    - �u�f�o�C�X�̐����v�܂��́u�t�����h�����v��I��:
#      ��: Prolific RAID0 USB Device
#      ���̒l�� $targetModel �ɐݒ肵�܂��B(����̓f�B�X�N�h���C�u�̃��f�����Ƃ��ĔF������܂�)
#      ���������܂��F������Ȃ��ꍇ�́AGet-WmiObject -Class Win32_DiskDrive �ŕ\�������
#        �Ώۃh���C�u�́uModel�v�v���p�e�B�̒l��ݒ肵�Ă��������B

$targetVid = "067B"
$targetPid = "2775"
$targetModel = "Prolific RAID0 USB Device"

# ��������������������������������������������������������������������������������������������������������������������������������������
# ������������������������������������������������������ �ݒ�͂����܂� ������������������������������������������������������


# --- �����`�F�b�N --------------------------------------------------------------

# 1. �Ǘ��Ҍ����`�F�b�N
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Warning "���̃X�N���v�g�͊Ǘ��Ҍ������K�v�ł��B�X�N���v�g���E�N���b�N���A�u�Ǘ��҂Ƃ��Ď��s�v���Ă��������B"
    Read-Host "Enter�L�[�������ďI�����܂�"
    exit 1
}

# 2. Handle.exe �̑��݃`�F�b�N
$handleExePath = Join-Path $PSScriptRoot "handle64.exe"
if (-not (Test-Path -Path $handleExePath -PathType Leaf)) {
    Write-Error "handle64.exe ��������܂���B"
    Write-Warning "���̃X�N���v�g�̑S�@�\�𗘗p����ɂ́AMicrosoft Sysinternals �� Handle �c�[�����K�v�ł��B"
    Write-Warning "1. https://learn.microsoft.com/ja-jp/sysinternals/downloads/handle �ɃA�N�Z�X���ă_�E�����[�h���Ă��������B"
    Write-Warning "2. �_�E�����[�h�����t�@�C������ handle64.exe ���A���̃X�N���v�g�Ɓw�����t�H���_�x�ɔz�u���Ă��������B"
    Read-Host "Enter�L�[�������ďI�����܂�"
    exit 1
}


# =============================================================================
# �֐���`
# =============================================================================

# --- �Ώ�USB�f�o�C�X�̓���֐� ---
function Find-TargetUsbDevice {
    try {
        $usbDevices = Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            $_.InstanceId -like "*VID_$($targetVid)&PID_$($targetPid)*" -and
            $_.Class -eq "USB"
        } | Sort-Object Status
        if ($usbDevices.Count -eq 0) { return $null }
        $activeDevice = $usbDevices | Where-Object { $_.Status -eq "OK" } | Select-Object -First 1
        if (-not $activeDevice) { $activeDevice = $usbDevices | Where-Object { $_.Status -eq "Error" } | Select-Object -First 1 }
        if (-not $activeDevice) { $activeDevice = $usbDevices | Select-Object -First 1 }
        $targetDisk = $null
        if ($activeDevice -and $activeDevice.Status -eq "OK") {
            $targetDisk = Get-WmiObject -Class Win32_DiskDrive -ErrorAction SilentlyContinue | Where-Object { $_.Model -eq $targetModel } | Select-Object -First 1
        }
        return @{ DiskDrive = $targetDisk; USBDevice = $activeDevice; InstanceId = $activeDevice.InstanceId; AllDevices = $usbDevices }
    }
    catch { Write-Warning "�f�o�C�X�������ɃG���[: $($_.Exception.Message)"; return $null }
}

# --- �Ώۃf�o�C�X�̃{�����[���擾�֐� ---
function Get-TargetDeviceVolumes {
    param($DeviceInfo)
    try {
        if (-not $DeviceInfo -or -not $DeviceInfo.DiskDrive) { return @() }
        $disk = $DeviceInfo.DiskDrive
        $diskNumber = ($disk.DeviceID -replace '\\\\.\\PHYSICALDRIVE', '')
        if (-not ($diskNumber -match '^\d+$')) { return @() }
        return Get-Partition -DiskNumber $diskNumber -ErrorAction SilentlyContinue | Get-Volume -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter }
    }
    catch { return @() }
}

# --- �f�o�C�X�g�p�󋵃`�F�b�N�֐� ---
function Test-DeviceInUse {
    param($DeviceInfo)
    try {
        $targetVolumes = Get-TargetDeviceVolumes -DeviceInfo $DeviceInfo
        if (-not $targetVolumes) { return $false }
        foreach ($volume in $targetVolumes) {
            try {
                $drivePath = Join-Path -Path ($volume.DriveLetter + ":") -ChildPath '\'
                $testFile = Join-Path $drivePath ([System.IO.Path]::GetRandomFileName())
                [System.IO.File]::Create($testFile).Close(); Remove-Item $testFile -Force -ErrorAction Stop
            } catch {
                Write-Host "�{�����[�� $($volume.DriveLetter): �̃��b�N�����o�i�g�p���̉\���j" -ForegroundColor Red
                return $true
            }
        }
        return $false
    } catch { Write-Warning "�g�p�󋵃`�F�b�N���ɃG���[: $($_.Exception.Message)"; return $true }
}

# --- ���S�Ȏ��O���֐� ---
function Safe-RemoveDevice {
    param($DeviceInfo)
    try {
        Write-Host "`n--- ���S�Ȏ��O���`�F�b�N���J�n���܂� ---" -ForegroundColor Cyan
        $volumes = Get-TargetDeviceVolumes -DeviceInfo $DeviceInfo
        if ($volumes) {
            foreach ($volume in $volumes) {
                $driveLetter = $volume.DriveLetter; $targetPath = "${driveLetter}:\"
                $openFilesOnDrive = Get-SmbOpenFile | Where-Object { $_.Path -like "$targetPath*" }
                if ($openFilesOnDrive) { Write-Error "���O���s��: �h���C�u $targetPath �ŃA�N�e�B�u�ȋ��L�Z�b�V���������o����܂����B"; Write-Warning "���j���[ [2] �܂��� [4] ���g�p���ăZ�b�V������������Ă��������B"; return $false }
                $handleOutput = & $handleExePath -nobanner $targetPath
                $localOpenFiles = $handleOutput | Where-Object { $_ -match '^\s*([\w.-]+)\s+pid:\s*(\d+)\s+.*:\s+([A-Z]:\\.*)$' }
                if ($localOpenFiles) { Write-Error "���O���s��: �h���C�u $targetPath �Ń��[�J���ŊJ����Ă���t�@�C�������o����܂����B"; Write-Warning "���j���[ [3] �܂��� [4] ���g�p���ăt�@�C���n���h�����m�F�E������Ă��������B"; return $false }
            }
        }
        Write-Host "[OK] �S�Ẵp�[�e�B�V�����ŋ��L/���[�J���̃��b�N�͂���܂���B" -ForegroundColor Green
        if (Test-DeviceInUse -DeviceInfo $DeviceInfo) { Write-Error "`n>> ���O���s��: �f�o�C�X���g�p���̉\��������܂��i�{�����[�����b�N�Ȃǁj�B"; Write-Warning ">> �J���Ă���A�v���P�[�V������G�N�X�v���[���[����Ă���Ď��s���Ă��������B"; return $false }
        Write-Host "[OK] �{�����[�����b�N�͌��o����܂���ł����B" -ForegroundColor Green
        Write-Host "`n���ׂẴ`�F�b�N���N���A���܂����B�f�o�C�X�̖������������J�n���܂��B" -ForegroundColor Green
        Disable-PnpDevice -InstanceId $DeviceInfo.InstanceId -Confirm:$false -ErrorAction Stop; Start-Sleep -Seconds 2
        $deviceAfter = Get-PnpDevice -InstanceId $DeviceInfo.InstanceId -ErrorAction SilentlyContinue
        if ($deviceAfter -and $deviceAfter.Status -eq 'Error') { Write-Host "�f�o�C�X�͐���ɖ���������܂��� (Status: Error - �R�[�h22)" -ForegroundColor Green; return $true }
        Write-Warning "�f�o�C�X�̖������Ɏ��s�����\��������܂��B"; return $false
    } catch { Write-Error "���S�Ȏ��O���������ɃG���[: $($_.Exception.Message)"; return $false }
}

# --- ���S�Ȑڑ��֐� ---
function Safe-ConnectDevice {
    param($DeviceInfo)
    try {
        Write-Host "USB�f�o�C�X��L������..." -ForegroundColor Cyan
        Enable-PnpDevice -InstanceId $DeviceInfo.InstanceId -Confirm:$false -ErrorAction Stop
        Write-Host "�f�o�C�X�L�����R�}���h�����s�B�{�����[���̏�����ҋ@��..." -ForegroundColor Cyan
        $timeout = 20; $progressInterval = 4
        foreach ($i in 1..$timeout) {
            Start-Sleep -Seconds 1
            $updatedDeviceInfo = Find-TargetUsbDevice
            if ($updatedDeviceInfo.USBDevice.Status -eq 'OK' -and $updatedDeviceInfo.DiskDrive) {
                if (Get-TargetDeviceVolumes -DeviceInfo $updatedDeviceInfo) { Write-Host "`n[OK] �f�o�C�X�̔F���ƃ}�E���g���������܂����I" -ForegroundColor Green; return $true }
            }
            if ($i % $progressInterval -eq 0) { Write-Host "�F���ҋ@��... ($i �b)" -ForegroundColor Cyan }
        }
        Write-Warning "�f�o�C�X�̊��S�ȔF�����^�C���A�E�g���܂���"; return $false
    } catch { Write-Error "���S�Ȑڑ��������ɃG���[: $($_.Exception.Message)"; return $false }
}

# --- �f�o�C�X��Ԕ���֐� ---
function Get-DeviceActualStatus {
    param($DeviceInfo)
    if ($DeviceInfo.USBDevice.Status -eq 'OK') {
        if ($DeviceInfo.DiskDrive) { if (Get-TargetDeviceVolumes -DeviceInfo $DeviceInfo) { return "Active" } else { return "Inactive" } } else { return "Connecting" }
    } elseif ($DeviceInfo.USBDevice.Status -eq 'Error') { return "Disabled" }
    else { return "Other" }
}

# --- ���L�Z�b�V�����֘A ---
function Manage-SharedSessions {
    Write-Host "`n--- ���L�Z�b�V�����̊m�F�E��� ---" -ForegroundColor Yellow
    $targetDriveLetters = @()
    $deviceInfo = Find-TargetUsbDevice
    if ($deviceInfo -and $deviceInfo.DiskDrive) {
        $volumes = Get-TargetDeviceVolumes -DeviceInfo $deviceInfo
        if ($volumes) {
            $autoDrives = $volumes.DriveLetter -join ', '
            Write-Host "�Ώۃf�o�C�X�̃h���C�u '$($autoDrives):' ���������o���܂����B" -ForegroundColor Cyan
            Write-Host "�����̃h���C�u���`�F�b�N���܂����H �ʂ̃h���C�u���w�肷��ꍇ�� 'n' ����͂��Ă������� [Y/n]:" -NoNewline
            if ((Read-Host) -ne 'n') { $targetDriveLetters = $volumes.DriveLetter }
        }
    }
    if ($targetDriveLetters.Count -eq 0) {
        Write-Host "�`�F�b�N�������h���C�u���^�[ (��: D) ����͂��Ă�������:" -ForegroundColor Cyan -NoNewline
        $manualDrive = Read-Host
        if (-not ($manualDrive -match '^[a-zA-Z]$')) { Write-Host "�����ȃh���C�u���^�[�ł��B" -ForegroundColor Red; return }
        $targetDriveLetters += $manualDrive
    }
    $allOpenFiles = @()
    Write-Host "`n�ȉ��̃h���C�u���`�F�b�N���܂�:" -ForegroundColor Cyan
    foreach ($driveLetter in $targetDriveLetters) {
        $targetPath = "$($driveLetter):\"
        Write-Host " - �h���C�u: $targetPath"
        $allOpenFiles += Get-SmbOpenFile | Where-Object { $_.Path -like "$targetPath*" }
    }
    if (-not $allOpenFiles) { Write-Host "`n�Ώۃh���C�u�Ɋ֘A����A�N�e�B�u�ȋ��L�Z�b�V�����͂���܂���B" -ForegroundColor Green; return }
    Write-Host "`n�Ώۃh���C�u��̃t�@�C���ɃA�N�Z�X���Ă���Z�b�V������������܂����F" -ForegroundColor Red
    $relevantSessionIds = $allOpenFiles | Select-Object -ExpandProperty SessionId -Unique
    $relevantSessions = Get-SmbSession | Where-Object { $_.SessionId -in $relevantSessionIds }
    $relevantSessions | ForEach-Object -Begin { $i = 1 } -Process { Write-Host "[$i] �N���C�A���g: $($_.ClientComputerName), ���[�U�[: $($_.ClientUserName)"; $i++ }
    Write-Host "`n�����̋��L�Z�b�V���������ׂĕ��܂����H [y/n]:" -ForegroundColor Yellow -NoNewline
    if ((Read-Host) -eq 'y') {
        $relevantSessions | ForEach-Object { try { Close-SmbSession -SessionId $_.SessionId -Force; Write-Host "�Z�b�V���� $($_.SessionId) ����܂����B" -ForegroundColor Green } catch { Write-Host "�Z�b�V���� $($_.SessionId) �̐ؒf�Ɏ��s: $($_.Exception.Message)" -ForegroundColor Red } }
    } else { Write-Host "�Z�b�V�����̐ؒf���L�����Z�����܂����B" -ForegroundColor Gray }
}

# --- ���[�J���t�@�C���n���h���֘A ---
function Check-LocalFileHandles {
    $targetDriveLetters = @()
    $deviceInfo = Find-TargetUsbDevice
    if ($deviceInfo -and $deviceInfo.DiskDrive) {
        $volumes = Get-TargetDeviceVolumes -DeviceInfo $deviceInfo
        if ($volumes) {
            $autoDrives = $volumes.DriveLetter -join ', '
            Write-Host "�Ώۃf�o�C�X�̃h���C�u '$($autoDrives):' ���������o���܂����B" -ForegroundColor Cyan
            Write-Host "�����̃h���C�u���`�F�b�N���܂����H �ʂ̃h���C�u���w�肷��ꍇ�� 'n' ����͂��Ă������� [Y/n]:" -NoNewline
            if ((Read-Host) -ne 'n') {
                $targetDriveLetters = $volumes.DriveLetter
            }
        }
    }

    if ($targetDriveLetters.Count -eq 0) {
        Write-Host "�`�F�b�N�������h���C�u���^�[ (��: D) ����͂��Ă�������:" -ForegroundColor Cyan -NoNewline
        $manualDrive = Read-Host
        if (-not ($manualDrive -match '^[a-zA-Z]$')) { Write-Host "�����ȃh���C�u���^�[�ł��B" -ForegroundColor Red; return }
        $targetDriveLetters += $manualDrive
    }

    Write-Host ""
    foreach ($driveLetter in $targetDriveLetters) {
        $targetDrive = "${driveLetter}:"
        Write-Host "--- �h���C�u $targetDrive �ŊJ����Ă���t�@�C���̈ꗗ ---" -ForegroundColor Yellow
        $handleOutput = & $handleExePath -nobanner $targetDrive
        $openFiles = $handleOutput | ForEach-Object { if ($_ -match '^\s*([\w.-]+)\s+pid:\s*(\d+)\s+.*:\s+([A-Z]:\\.*)$') { [PSCustomObject]@{ ProcessName = $matches[1]; PID = $matches[2]; FilePath = $matches[3].Trim() } } }

        if ($openFiles.Count -eq 0) {
            Write-Host "�h���C�u $targetDrive ���g�p���Ă���v���Z�X�͌�����܂���ł����B" -ForegroundColor Green
        } else {
            Write-Host "�ȉ��̃v���Z�X���h���C�u $targetDrive ��̃t�@�C�����J���Ă��܂��B" -ForegroundColor Red
            $openFiles | Format-Table -AutoSize
            Write-Warning "�����̃t�@�C���ւ̃A�N�Z�X�����ɂ́A��L���X�g�̃v���Z�X�i�A�v���P�[�V�����j���蓮�ŏI�����Ă��������B"
        }
        Write-Host ""
    }
}


# =============================================================================
# ���C������
# =============================================================================
function Show-Menu {
    Write-Host "=== USB�h���C�u �����Ǘ��c�[�� ===" -ForegroundColor Green
    Write-Host "-------------------------------------------"
    $deviceInfo = Find-TargetUsbDevice
    if ($deviceInfo) {
        $status = Get-DeviceActualStatus -DeviceInfo $deviceInfo
        Write-Host "�Ώۃf�o�C�X: $($deviceInfo.USBDevice.FriendlyName) ($($targetVid):$($targetPid))" -ForegroundColor Cyan
        Write-Host "�f�o�C�X���: $status" -ForegroundColor Cyan
    } else {
        Write-Host "�Ώۃf�o�C�X: ($($targetVid):$($targetPid)) �͌�����܂���" -ForegroundColor Yellow
    }
    Write-Host "-------------------------------------------"
    Write-Host "�I�����Ă�������:" -ForegroundColor Cyan
    Write-Host "1. �f�o�C�X�̐ڑ� / �ؒf (�g�O������)"
    Write-Host "2. �y���O�������z���L�Z�b�V�����̉��"
    Write-Host "3. �y���O�������z���[�J���t�@�C���n���h���̊m�F"
    Write-Host "4. �y��������z������������s���Ĉ��S�Ɏ��O��"
    Write-Host "5. �I��"
    Write-Host ""
}

while ($true) {
    Show-Menu
    $choice = Read-Host "�I��"
    switch ($choice) {
        "1" {
            $deviceInfo = Find-TargetUsbDevice
            if (-not $deviceInfo) { Write-Error "�Ώۂ�USB�f�o�C�X��������܂���B�ݒ���m�F���Ă��������B"; continue }
            $actualStatus = Get-DeviceActualStatus -DeviceInfo $deviceInfo
            Write-Host "���݂̏�Ԃ� '$actualStatus' �ł��B" -ForegroundColor Cyan
            switch ($actualStatus) {
                "Active"   { Write-Host "�ؒf���������s���܂�..."; Safe-RemoveDevice -DeviceInfo $deviceInfo }
                "Disabled" { Write-Host "�ڑ����������s���܂�..."; Safe-ConnectDevice -DeviceInfo $deviceInfo }
                default    { Write-Warning "�f�o�C�X�����ԓI�ȏ�� ($actualStatus) �ł��B���Z�b�g�����݂܂�..."; Safe-RemoveDevice -DeviceInfo $deviceInfo }
            }
        }
        "2" { Manage-SharedSessions }
        "3" { Check-LocalFileHandles }
        "4" {
            Write-Host "--- �y��������z��������ƈ��S�Ȏ��O�����J�n���܂� ---" -ForegroundColor Magenta
            $deviceInfo = Find-TargetUsbDevice
            if (-not $deviceInfo -or (Get-DeviceActualStatus -DeviceInfo $deviceInfo) -ne "Active") { Write-Warning "�f�o�C�X�� 'Active' ��Ԃł͂���܂���B���̑���͎��s�ł��܂���B"; continue }
            Manage-SharedSessions
            $volumes = Get-TargetDeviceVolumes -DeviceInfo $deviceInfo
            if ($volumes) { foreach ($volume in $volumes) { Check-LocalFileHandles -driveLetter $volume.DriveLetter } } # ���̕����͎蓮�m�F�����܂Ȃ�
            Write-Host "`n����������������܂����B�f�o�C�X�̈��S�Ȏ��O�������s���܂����H [y/n]:" -ForegroundColor Yellow -NoNewline
            if ((Read-Host) -eq 'y') { Safe-RemoveDevice -DeviceInfo $deviceInfo } else { Write-Host "���O�����L�����Z�����܂����B" -ForegroundColor Gray }
        }
        "5" { Write-Host "�I�����܂��B"; exit 0 }
        default { Write-Host "�����ȑI���ł��B1-5����͂��Ă��������B" -ForegroundColor Red }
    }
    Write-Host "`n���������BEnter�L�[�������ă��j���[�ɖ߂�܂�..." -ForegroundColor Gray
    Read-Host
    Clear-Host
}