# =============================================================================
# USBドライブの管理と安全な取り外し支援
#
# 目的:
# 1. ユーザーが設定したUSBドライブの接続(有効化)と切断(無効化)をトグル操作する。
# 2. ネットワーク経由の共有セッションやローカルで開かれているファイルハンドルを解放し、
#    USBドライブを安全に取り外せる状態にすることを支援する。
#
# =============================================================================

# --- 事前準備 ----------------------------------------------------------------
# 1. 管理者権限:
#    このスクリプトは管理者権限で実行する必要があります。
#    スクリプトファイルを右クリックし、「管理者として実行」を選択してください。
#
# 2. Handle ツールの配置:
#    Microsoft Sysinternals の "Handle" ツールが必要です。
#    - ダウンロード: https://learn.microsoft.com/ja-jp/sysinternals/downloads/handle
#    - 準備: ダウンロードしたファイル内の `handle64.exe` を、
#            このスクリプトファイルと「同じフォルダ」に配置してください。
# -----------------------------------------------------------------------------


# ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
# ユーザー設定: あなたのUSBデバイスに合わせて以下の3つの値を変更してください
# ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

# --- 設定方法 ---
# 1. 「デバイス マネージャー」を開きます。
# 2. 該当のUSBドライブを右クリックし、「プロパティ」を選択します。
# 3. 「詳細」タブに移動します。
# 4. 「プロパティ」のドロップダウンリストから以下の項目を選択し、「値」をコピーします。
#
#    - 「ハードウェア ID」を選択:
#      例: USB\VID_067B&PID_2775...
#      この場合、VIDは "067B"、PIDは "2775" です。
#      $targetVid と $targetPid にその値を設定します。
#
#    - 「デバイスの説明」または「フレンドリ名」を選択:
#      例: Prolific RAID0 USB Device
#      この値を $targetModel に設定します。(これはディスクドライブのモデル名として認識されます)
#      ※もしうまく認識されない場合は、Get-WmiObject -Class Win32_DiskDrive で表示される
#        対象ドライブの「Model」プロパティの値を設定してください。

$targetVid = "067B"
$targetPid = "2775"
$targetModel = "Prolific RAID0 USB Device"

# ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
# ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ 設定はここまで ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲


# --- 初期チェック --------------------------------------------------------------

# 1. 管理者権限チェック
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Warning "このスクリプトは管理者権限が必要です。スクリプトを右クリックし、「管理者として実行」してください。"
    Read-Host "Enterキーを押して終了します"
    exit 1
}

# 2. Handle.exe の存在チェック
$handleExePath = Join-Path $PSScriptRoot "handle64.exe"
if (-not (Test-Path -Path $handleExePath -PathType Leaf)) {
    Write-Error "handle64.exe が見つかりません。"
    Write-Warning "このスクリプトの全機能を利用するには、Microsoft Sysinternals の Handle ツールが必要です。"
    Write-Warning "1. https://learn.microsoft.com/ja-jp/sysinternals/downloads/handle にアクセスしてダウンロードしてください。"
    Write-Warning "2. ダウンロードしたファイル内の handle64.exe を、このスクリプトと『同じフォルダ』に配置してください。"
    Read-Host "Enterキーを押して終了します"
    exit 1
}


# =============================================================================
# 関数定義
# =============================================================================

# --- 対象USBデバイスの特定関数 ---
function Find-TargetUsbDevice {
    try {
        # 設定されたVIDとPIDでデバイスを検索
        $usbDevices = Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            $_.InstanceId -like "*VID_$($targetVid)&PID_$($targetPid)*" -and
            $_.Class -eq "USB"
        } | Sort-Object Status

        if ($usbDevices.Count -eq 0) { return $null }

        $activeDevice = $usbDevices | Where-Object { $_.Status -eq "OK" } | Select-Object -First 1
        if (-not $activeDevice) {
            $activeDevice = $usbDevices | Where-Object { $_.Status -eq "Error" } | Select-Object -First 1
        }
        if (-not $activeDevice) {
            $activeDevice = $usbDevices | Select-Object -First 1
        }

        $targetDisk = $null
        if ($activeDevice -and $activeDevice.Status -eq "OK") {
            # 設定されたモデル名でディスクドライブを検索
            $targetDisk = Get-WmiObject -Class Win32_DiskDrive -ErrorAction SilentlyContinue |
                           Where-Object { $_.Model -eq $targetModel } |
                           Select-Object -First 1
        }

        return @{
            DiskDrive = $targetDisk
            USBDevice = $activeDevice
            InstanceId = $activeDevice.InstanceId
            AllDevices = $usbDevices
        }
    }
    catch {
        Write-Warning "デバイス検索中にエラー: $($_.Exception.Message)"
        return $null
    }
}

# --- 対象デバイスのボリューム取得関数 ---
function Get-TargetDeviceVolumes {
    param($DeviceInfo)
    try {
        if (-not $DeviceInfo -or -not $DeviceInfo.DiskDrive) { return @() }

        $disk = $DeviceInfo.DiskDrive
        $diskNumber = ($disk.DeviceID -replace '\\\\.\\PHYSICALDRIVE', '')

        if (-not ($diskNumber -match '^\d+$')) { return @() }

        $targetVolumes = @()
        $partitions = Get-Partition -DiskNumber $diskNumber -ErrorAction SilentlyContinue
        if ($partitions) {
            foreach ($partition in $partitions) {
                $volume = Get-Volume -Partition $partition -ErrorAction SilentlyContinue
                if ($volume -and $volume.DriveLetter) {
                    $driveLetter = $volume.DriveLetter + ":"
                    $logicalDisk = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID = '$driveLetter'" -ErrorAction SilentlyContinue
                    if ($logicalDisk) {
                        $targetVolumes += $logicalDisk
                    }
                }
            }
        }
        return @($targetVolumes | Select-Object -Unique)
    }
    catch {
        return @()
    }
}

# --- デバイス使用状況チェック関数 ---
function Test-DeviceInUse {
    param($DeviceInfo)
    try {
        Write-Host "デバイス使用状況チェックを実行中..." -ForegroundColor Cyan
        $targetVolumes = Get-TargetDeviceVolumes -DeviceInfo $DeviceInfo
        if (-not $targetVolumes) {
            Write-Host "対象ボリュームが見つからないため、使用中ではないと判定" -ForegroundColor Green
            return $false
        }

        $inUse = $false
        foreach ($volume in $targetVolumes) {
            Write-Host "ボリューム $($volume.DeviceID) の使用状況をチェック中..." -ForegroundColor Cyan
            try {
                $testFile = Join-Path $volume.DeviceID ([System.IO.Path]::GetRandomFileName())
                [System.IO.File]::Create($testFile).Close(); Remove-Item $testFile -Force -ErrorAction Stop
                Write-Host "ボリューム $($volume.DeviceID) は正常にアクセス可能" -ForegroundColor Green
            } catch {
                Write-Host "ボリューム $($volume.DeviceID) のロックを検出（使用中の可能性）" -ForegroundColor Red
                $inUse = $true; break
            }
        }
        return $inUse
    } catch {
        Write-Warning "使用状況チェック中にエラー: $($_.Exception.Message)"; return $true
    }
}

# --- 安全な取り外し関数 ---
function Safe-RemoveDevice {
    param($DeviceInfo)
    try {
        Write-Host "`n--- 安全な取り外しチェックを開始します ---" -ForegroundColor Cyan

        # チェック1: 共有セッションの存在確認
        $sessions = Get-SmbSession | Where-Object { $_.ClientComputerName -notin @($env:COMPUTERNAME, "127.0.0.1", "::1") }
        if ($sessions) {
            Write-Error "取り外し不可: アクティブな共有セッションが検出されました。"
            Write-Warning "メニュー [2] または [4] を使用してセッションを解放してください。"
            $sessions | ForEach-Object { Write-Host "  - クライアント: $($_.ClientComputerName), ユーザー: $($_.ClientUserName)" }
            return $false
        }
        Write-Host "[OK] 共有セッションはありません。" -ForegroundColor Green

        # チェック2: ローカルファイルハンドルの存在確認
        $volumes = Get-TargetDeviceVolumes -DeviceInfo $DeviceInfo
        if ($volumes -and $volumes.Count -gt 0) {
            # DeviceIDの存在と値をチェック
            $firstVolume = $volumes[0]
            if ($firstVolume -and $firstVolume.DeviceID -and $firstVolume.DeviceID.ToString().Length -gt 0) {
                $driveLetter = $firstVolume.DeviceID.ToString().Trim(":")
                $targetDrive = "${driveLetter}:"
                
                Write-Host "ドライブ $targetDrive のファイルハンドルをチェック中..." -ForegroundColor Cyan
                
                try {
                    $handleOutput = & $handleExePath -nobanner $targetDrive 2>$null
                    if ($LASTEXITCODE -eq 0 -and $handleOutput) {
                        $openFiles = $handleOutput | Where-Object { $_ -match '^\s*([\w.-]+)\s+pid:\s*(\d+)\s+.*:\s+([A-Z]:\\.*)$' }
                        if ($openFiles -and $openFiles.Count -gt 0) {
                            Write-Error "取り外し不可: ローカルで開かれているファイルが検出されました。"
                            Write-Warning "メニュー [3] または [4] を使用してファイルハンドルを確認・解放してください。"
                            return $false
                        }
                    }
                } catch {
                    Write-Warning "Handle.exe の実行中にエラーが発生しました: $($_.Exception.Message)"
                }
            } else {
                Write-Warning "ボリュームのDeviceIDが取得できませんでした。ファイルハンドルチェックをスキップします。"
            }
        } else {
            Write-Host "対象ボリュームが見つからないため、ファイルハンドルチェックをスキップします。" -ForegroundColor Yellow
        }
        Write-Host "[OK] ローカルで開かれているファイルはありません。" -ForegroundColor Green

        # チェック3: 従来の簡易チェック (念のため)
        if (Test-DeviceInUse -DeviceInfo $DeviceInfo) {
            Write-Error "`n>> 取り外し不可: デバイスが使用中の可能性があります（ボリュームロックなど）。"
            Write-Warning ">> 開いているアプリケーションやエクスプローラーを閉じてから再試行してください。"
            return $false
        }
        Write-Host "[OK] ボリュームロックは検出されませんでした。" -ForegroundColor Green

        Write-Host "`nすべてのチェックをクリアしました。デバイスの無効化処理を開始します。" -ForegroundColor Green
        Disable-PnpDevice -InstanceId $DeviceInfo.InstanceId -Confirm:$false -ErrorAction Stop
        Start-Sleep -Seconds 2
        $deviceAfter = Get-PnpDevice -InstanceId $DeviceInfo.InstanceId -ErrorAction SilentlyContinue
        if ($deviceAfter -and $deviceAfter.Status -eq 'Error') {
            Write-Host "デバイスは正常に無効化されました (Status: Error - コード22)" -ForegroundColor Green
            return $true
        }
        Write-Warning "デバイスの無効化に失敗した可能性があります。"; return $false
    } catch {
        Write-Error "安全な取り外し処理中にエラー: $($_.Exception.Message)"; return $false
    }
}

# --- 安全な接続関数 ---
function Safe-ConnectDevice {
    param($DeviceInfo)
    try {
        Write-Host "USBデバイスを有効化中..." -ForegroundColor Cyan
        Enable-PnpDevice -InstanceId $DeviceInfo.InstanceId -Confirm:$false -ErrorAction Stop
        Write-Host "デバイス有効化コマンドを実行。ボリュームの準備を待機中..." -ForegroundColor Cyan

        $timeout = 20
        $progressInterval = 4

        foreach ($i in 1..$timeout) {
            Start-Sleep -Seconds 1
            $updatedDeviceInfo = Find-TargetUsbDevice # ★ 変更
            if ($updatedDeviceInfo.USBDevice.Status -eq 'OK' -and $updatedDeviceInfo.DiskDrive) {
                $volumes = Get-TargetDeviceVolumes -DeviceInfo $updatedDeviceInfo
                if ($volumes) {
                    Write-Host "`n[OK] デバイスの認識とマウントが完了しました！" -ForegroundColor Green
                    return $true
                }
            }
            if ($i % $progressInterval -eq 0) { Write-Host "認識待機中... ($i 秒)" -ForegroundColor Cyan }
        }

        Write-Warning "デバイスの完全な認識がタイムアウトしました"; return $false
    } catch {
        Write-Error "安全な接続処理中にエラー: $($_.Exception.Message)"; return $false
    }
}

# --- デバイス状態判定関数 ---
function Get-DeviceActualStatus {
    param($DeviceInfo)
    if ($DeviceInfo.USBDevice.Status -eq 'OK') {
        if ($DeviceInfo.DiskDrive) {
            $volumes = Get-TargetDeviceVolumes -DeviceInfo $DeviceInfo
            if ($volumes) { return "Active" }
            else { return "Inactive" }
        } else { return "Connecting" }
    } elseif ($DeviceInfo.USBDevice.Status -eq 'Error') { return "Disabled" }
    else { return "Other" }
}

# --- 共有セッション関連 ---
function Manage-SharedSessions {
    Write-Host "`n--- 共有セッション一覧 ---" -ForegroundColor Yellow
    $sessions = Get-SmbSession | Where-Object { $_.ClientComputerName -notin @($env:COMPUTERNAME, "127.0.0.1", "::1") }

    if ($sessions.Count -eq 0) {
        Write-Host "アクティブな共有セッションはありません。" -ForegroundColor Green
        return
    }

    $sessions | ForEach-Object -Begin { $i = 1 } -Process {
        Write-Host "[$i] クライアント: $($_.ClientComputerName), ユーザー: $($_.ClientUserName)"
        $i++
    }

    Write-Host ""
    Write-Host "これらの共有セッションをすべて閉じますか？ [y/n]:" -ForegroundColor Yellow -NoNewline
    $choice = Read-Host

    if ($choice -eq 'y') {
        $sessions | ForEach-Object {
            try {
                Close-SmbSession -SessionId $_.SessionId -Force
                Write-Host "セッション $($_.SessionId) を閉じました。" -ForegroundColor Green
            }
            catch {
                Write-Host "セッション $($_.SessionId) の切断に失敗: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "セッションの切断をキャンセルしました。" -ForegroundColor Gray
    }
}

# --- ローカルファイルハンドル関連 ---
function Check-LocalFileHandles {
    param([string]$DriveLetter)

    if (-not $DriveLetter) {
        $deviceInfo = Find-TargetUsbDevice # ★ 変更
        if ($deviceInfo -and $deviceInfo.DiskDrive) {
            $volumes = Get-TargetDeviceVolumes -DeviceInfo $deviceInfo
            if ($volumes) {
                $DriveLetter = $volumes[0].DeviceID.Trim(":")
                Write-Host "対象デバイスのドライブ '$($DriveLetter):' を自動検出しました。" -ForegroundColor Cyan
            }
        }
    }

    if (-not $DriveLetter) {
        Write-Host "ドライブレター (例: D) を入力してください:" -ForegroundColor Cyan -NoNewline
        $DriveLetter = Read-Host
    }

    if (-not ($DriveLetter -match '^[a-zA-Z]$')) {
        Write-Host "無効なドライブレターです。" -ForegroundColor Red
        return
    }

    $targetDrive = "${DriveLetter}:"
    Write-Host "`n--- ドライブ $targetDrive で開かれているファイルの一覧 ---" -ForegroundColor Yellow

    $handleOutput = & $handleExePath -nobanner $targetDrive

    $openFiles = $handleOutput | ForEach-Object {
        if ($_ -match '^\s*([\w.-]+)\s+pid:\s*(\d+)\s+.*:\s+([A-Z]:\\.*)$') {
            [PSCustomObject]@{
                ProcessName = $matches[1]
                PID         = $matches[2]
                FilePath    = $matches[3].Trim()
            }
        }
    }

    if ($openFiles.Count -eq 0) {
        Write-Host "ドライブ $targetDrive を使用しているプロセスは見つかりませんでした。" -ForegroundColor Green
    } else {
        Write-Host "以下のプロセスがドライブ $targetDrive 上のファイルを開いています。" -ForegroundColor Red
        $openFiles | Format-Table -AutoSize
        Write-Warning "これらのファイルへのアクセスを閉じるには、上記リストのプロセス（アプリケーション）を手動で終了してください。"
        Write-Warning "例: 'explorer.exe' ならフォルダを閉じ、'msword.exe' ならWordを終了します。"
    }
}


# =============================================================================
# メイン処理
# =============================================================================
function Show-Menu {
    Write-Host "=== USBドライブ 総合管理ツール ===" -ForegroundColor Green # ★ 変更
    Write-Host "-------------------------------------------"
    $deviceInfo = Find-TargetUsbDevice # ★ 変更
    if ($deviceInfo) {
        $status = Get-DeviceActualStatus -DeviceInfo $deviceInfo
        Write-Host "対象デバイス: $($deviceInfo.USBDevice.FriendlyName) ($($targetVid):$($targetPid))" -ForegroundColor Cyan
        Write-Host "デバイス状態: $status" -ForegroundColor Cyan
    } else {
        Write-Host "対象デバイス: ($($targetVid):$($targetPid)) は見つかりません" -ForegroundColor Yellow
    }
    Write-Host "-------------------------------------------"

    Write-Host "選択してください:" -ForegroundColor Cyan
    Write-Host "1. デバイスの接続 / 切断 (トグル操作)"
    Write-Host "2. 【取り外し準備】共有セッションの解放"
    Write-Host "3. 【取り外し準備】ローカルファイルハンドルの確認"
    Write-Host "4. 【統合操作】解放処理を実行して安全に取り外す"
    Write-Host "5. 終了"
    Write-Host ""
}

while ($true) {
    Show-Menu
    $choice = Read-Host "選択"

    switch ($choice) {
        "1" {
            $deviceInfo = Find-TargetUsbDevice # ★ 変更
            if (-not $deviceInfo) { Write-Error "対象のUSBデバイスが見つかりません。設定を確認してください。"; continue } # ★ 変更

            $actualStatus = Get-DeviceActualStatus -DeviceInfo $deviceInfo
            Write-Host "現在の状態は '$actualStatus' です。" -ForegroundColor Cyan

            switch ($actualStatus) {
                "Active"   { Write-Host "切断処理を実行します..."; Safe-RemoveDevice -DeviceInfo $deviceInfo }
                "Disabled" { Write-Host "接続処理を実行します..."; Safe-ConnectDevice -DeviceInfo $deviceInfo }
                default    {
                    Write-Warning "デバイスが中間的な状態 ($actualStatus) です。リセットを試みます..."
                    Safe-RemoveDevice -DeviceInfo $deviceInfo
                }
            }
        }
        "2" {
            Manage-SharedSessions
        }
        "3" {
            Check-LocalFileHandles
        }
        "4" {
            Write-Host "--- 【統合操作】解放処理と安全な取り外しを開始します ---" -ForegroundColor Magenta
            
            $deviceInfo = Find-TargetUsbDevice # ★ 変更
            if (-not $deviceInfo -or (Get-DeviceActualStatus -DeviceInfo $deviceInfo) -ne "Active") {
                Write-Warning "デバイスが 'Active' 状態ではありません。この操作は実行できません。"
                continue
            }

            Manage-SharedSessions
            Check-LocalFileHandles

            Write-Host "`n解放処理が完了しました。デバイスの安全な取り外しを実行しますか？ [y/n]:" -ForegroundColor Yellow -NoNewline
            $confirm = Read-Host
            if ($confirm -eq 'y') {
                Safe-RemoveDevice -DeviceInfo $deviceInfo
            } else {
                Write-Host "取り外しをキャンセルしました。" -ForegroundColor Gray
            }
        }
        "5" {
            Write-Host "終了します。"
            exit 0
        }
        default {
            Write-Host "無効な選択です。1-5を入力してください。" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "処理完了。Enterキーを押してメニューに戻ります..." -ForegroundColor Gray
    Read-Host
    Clear-Host
}