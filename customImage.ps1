# The character code of this file is ShiftJIS!!!
# 管理者権限で実行
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole("Administrators")) { Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs; exit }

######################################
# 関数のセット
######################################

# Yes/No ダイアログ
function Get-YesNoDialog {
    $title = $args.Title
    $message = $args.Message

    $tChoiceDescription = "System.Management.Automation.Host.ChoiceDescription"
    $options = @(
        New-Object $tChoiceDescription ("はい(&Yes)", $args.YesMsg)
        New-Object $tChoiceDescription ("いいえ(&No)", $args.NoMsg)
    )
    $result = $host.ui.PromptForChoice($title, $message, $options, 0)
    switch ($result) {
        0 { return $true }        
        1 { return $false }
    }
}

# フォルダ選択ダイアログ
function Set-Directory {
    Add-Type -AssemblyName System.Windows.Forms
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
        Description = $args[0]
    }

    # フォルダ選択の有無を判定
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
        return $true, $folderBrowser.SelectedPath 
    }
    else { 
        return $false, $folderBrowser.SelectedPath 
    }
}

# OSバージョンの判定
function Get-MountImageVer {
    # wimのバージョン取得
    $Logs = Dism /Image:$offlineDir /Get-Help
    foreach ( $Log in $Logs ) {
    
        # マッチング文字
        $MatchString = "イメージのバージョン: "
    
        # 文字列ヒット
        if ( $Log -match "(?<Ver>$MatchString *[0-9.]+)") {
            # ヒットした文字列からマッチング文字を削除
            $SelectData = $Matches.Ver -replace $MatchString, ""
        
            # 分割
            $BuildNumber = $SelectData.split(".")
        }
    }
    $OSSupportFlag = $true
    # Windowsバージョン判定
    switch ($BuildNumber[2]) {
        { $_ -ge 22000 } {
            $OSName = "Windows 11"
            break
        }
        20348 {
            $OSName = "Windows Server 2022"
            break
        }
        { ($_ -le 19046) -And ($_ -ge 10240) } {
            $OSName = "Windows 10"
            break
        }
        { ($_ -le 9600) -And ($_ -ge 9200) } {
            $OSName = "Windows 8"
            $OSSupportFlag = $false
            break
        }
        { ($_ -le 7601) -And ($_ -ge 7600) } {
            $OSName = "Windows 7"
            $OSSupportFlag = $false
            break
        }
        default {
            $OSName = "UnknownOS"
            $OSSupportFlag = $false
        }    
    }
    return @{
        Name        = $OSName
        Build       = $BuildNumber[2]
        Update      = $BuildNumber[3]
        SupportFlag = $OSSupportFlag
    }
}

# ドライバのインストール
function Install-Driver {
    $driverInstallMsg = @{
        Title   = "ドライバのインストール"
        Message = "ドライバをインストールしてよろしいですか？"
        YesMsg  = "ドライバのインストールをします。"
        NoMsg   = "ドライバのインストールを中止します。"
    }
    $driverInstall = Get-YesNoDialog $driverInstallMsg
    if ($driverInstall -eq $true) {
        Write-Output ("OS共通のドライバのインストールをします。")
        $driverCommonDir = Join-Path $driverDir "Common"
        Dism /Image:$offlineDir /Add-Driver /Driver:$driverCommonDir /recurse

        if ($OSInfo.SupportFlag -ne $false) {
            if ($OSInfo.Name -eq "Windows 11") {
                $driverOSDir = Join-Path $driverDir "Win11"
            }
            elseif ($OSInfo.Name -eq "Windows Server 2022") {
                $driverOSDir = Join-Path $driverDir "WS2022"
            }
            elseif ($OSInfo.Name -eq "Windows 10") {
                $driverOSDir = Join-Path $driverDir "Win10"
            }
            Write-Output ($OSInfo.Name + "用のドライバをインストールします。")
            Dism /Image:$offlineDir /Add-Driver /Driver:$driverOSDir /recurse
        }
        Write-Output ("ドライバのインストールが完了しました。")
    }
    else {
        Write-Output ("ドライバのインストールを中止しました。")
    }
    pause
}

function Remove-UWP {
    # プロビジョニングされたUWPアプリの削除
    if ($editTerget -eq $installWim) {
        if ($OSInfo.Build -ge 19041) {
            $deleteUwpAppsMsg = @{
                Title   = "プロビジョニングされたUWPアプリの削除"
                Message = "プロビジョニングされたUWPアプリを削除してよろしいですか？"
                YesMsg  = "プロビジョニングされたUWPアプリの削除をします。"
                NoMsg   = "プロビジョニングされたUWPアプリの削除を中止します。"
            }
            $deleteUwpApps = Get-YesNoDialog $deleteUwpAppsMsg
            if ($deleteUwpApps -eq $true) {
                if ($OSInfo.Build -eq 22621) {
                    Write-Output ("Windows 11 22H2のプロビジョニングされたUWPアプリの削除をします。")
                }
                elseif ($OSInfo.Build -eq 22000) {
                    Write-Output ("Windows 11 21H2のプロビジョニングされたUWPアプリの削除をします。")
                }
                elseif (($OSInfo.Build -le 19046) -And ($OSInfo.Build -ge 19041)) {
                    Write-Output ("Windows 10 2004-22H2のプロビジョニングされたUWPアプリの削除をします。")
                }
            }
            else {
                Write-Output ("プロビジョニングされたUWPアプリの削除を中止しました。")
            }
        }
        else {
            Write-Output ("プロビジョニングされたUWPアプリの削除は、Windows 10 2004-22H2、Windows 11以外はできません。")
        }
    
    }
    else {
        Write-Output ("プロビジョニングされたUWPアプリの削除はinstall.wim以外には行えません。")
    }
    pause
}

function Install-WindowsUpdate {
    # 更新プログラムのインストール
    if ($OSInfo.SupportFlag -ne $false) {
        $updateProgramMsg = @{
            Title   = "更新プログラム"
            Message = "更新プログラムをインストールしてよろしいですか？"
            YesMsg  = "更新プログラムのインストールをします。"
            NoMsg   = "更新プログラムのインストールを中止します。"
        }
        $updateProgram = Get-YesNoDialog $updateProgramMsg
        if ($updateProgram -eq $true) {
            Write-Output ("更新プログラムのインストールをします。")
            if ($OSInfo.Name -eq "Windows 11") {
                $updateOSDir = Join-Path $updateDir "Win11"
            }
            elseif ($OSInfo.Name -eq "Windows Server 2022") {
                $updateOSDir = Join-Path $updateDir "WS2022"
            }
            elseif ($OSInfo.Name -eq "Windows 10") {
                $updateOSDir = Join-Path $updateDir "Win10"
            }
            # KB 適用
            Dism /Image:$offlineDir /Add-Package /PackagePath:$updateOSDir
            Write-Output ("更新プログラムのインストールが完了しました。")
        }
        else {
            Write-Output ("更新プログラムのインストールを中止しました。")
        }
    }
    else {
        Write-Output ("更新プログラムのインストールはWindows 10以上でないと行えません。")
    }
    pause
}

function Set-Registry {
    # レジストリの編集
    Write-Output ("レジストリの編集をします。")
    reg load HKEY_USERS\DefaultUser (Join-Path $offlineDir "Users\Default\ntuser.dat")
    regedit.exe
    Write-Output ("HKEY_LOCAL_MACHINE\Offline を編集してください。")
    Write-Output ("regeditでの編集が終了したら続行してください。")
    pause
    reg unload HKEY_USERS\DefaultUser
    Write-Output ("レジストリの編集が完了しました。")
    pause
}

# Dismマウント後
function Mount-Dism {
    $offlineEscape = $false
    while ($offlineEscape -eq $false) {
        $OSInfo = Get-MountImageVer
        Clear-Host
        Write-Output ***********************************************
        Write-Output ("Windows インストールイメージ カスタムスクリプト")
        Write-Output ("オフラインイメージ編集メニュー")
        Write-Output ***********************************************
        Write-Output ($OSInfo.Name + " Build " + $OSInfo.Build + "." + $OSInfo.Update)
        Write-Output ("ドライバ配置先:      " + $driverDir)
        Write-Output ("更新プログラム配置先: " + $updateDir)
        Write-Output ("")
        Write-Output ("1: inf形式のドライバのインストール")
        Write-Output ("2: プロビジョニングされたUWPアプリの削除")
        Write-Output ("3: 更新プログラムのインストール")
        Write-Output ("4: レジストリの編集")
        Write-Output ("5: SMB1の有効化")
        Write-Output ("")
        Write-Output ("7: 変更を保存")
        Write-Output ("8: 変更を破棄して終了")
        Write-Output ("9: 変更を保存して終了")
        Write-Output ("")

        $offlineSelectNumber = Read-Host "実行する作業の番号を入力してください。"
        switch ($offlineSelectNumber) {
            1 {
                Install-Driver
                break
            }
            2 {
                Remove-UWP
                break
            }
            3 {
                Install-WindowsUpdate
                break
            }
            4 {
                Set-Registry
                break
            }
            5 {
                # SMB1の有効化
                Dism /Image:$offlineDir /Enable-Feature /Featurename:"SMB1Protocol" -All
                break
            }
            7 {
                # 変更を保存
                Dism /Image:$OfflineDir /Cleanup-Image /StartComponentCleanup /ResetBase
                Dism /Commit-Image /MountDir:$offlineDir
                break
            }
            8 {
                # 変更を破棄して終了
                Dism /Unmount-WIM /MountDir:$offlineDir /Discard
                $offlineEscape = $true
                pause
                break
            }
            9 {
                # 変更を保存して終了
                Dism /Image:$OfflineDir /Cleanup-Image /StartComponentCleanup /ResetBase
                Dism /Unmount-WIM /MountDir:$offlineDir /Commit
                $offlineEscape = $true
                pause
                break
            }
            default {
                Write-Output ("不正な値が入力がされました。")
                pause
                break
            }
        }
    }
}

Clear-Host

######################################
# 作業ディレクトリの設定
######################################
# 作業ディレクトリ
$workDir = Get-Location

$selectWorkDirMsg = @{
    Title   = "作業ディレクトリ"
    Message = ("作業ディレクトリの場所は " + $workDir + " でよろしいですか？")
    YesMsg  = ("作業ディレクトリの場所 " + $workDir + " で続行します。")
    NoMsg   = "作業ディレクトリの場所を変更します。"
}

$selectWorkDir = Get-YesNoDialog $selectWorkDirMsg
if ($selectWorkDir -eq $false) {
    $SetWorkDir = Set-Directory "作業ディレクトリを選択してください"
    if ($SetWorkDir[0] -eq $true) {
        $workDir = $SetWorkDir[1]
    }
    else {
        $workDirException = Join-Path  ([System.Environment]::GetFolderPath("Desktop")) "DismTemp"
        Write-Output ("作業ディレクトリが選択されていないため " + $workDirException + " を作業ディレクトリにします。")
        $workDir = $workDirException
    }
}
Write-Output ("作業ディレクトリを " + $workDir + " に設定しました。")

# ブート モジュールをセットした先
$windowsPERoot = Join-Path $workDir "WindowsPE"

######################################
# 必要変数(編集不要)
######################################
# ブートモジュール 
$biosBoot = Join-Path $windowsPERoot "fwfiles\etfsboot.com"
$uefiBoot = Join-Path $windowsPERoot "fwfiles\efisys.bin"

# ワークディレクトリ
$isoDir = Join-Path $workDir "ISO"
$global:offlineDir = Join-Path $workDir "Offline"
$driverDir = Join-Path $workDir "Driver"
$updateDir = Join-Path $workDir "KB"
$bootWim = Join-Path $isoDir "sources\boot.wim"
$installWim = Join-Path $isoDir "sources\install.wim"

$editTerget = $installWim

######################################
# ディレクトリ作成
######################################
if ( -not (Test-Path $driverDir)) { mkdir $driverDir }
if ( -not (Test-Path $isoDir)) { mkdir $isoDir }
if ( -not (Test-Path $updateDir)) { mkdir $updateDir }

Write-Output ($isoDir + " フォルダに、isoファイルの中身を展開してください。")
pause

$escape = $false
while ($escape -eq $false) {
    Clear-Host
    Write-Output ***********************************************
    Write-Output ("Windows インストールイメージ カスタムスクリプト")
    Write-Output ("トップメニュー")
    Write-Output ***********************************************
    Write-Output ("作業ディレクトリ:     " + $workDir)
    Write-Output ("WindowsPE:            " + $windowsPERoot)
    Write-Output ("ISOファイル展開先:    " + $isoDir)
    Write-Output ("")
    Write-Output ("現在の編集対象は " + $editTerget + " です。")
    Write-Output ("")
    Write-Output ("1: wimマウント")
    Write-Output ("2: 編集対象の変更")
    Write-Output ("3: wimファイルエクスポート")
    Write-Output ("")
    Write-Output ("5: isoファイルの作成")
    Write-Output ("")
    Write-Output ("9: 終了")
    Write-Output ("")

    $mainSelectNumber = Read-Host "実行する作業の番号を入力してください。"
    switch ($mainSelectNumber) {
        1 {
            # wimマウント

            # インデックスの確認
            Dism /Get-ImageInfo /ImageFile:$editTerget
            $index = Read-Host "インデックス番号を入力してください。"
            Write-Output ("wimファイルをマウントします。")

            # オフライン操作用ディレクトリ作成
            if ( -not (Test-Path $offlineDir)) { mkdir $offlineDir }

            # wim マウント
            Dism /Mount-Image /ImageFile:$editTerget /Index:$index /MountDir:$offlineDir

            pause
            # wimマウント後用の関数呼び出し
            Mount-Dism
            break
        }
        2 {
            ######################################
            # 編集対象の変更
            ######################################
            # 現在の$editTergetが$installWimと同値なら$bootWim、異なるなら$installWimに設定する。
            if ($editTerget -eq $installWim) {
                $editTerget = $bootWim
            }
            else {
                $editTerget = $installWim
            }
            break
        }
        3 {
            # 編集する .wim 指定
            if ($editTerget -eq $installWim) {
                $DestTarget = Join-Path $WorkDir "install.wim"
            }
            else {
                $DestTarget = Join-Path $WorkDir "boot.wim"
            }

            # インデックスの確認
            Dism /Get-ImageInfo /ImageFile:$editTerget
            Write-Output ("wimファイルをマウントします。")
            $index = Read-Host "インデックス番号を入力してください。"
            
            Dism /Export-Image /SourceImageFile:$editTerget /SourceIndex:$index /DestinationImageFile:$DestTarget /Compress:max
            break
        }
        5 {
            $oscdimg = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"
            $IsoFileName = "Win10_22H2_Japanese_x64"
            $OutputIsoFile = Join-Path $WorkDir $IsoFileName
            $NowTime = Get-Date -Format "_yyMMdd-HHmm"

            # ISO 作成オプション
            $IsoOption = " -m -o -u2 -bootdata:2#p0,e,b" + $BiosBoot + "#pEF,e,b" + $UefiBoot + " " + $IsoDir + " " + $OutputIsoFile

            # ファイル名に日時追加
            $IsoOptionDateTime = $IsoOption + $NowTime + ".iso -l" + $IsoFileName
            Start-Process $oscdimg -ArgumentList $IsoOptionDateTime
            break
        }
        9 {
            $escape = $true
            break
        }
        default {
            Write-Output ("不正な値が入力がされました。")
            pause
            break
        }
    }
}
Write-Output ("終了します。")
