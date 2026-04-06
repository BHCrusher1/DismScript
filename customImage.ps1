# The character code of this file is UTF-8 with BOM !!!
#Requires -RunAsAdministrator
Import-LocalizedData -BindingVariable LangData -filename Message.psd1

######################################
# 共通で使用するYes/Noダイアログの定義
######################################
<#
.SYNOPSIS
ユーザーにYes/Noの選択を促すダイアログを表示します。

.DESCRIPTION
指定されたタイトルとメッセージを持つダイアログを表示し、ユーザーの選択に基づいて結果を返します。

.PARAMETER Msg
ハッシュテーブル。ダイアログのタイトル、メッセージ、および各選択肢の説明を含みます。
- Title: ダイアログのタイトル
- Message: ダイアログのメッセージ
- YesMsg: 「はい」選択肢の説明
- NoMsg: 「いいえ」選択肢の説明

.OUTPUTS
Boolean. ユーザーが「はい」を選択した場合は$true、そうでない場合は$falseを返します。
#>
function Show-YesNoDialog {
    [CmdletBinding()]
    param (
        # ダイアログの設定を含むハッシュテーブル
        [hashtable]$Msg
    )

    [Array]$Choices = @(
        New-Object System.Management.Automation.Host.ChoiceDescription ( $LangData.Yes, $Msg.YesMsg )
        New-Object System.Management.Automation.Host.ChoiceDescription ( $LangData.No, $Msg.NoMsg )
    )

    [Boolean]$result = $Host.UI.PromptForChoice($Msg.Title, $Msg.Message, $Choices, 0)
    if ( $result -eq 0 ) {
        $result = $true
    } else {
        $result = $false
    }
    return $result

}

# Folder selection dialog
function Set-Directory {
    Add-Type -AssemblyName System.Windows.Forms

    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        Description = $args[0]
    }

    # フォルダ選択の有無を判定
    if ( $FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK ) {
        return $true, $FolderBrowser.SelectedPath
    } else {
        return $false
    }
}

# ファイル上書き確認ダイアログ
<#
.SYNOPSIS
ファイルの上書き確認ダイアログを表示します。

.PARAMETER FilePath
上書き対象ファイルのパス

.OUTPUTS
Boolean. ユーザーが「はい」を選択した場合は$true、そうでない場合は$falseを返します。
#>
function Confirm-Overwrite {
    [CmdletBinding()]
    param (
        [string]$FilePath
    )

    # 上書き確認メッセージ
    $overwriteMsg = @{
        Title   = $LangData.Overwrite_Title
        Message = $LangData.Overwrite_Message -f $FilePath
        YesMsg  = $LangData.Overwrite_YesMsg -f $FilePath
        NoMsg   = $LangData.Overwrite_NoMsg -f $FilePath
    }

    # 上書きするか？
    [Boolean]$overwrite = Show-YesNoDialog $overwriteMsg

    # Yesの場合$true、Noの場合$falseを返す
    if ( $overwrite -eq $true ) {
        return $true
    } else {
        return $false
    }
}

# Windowsイメージのアンマウント（変更を破棄）
<#
.SYNOPSIS
マウント中のWindowsイメージをアンマウントし、変更を破棄します。
#>
function Dismount-WindowsImageDiscard {
    [CmdletBinding()]
    param ()

    # マウントされているイメージの取得
    $mountedImages = Get-WindowsImage -Mounted

    # マウント中のイメージが存在する場合、全てアンマウント
    if ( $null -ne $mountedImages ) {
        Write-Host ( $LangData.Dism_CleanUp )
        foreach ( $image in $mountedImages ) {
            Dismount-WindowsImage -Path $image.MountPath -Discard
        }
        Clear-WindowsCorruptMountPoint
    }
}

# ドライバ追加関数
<#
.SYNOPSIS
指定されたディレクトリのドライバをオフラインイメージに追加します。
#>
function Add-WindowsDriverToImage {
    param (
        [switch]$Install,
        [switch]$Boot
    )

    if (-not $Install -and -not $Boot) {
        Write-Warning $LangData.Driver_SwitchMissing
        return
    }

    $targets = @($driverCommonDir)

    if ($Install) {
        $targets += $driverInstallDir
    }

    if ($Boot) {
        $targets += $driverBootDir
    }

    # ドライバ追加確認メッセージ
    $actionName = $LangData.Driver_Add
    $addDriverMsg = @{
        Title   = $actionName
        Message = $LangData.Confirm_Message -f $actionName
        YesMsg  = $LangData.Confirm_YesMsg -f $actionName
        NoMsg   = $LangData.Confirm_NoMsg -f $actionName
    }

    # ドライバを追加するか？
    [Boolean]$addDriver = Show-YesNoDialog $addDriverMsg

    # Yesの場合、ドライバの追加
    if ($addDriver) {
        foreach ($driverPath in $targets) {
            Write-Host ($LangData.Driver_Adding -f $driverPath)
            Add-WindowsDriver -Path $offlineDir -Driver $driverPath -Recurse -ForceUnsigned -ErrorAction SilentlyContinue | Out-Null
        }
    }
}

# ISO作成関数
<#
.SYNOPSIS
ISOファイルを作成します。
#>
function New-IsoFile {
    [CmdletBinding()]
    param ()

    # ISOファイルの作成確認メッセージ
    $actionName = $LangData.ISOFile_Create
    $createIsoMsg = @{
        Title   = $actionName
        Message = $LangData.Confirm_Message -f $actionName
        YesMsg  = $LangData.Confirm_YesMsg -f $actionName
        NoMsg   = $LangData.Confirm_NoMsg -f $actionName
    }

    # ISOファイルを作成するか？
    [Boolean]$createIso = Show-YesNoDialog $createIsoMsg

    # Yesの場合、ISOファイルの作成
    if ( $createIso -eq $true ) {
        # ISO出力ディレクトリの作成
        $isoDir = Join-Path $workDir "ISO"
        if ( -not ( Test-Path $isoDir )) { New-Item -ItemType Directory $isoDir }

        [String]$Oscdimg = Join-Path $workDir "Bin\oscdimg.exe"

        # oscdimg.exeの検証
        if ( -not ( Test-Path $Oscdimg )) {
            Write-Error ( $LangData.OSCDImg_NotFound -f $Oscdimg )
            return
        }

        [String]$BIOSBoot = Join-Path $dvdDir "boot\etfsboot.com"
        [String]$UEFIBoot = Join-Path $dvdDir "efi\microsoft\boot\efisys.bin"
        [String]$ISOLabel = Read-Host $LangData.ISOFile_VolumeLabel
        [String]$ISOFileName = Read-Host $LangData.ISOFile_FileName
        $ISOFileName = Join-Path $isoDir "${ISOFileName}.iso"

        # 既にISOファイルが存在する場合、上書き確認
        if ( Test-Path $ISOFileName ) {
            if ( -not ( Confirm-Overwrite $ISOFileName ) ) {
                Write-Host ( $LangData.Overwrite_NoMsg -f $ISOFileName )
                return
            }
        }

        if ( [string]::IsNullOrWhiteSpace($ISOLabel) ) {
            Start-Process -FilePath $Oscdimg -ArgumentList "-bootdata:2#p0,e,b${BIOSBoot}#pEF,e,b${UEFIBoot} -o -h -m -u2 -udfver102 ${dvdDir} ${ISOFileName}" -Wait
        } else {
            Start-Process -FilePath $Oscdimg -ArgumentList "-bootdata:2#p0,e,b${BIOSBoot}#pEF,e,b${UEFIBoot} -o -h -m -u2 -udfver102 -l${ISOLabel} ${dvdDir} ${ISOFileName}" -Wait
        }
    }
}


######################################
# スクリプト本体
######################################

# 作業ディレクトリの設定
# デフォルトはカレントディレクトリ
$workDir = Get-Location

# 作業ディレクトリ確認メッセージ
$selectWorkDirMsg = @{
    Title   = $LangData.WorkDir_Title
    Message = $LangData.WorkDir_Message -f $workDir
    YesMsg  = $LangData.WorkDir_YesMsg -f $workDir
    NoMsg   = $LangData.WorkDir_NoMsg
}

# 作業ディレクトリはカレントでよいか？
[Boolean]$selectWorkDir = Show-YesNoDialog $selectWorkDirMsg

# Noの場合、作業ディレクトリを選択ダイアログを表示
if ( $selectWorkDir -eq $false ) {
    $SetWorkDir = Set-Directory $LangData.WorkDir_Select
    # 選択ダイアログで選択されたディレクトリを作業ディレクトリにする
    if ( $SetWorkDir[0] -eq $true ) {
        $workDir = $SetWorkDir[1]
    } else {
        # ディレクトリが選択されなかった場合、終了
        Write-Error $LangData.WorkDir_NotSelected
        exit 1
    }
}
Write-Host ( $LangData.WorkDir_Set -f $workDir )

######################################
# リソースディレクトリ作成
######################################
[String]$driverDir = Join-Path $workDir "Driver"
[String]$driverCommonDir = Join-Path $driverDir "Common"
[String]$driverBootDir = Join-Path $driverDir "Boot"
[String]$driverInstallDir = Join-Path $driverDir "Install"
[String]$dvdDir = Join-Path $workDir "DVD"
[String]$offlineDir = Join-Path $workDir "Offline"
[String]$updateDir = Join-Path $workDir "Update"
[String]$bootWim = Join-Path $dvdDir "sources/boot.wim"
[String]$installWim = Join-Path $dvdDir "sources/install.wim"

Dismount-WindowsImageDiscard

if ( -not ( Test-Path $driverCommonDir )) { New-Item -ItemType Directory $driverCommonDir }
if ( -not ( Test-Path $driverBootDir )) { New-Item -ItemType Directory $driverBootDir }
if ( -not ( Test-Path $driverInstallDir )) { New-Item -ItemType Directory $driverInstallDir }
if ( -not ( Test-Path $dvdDir )) { New-Item -ItemType Directory $dvdDir }
if ( -not ( Test-Path $offlineDir )) { New-Item -ItemType Directory $offlineDir }
if ( -not ( Test-Path $updateDir )) { New-Item -ItemType Directory $updateDir }

# isoファイルの中身を展開してください
Write-Host ( $LangData.Iso_Extract -f $dvdDir )

Pause

# installWimファイルの存在チェック
if ( -not ( Test-Path $installWim )) {
    Write-Error ( $LangData.InstallWim_NotFound -f $installWim )
    exit 1
}

$windowsImage = Get-WindowsImage -ImagePath $installWim
$windowsImage | Format-Table -Property ImageIndex, ImageName, ImageDescription
[int]$index = Read-Host $LangData.InstallWim_ImageIndex

# インデックスの検証
if ( -not ( $windowsImage | Where-Object { $_.ImageIndex -eq $index } )) {
    Write-Error $LangData.InstallWim_ImageIndex_Invalid
    exit 1
}

# 選択したイメージの情報取得
$windowsImageInfo = Get-WindowsImage -ImagePath $installWim -Index $index

# Windowsイメージのマウント確認メッセージ
$actionName = $LangData.InstallWim_Mount -f $windowsImageInfo.ImageName
$mountWindowsImageMsg = @{
    Title   = $LangData.InstallWim_Mount -f $windowsImageInfo.ImageName
    Message = $LangData.Confirm_Message -f $actionName
    YesMsg  = $LangData.Confirm_YesMsg -f $actionName
    NoMsg   = $LangData.Confirm_NoMsg -f $actionName
}
# Windowsイメージ(wimファイル)をマウントするか？
[Boolean]$mountWindowsImage = Show-YesNoDialog $mountWindowsImageMsg

# Yesの場合、wimファイルのマウント
if ( $mountWindowsImage -eq $true ) {
    try {
        # イメージのマウント
        Mount-WindowsImage -ImagePath $installWim -Index:$index -Path:$offlineDir

        # マウントしたイメージのバージョン情報取得
        Write-Host ( $LangData.InstallWim_MountedInfo -f $windowsImageInfo.ImageName, $windowsImageInfo.Version )

        # プロビジョニングパッケージの削除確認メッセージ
        $actionName = $LangData.AppxProvisionedPackage_Remove
        $removeAppxProvisionedPackageMsg = @{
            Title   = $actionName
            Message = $LangData.Confirm_Message -f $actionName
            YesMsg  = $LangData.Confirm_YesMsg -f $actionName
            NoMsg   = $LangData.Confirm_NoMsg -f $actionName
        }

        # プロビジョニングパッケージを削除するか？
        [Boolean]$removeAppxProvisionedPackage = Show-YesNoDialog $removeAppxProvisionedPackageMsg

        # Yesの場合、プロビジョニングパッケージの削除を実行
        if ( $removeAppxProvisionedPackage -eq $true ) {
            # プロビジョニングパッケージ削除リストの取得
            [string]$appsListPath = Join-Path $workDir "RemoveAppsList.txt"
            if (Test-Path $appsListPath) {
                $RemoveAppsList = Get-Content -Path $appsListPath | ForEach-Object {
                    ($_ -split '#')[0].Trim()  # '#'以降を削除してトリミング
                } | Where-Object { $_ -ne '' }  # 空行を除外
            } else {
                Write-Warning ( $LangData.AppxProvisionedPackage_ListFile_NotFound -f $appsListPath )
                $RemoveAppsList = @()
            }

            # オフラインイメージのプロビジョニングパッケージの取得
            $AppxProvisionedPackage = Get-AppxProvisionedPackage -Path $offlineDir

            # プロビジョニングパッケージの削除
            foreach ( $DisplayName in $RemoveAppsList ) {
                Write-Host ( $LangData.AppxProvisionedPackage_Removing -f $DisplayName )
                $AppxProvisionedPackage | Where-Object { $_.DisplayName -like $DisplayName } | Remove-AppxProvisionedPackage -Path $offlineDir | Out-Null
            }
            # プロビジョニングパッケージの最適化
            Write-Host ( $LangData.AppxProvisionedPackage_Optimize )
            Optimize-AppXProvisionedPackages -Path $offlineDir | Out-Null
        }

        # インボックスドライバーの削除確認メッセージ
        $actionName = $LangData.InboxDriver_Remove
        $removeInboxDriverMsg = @{
            Title   = $actionName
            Message = $LangData.Confirm_Message -f $actionName
            YesMsg  = $LangData.Confirm_YesMsg -f $actionName
            NoMsg   = $LangData.Confirm_NoMsg -f $actionName
        }

        # インボックスドライバーを削除するか？
        [Boolean]$removeInboxDriver = Show-YesNoDialog $removeInboxDriverMsg

        # Yesの場合、インボックスドライバーの削除を実行
        if ( $removeInboxDriver -eq $true) {
            # インボックスドライバーの削除
            Get-WindowsPackage -Path $offlineDir | Where-Object { $_.PackageName -like "Microsoft-Windows-Ethernet-Client*" -and $_.PackageState -eq "Installed" } | Remove-WindowsPackage -Path $offlineDir | Out-Null
            Get-WindowsPackage -Path $offlineDir | Where-Object { $_.PackageName -like "Microsoft-Windows-Wifi-Client*" -and $_.PackageState -eq "Installed" } | Remove-WindowsPackage -Path $offlineDir | Out-Null
        }

        # ドライバの追加
        Add-WindowsDriverToImage -Install

        # イメージの保存確認メッセージ
        $actionName = $LangData.InstallWim_Save
        $saveImageMsg = @{
            Title   = $actionName
            Message = $LangData.Confirm_Message -f $actionName
            YesMsg  = $LangData.Confirm_YesMsg -f $actionName
            NoMsg   = $LangData.Confirm_NoMsg -f $actionName
        }

        # 編集したイメージを保存するか？
        [Boolean]$saveImage = Show-YesNoDialog $saveImageMsg

        # アンマウント前の確認
        Write-Host ( $LangData.MountWindowsDir_DontOpen -f $offlineDir )
        Start-Sleep -Seconds 5

        # Yesの場合、イメージの保存を実行
        if ( $saveImage -eq $true ) {
            Write-Host ( $LangData.InstallWim_Saving )

            # イメージ全体の最適化
            Repair-WindowsImage -Path $offlineDir -StartComponentCleanup -ResetBase

            Dismount-WindowsImage -Path $offlineDir -Save
            Export-WindowsImage -SourceImagePath $installWim -SourceIndex $index -DestinationImagePath (Join-Path $dvdDir "sources/install2.wim") -CheckIntegrity -CompressionType "max"
            Move-Item -Path ( Join-Path $dvdDir "sources/install2.wim" ) -Destination ( Join-Path $dvdDir "sources/install.wim" ) -Force
        } else {
            Write-Host ( $LangData.InstallWim_Discard )
        }
    } catch {
        Write-Error ( $LangData.ImageProcessing_Error -f $_.Exception.Message )
    } finally {
        $mountedCheck = Get-WindowsImage -Mounted | Where-Object { $_.MountPath -eq $offlineDir }
        if ($null -ne $mountedCheck) {
            Dismount-WindowsImageDiscard
            exit 1
        }
    }
}

# ISOファイルの作成を実行
New-IsoFile
