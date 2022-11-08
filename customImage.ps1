# The character code of this file is ShiftJIS!!!
# �Ǘ��Ҍ����Ŏ��s
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole("Administrators")) { Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs; exit }

######################################
# �֐��̃Z�b�g
######################################

# Yes/No �_�C�A���O
function Get-YesNoDialog {
    $title = $args.Title
    $message = $args.Message

    $tChoiceDescription = "System.Management.Automation.Host.ChoiceDescription"
    $options = @(
        New-Object $tChoiceDescription ("�͂�(&Yes)", $args.YesMsg)
        New-Object $tChoiceDescription ("������(&No)", $args.NoMsg)
    )
    $result = $host.ui.PromptForChoice($title, $message, $options, 0)
    switch ($result) {
        0 { return $true }        
        1 { return $false }
    }
}

# �t�H���_�I���_�C�A���O
function Set-Directory {
    Add-Type -AssemblyName System.Windows.Forms
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
        Description = $args[0]
    }

    # �t�H���_�I���̗L���𔻒�
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
        return $true, $folderBrowser.SelectedPath 
    }
    else { 
        return $false, $folderBrowser.SelectedPath 
    }
}

# �h���C�o�̃C���X�g�[��
function Install-Driver {
    $driverInstallMsg = @{
        Title   = "�h���C�o�̃C���X�g�[��"
        Message = "�h���C�o���C���X�g�[�����Ă�낵���ł����H"
        YesMsg  = "�h���C�o�̃C���X�g�[�������܂��B"
        NoMsg   = "�h���C�o�̃C���X�g�[���𒆎~���܂��B"
    }
    $driverInstall = Get-YesNoDialog $driverInstallMsg
    if ($driverInstall -eq $true) {
        Write-Output ("�h���C�o�̃C���X�g�[�������܂��B")
        Dism /Image:$offlineDirectory /Add-Driver /Driver:$driverDirectory /recurse
        Write-Output ("�h���C�o�̃C���X�g�[�����������܂����B")
    }
    else {
        Write-Output ("�h���C�o�̃C���X�g�[���𒆎~���܂����B")
    }
    pause
}

function Remove-UWP {
    # �v���r�W���j���O���ꂽUWP�A�v���̍폜
    if ($editTerget -eq $installWim) {
        $deleteUwpAppsMsg = @{
            Title   = "�v���r�W���j���O���ꂽUWP�A�v���̍폜"
            Message = "�v���r�W���j���O���ꂽUWP�A�v�����폜���Ă�낵���ł����H"
            YesMsg  = "�v���r�W���j���O���ꂽUWP�A�v���̍폜�����܂��B"
            NoMsg   = "�v���r�W���j���O���ꂽUWP�A�v���̍폜�𒆎~���܂��B"
        }
        $deleteUwpApps = Get-YesNoDialog $deleteUwpAppsMsg
        if ($deleteUwpApps -eq $true) {
            $isWindows11Msg = @{
                Title   = "OS�̊m�F"
                Message = "���̃C���[�W��Windows 11 22H2�ł����H"
                YesMsg  = "�͂��B���̃C���[�W��Windows 11 22H2�ł��B"
                NoMsg   = "�������B���̃C���[�W��Windows 10 21H2�ł��B"
            }
            $isWindows11 = Get-YesNoDialog $isWindows11Msg
            if ($isWindows11 -eq $true) {
                Write-Output ("Windows 11 22H2�̃v���r�W���j���O���ꂽUWP�A�v���̍폜�����܂��B")
                .\Get-ProvisionedAppxPackages_Win10_22H2.ps1
            }
            else {
                Write-Output ("Windows 10 2004-22H2�̃v���r�W���j���O���ꂽUWP�A�v���̍폜�����܂��B")
                .\Get-ProvisionedAppxPackages_Win11_22H2.ps1
            }
            Write-Output ("�v���r�W���j���O���ꂽUWP�A�v���̍폜���������܂����B")
        }
        else {
            Write-Output ("�v���r�W���j���O���ꂽUWP�A�v���̍폜�𒆎~���܂����B")
        }
    }
    else {
        Write-Output ("�v���r�W���j���O���ꂽUWP�A�v���̍폜��install.wim�ȊO�ɂ͍s���܂���B")
    }
    pause
}

function Install-WindowsUpdate {
    # �X�V�v���O�����̃C���X�g�[��
    if ($editTerget -eq $installWim) {
        $updateProgramMsg = @{
            Title   = "�X�V�v���O����"
            Message = "�X�V�v���O�������C���X�g�[�����Ă�낵���ł����H"
            YesMsg  = "�X�V�v���O�����̃C���X�g�[�������܂��B"
            NoMsg   = "�X�V�v���O�����̃C���X�g�[���𒆎~���܂��B"
        }
        $updateProgram = Get-YesNoDialog $updateProgramMsg
        if ($updateProgram -eq $true) {
            Write-Output ("�X�V�v���O�����̃C���X�g�[�������܂��B")
            # �K�p���� KB �擾
            [array]$kbs = Get-ChildItem $updateDirectory\*.msu -Recurse
            # KB �K�p
            foreach ( $kb in $kbs ) {
                $kbFullName = $kb.FullName
                Dism /Image:$offlineDirectory /Add-Package /PackagePath:$kbFullName
            }
            Write-Output ("�X�V�v���O�����̃C���X�g�[�����������܂����B")
        }
        else {
            Write-Output ("�X�V�v���O�����̃C���X�g�[���𒆎~���܂����B")
        }
    }
    else {
        Write-Output ("�X�V�v���O�����̃C���X�g�[����install.wim�ȊO�ɂ͍s���܂���B")
    }
    pause
}

function Set-Registry {
    # ���W�X�g���̕ҏW
    Write-Output ("���W�X�g���̕ҏW�����܂��B")
    reg load HKEY_USERS\DefaultUser (Join-Path $offlineDirectory "Users\Default\ntuser.dat")
    regedit.exe
    Write-Output ("HKEY_LOCAL_MACHINE\Offline ��ҏW���Ă��������B")
    Write-Output ("regedit�ł̕ҏW���I�������瑱�s���Ă��������B")
    pause
    reg unload HKEY_USERS\DefaultUser
    Write-Output ("���W�X�g���̕ҏW���������܂����B")
    pause
}

# Dism�}�E���g��
function Mount-Dism {
    $offlineEscape = $false
    while ($offlineEscape -eq $false) {
        Clear-Host
        Write-Output ***********************************************
        Write-Output ("Windows �C���X�g�[���C���[�W �J�X�^���X�N���v�g")
        Write-Output ("�I�t���C���C���[�W�ҏW���j���[")
        Write-Output ***********************************************
        Write-Output ("�h���C�o�z�u��:      " + $driverDirectory)
        Write-Output ("�X�V�v���O�����z�u��: " + $updateDirectory)
        Write-Output ("")
        Write-Output ("1: inf�`���̃h���C�o�̃C���X�g�[��")
        Write-Output ("2: �v���r�W���j���O���ꂽUWP�A�v���̍폜")
        Write-Output ("3: �X�V�v���O�����̃C���X�g�[��")
        Write-Output ("4: ���W�X�g���̕ҏW")
        Write-Output ("5: SMB1�̗L����")
        Write-Output ("")
        Write-Output ("7: �ύX��ۑ�")
        Write-Output ("8: �ύX��j�����ďI��")
        Write-Output ("9: �ύX��ۑ����ďI��")
        Write-Output ("")

        $offlineSelectNumber = Read-Host "���s�����Ƃ̔ԍ�����͂��Ă��������B"
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
                # SMB1�̗L����
                Dism /Image:$offlineDirectory /Enable-Feature /Featurename:"SMB1Protocol" -All
            }
            7 {
                # �ύX��ۑ�
                Dism /Image:$OfflineDirectory /Cleanup-Image /StartComponentCleanup /ResetBase
                Dism /Commit-Image /MountDir:$offlineDirectory
            }
            8 {
                # �ύX��j�����ďI��
                Dism /Unmount-WIM /MountDir:$offlineDirectory /Discard
                $offlineEscape = $true
                pause
                break
            }
            9 {
                # �ύX��ۑ����ďI��
                Dism /Image:$OfflineDirectory /Cleanup-Image /StartComponentCleanup /ResetBase
                Dism /Unmount-WIM /MountDir:$offlineDirectory /Commit
                $offlineEscape = $true
                pause
                break
            }
            default {
                Write-Output ("�s���Ȓl�����͂�����܂����B")
                pause
                break
            }
        }
    }
}

Clear-Host

######################################
# ��ƃf�B���N�g���̐ݒ�
######################################
# ��ƃf�B���N�g��
$workDirectory = Get-Location

$selectWorkDirectoryMsg = @{
    Title   = "��ƃf�B���N�g��"
    Message = ("��ƃf�B���N�g���̏ꏊ�� " + $workDirectory + " �ł�낵���ł����H")
    YesMsg  = ("��ƃf�B���N�g���̏ꏊ " + $workDirectory + " �ő��s���܂��B")
    NoMsg   = "��ƃf�B���N�g���̏ꏊ��ύX���܂��B"
}

$selectWorkDirectory = Get-YesNoDialog $selectWorkDirectoryMsg
if ($selectWorkDirectory -eq $false) {
    $SetWorkDirectory = Set-Directory "��ƃf�B���N�g����I�����Ă�������"
    if ($SetWorkDirectory[0] -eq $true) {
        $workDirectory = $SetWorkDirectory[1]
    }
    else {
        $workDirectoryException = Join-Path  ([System.Environment]::GetFolderPath("Desktop")) "DismTemp"
        Write-Output ("��ƃf�B���N�g�����I������Ă��Ȃ����� " + $workDirectoryException + " ����ƃf�B���N�g���ɂ��܂��B")
        $workDirectory = $workDirectoryException
    }
}
Write-Output ("��ƃf�B���N�g���� " + $workDirectory + " �ɐݒ肵�܂����B")

# �u�[�g ���W���[�����Z�b�g������
$windowsPERoot = Join-Path $workDirectory "WindowsPE"

######################################
# �K�v�ϐ�(�ҏW�s�v)
######################################
# �u�[�g���W���[�� 
$biosBoot = Join-Path $windowsPERoot "fwfiles\etfsboot.com"
$uefiBoot = Join-Path $windowsPERoot "fwfiles\efisys.bin"

# ���[�N�f�B���N�g��
$isoDirectory = Join-Path $workDirectory "ISO"
$global:offlineDirectory = Join-Path $workDirectory "Offline"
$driverDirectory = Join-Path $workDirectory "Driver"
$updateDirectory = Join-Path $workDirectory "KB"
$bootWim = Join-Path $isoDirectory "sources\boot.wim"
$installWim = Join-Path $isoDirectory "sources\install.wim"

$editTerget = $installWim

######################################
# �f�B���N�g���쐬
######################################
if ( -not (Test-Path $driverDirectory)) { mkdir $driverDirectory }
if ( -not (Test-Path $isoDirectory)) { mkdir $isoDirectory }
if ( -not (Test-Path $updateDirectory)) { mkdir $updateDirectory }

Write-Output ($isoDirectory + " �t�H���_�ɁAiso�t�@�C���̒��g��W�J���Ă��������B")
pause

$escape = $false
while ($escape -eq $false) {
    Clear-Host
    Write-Output ***********************************************
    Write-Output ("Windows �C���X�g�[���C���[�W �J�X�^���X�N���v�g")
    Write-Output ("�g�b�v���j���[")
    Write-Output ***********************************************
    Write-Output ("��ƃf�B���N�g��:     " + $workDirectory)
    Write-Output ("WindowsPE:            " + $windowsPERoot)
    Write-Output ("ISO�t�@�C���W�J��:    " + $isoDirectory)
    Write-Output ("")
    Write-Output ("���݂̕ҏW�Ώۂ� " + $editTerget + " �ł��B")
    Write-Output ("")
    Write-Output ("1: wim�}�E���g")
    Write-Output ("2: �ҏW�Ώۂ̕ύX")
    Write-Output ("3: wim�t�@�C���G�N�X�|�[�g")
    Write-Output ("")
    Write-Output ("5: iso�t�@�C���̍쐬")
    Write-Output ("")
    Write-Output ("9: �I��")
    Write-Output ("")

    $mainSelectNumber = Read-Host "���s�����Ƃ̔ԍ�����͂��Ă��������B"
    switch ($mainSelectNumber) {
        1 {
            # wim�}�E���g

            # �C���f�b�N�X�̊m�F
            Dism /Get-ImageInfo /ImageFile:$editTerget
            $index = Read-Host "�C���f�b�N�X�ԍ�����͂��Ă��������B"
            Write-Output ("wim�t�@�C�����}�E���g���܂��B")

            # �I�t���C������p�f�B���N�g���쐬
            if ( -not (Test-Path $offlineDirectory)) { mkdir $offlineDirectory }

            # wim �}�E���g
            Dism /Mount-Image /ImageFile:$editTerget /Index:$index /MountDir:$offlineDirectory

            pause
            # wim�}�E���g��p�̊֐��Ăяo��
            Mount-Dism
        }
        2 {
            ######################################
            # �ҏW�Ώۂ̕ύX
            ######################################
            # ���݂�$editTerget��$installWim�Ɠ��l�Ȃ�$bootWim�A�قȂ�Ȃ�$installWim�ɐݒ肷��B
            if ($editTerget -eq $installWim) {
                $editTerget = $bootWim
            }
            else {
                $editTerget = $installWim
            }
            break
        }
        3 {
            # �ҏW���� .wim �w��
            if ($editTerget -eq $installWim) {
                $DestTarget = Join-Path $WorkDirectory "install.wim"
            }
            else {
                $DestTarget = Join-Path $WorkDirectory "boot.wim"
            }

            # �C���f�b�N�X�̊m�F
            Dism /Get-ImageInfo /ImageFile:$editTerget
            Write-Output ("wim�t�@�C�����}�E���g���܂��B")
            $index = Read-Host "�C���f�b�N�X�ԍ�����͂��Ă��������B"
            
            Dism /Export-Image /SourceImageFile:$editTerget /SourceIndex:$index /DestinationImageFile:$DestTarget
        }
        5 {
            $oscdimg = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"
            $IsoFileName = "Win10_22H2_Japanese_x64"
            $OutputIsoFile = Join-Path $WorkDirectory $IsoFileName
            $NowTime = Get-Date -Format "_yyMMdd-HHmm"

            # ISO �쐬�I�v�V����
            $IsoOption = " -m -o -u2 -bootdata:2#p0,e,b" + $BiosBoot + "#pEF,e,b" + $UefiBoot + " " + $IsoDirectory + " " + $OutputIsoFile

            # �t�@�C�����ɓ����ǉ�
            $IsoOptionDateTime = $IsoOption + $NowTime + ".iso -l" + $IsoFileName
            Start-Process $oscdimg -ArgumentList $IsoOptionDateTime
        }
        9 {
            $escape = $true
            break
        }
        default {
            Write-Output ("�s���Ȓl�����͂�����܂����B")
            pause
            break
        }
    }
}
Write-Output ("�I�����܂��B")
