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

# OS�o�[�W�����̔���
function Get-MountImageVer {
    # wim�̃o�[�W�����擾
    $Logs = Dism /Image:$offlineDirectory /Get-Help
    foreach ( $Log in $Logs ) {
    
        # �}�b�`���O����
        $MatchString = "�C���[�W�̃o�[�W����: "
    
        # ������q�b�g
        if ( $Log -match "(?<Ver>$MatchString *[0-9.]+)") {
            # �q�b�g���������񂩂�}�b�`���O�������폜
            $SelectData = $Matches.Ver -replace $MatchString, ""
        
            # ����
            $BuildNumber = $SelectData.split(".")
        }
    }
    $OSSupportFlag = $true
    # Windows�o�[�W��������
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
    return $OSName, $BuildNumber[2], $OSSupportFlag
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
        Write-Output ("OS���ʂ̃h���C�o�̃C���X�g�[�������܂��B")
        $driverCommonDirectory = Join-Path $driverDirectory "Common"
        Dism /Image:$offlineDirectory /Add-Driver /Driver:$driverCommonDirectory /recurse

        if ($ImageVer[2] -ne $false) {
            if ($ImageVer[0] -eq "Windows 11") {
                $driverOSDirectory = Join-Path $driverDirectory "Win11"
            }
            elseif ($ImageVer[0] -eq "Windows Server 2022") {
                $driverOSDirectory = Join-Path $driverDirectory "WS2022"
            }
            elseif ($ImageVer[0] -eq "Windows 10") {
                $driverOSDirectory = Join-Path $driverDirectory "Win10"
            }
            Write-Output ($ImageVer[0] + "�p�̃h���C�o���C���X�g�[�����܂��B")
            Dism /Image:$offlineDirectory /Add-Driver /Driver:$driverOSDirectory /recurse
        }
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
        if ($ImageVer[1] -ge 19041) {
            $deleteUwpAppsMsg = @{
                Title   = "�v���r�W���j���O���ꂽUWP�A�v���̍폜"
                Message = "�v���r�W���j���O���ꂽUWP�A�v�����폜���Ă�낵���ł����H"
                YesMsg  = "�v���r�W���j���O���ꂽUWP�A�v���̍폜�����܂��B"
                NoMsg   = "�v���r�W���j���O���ꂽUWP�A�v���̍폜�𒆎~���܂��B"
            }
            $deleteUwpApps = Get-YesNoDialog $deleteUwpAppsMsg
            if ($deleteUwpApps -eq $true) {
                if ($ImageVer[1] -eq 22621) {
                    Write-Output ("Windows 11 22H2�̃v���r�W���j���O���ꂽUWP�A�v���̍폜�����܂��B")
                }
                elseif ($ImageVer[1] -eq 22000) {
                    Write-Output ("Windows 11 21H2�̃v���r�W���j���O���ꂽUWP�A�v���̍폜�����܂��B")
                }
                elseif (($ImageVer[1] -le 19046) -And ($ImageVer[1] -ge 19041)) {
                    Write-Output ("Windows 10 2004-22H2�̃v���r�W���j���O���ꂽUWP�A�v���̍폜�����܂��B")
                }
            }
            else {
                Write-Output ("�v���r�W���j���O���ꂽUWP�A�v���̍폜�𒆎~���܂����B")
            }
        }
        else {
            Write-Output ("�v���r�W���j���O���ꂽUWP�A�v���̍폜�́AWindows 10 2004-22H2�AWindows 11�ȊO�͂ł��܂���B")
        }
    
    }
    else {
        Write-Output ("�v���r�W���j���O���ꂽUWP�A�v���̍폜��install.wim�ȊO�ɂ͍s���܂���B")
    }
    pause
}

function Install-WindowsUpdate {
    # �X�V�v���O�����̃C���X�g�[��
    if ($ImageVer[2] -ne $false) {
        $updateProgramMsg = @{
            Title   = "�X�V�v���O����"
            Message = "�X�V�v���O�������C���X�g�[�����Ă�낵���ł����H"
            YesMsg  = "�X�V�v���O�����̃C���X�g�[�������܂��B"
            NoMsg   = "�X�V�v���O�����̃C���X�g�[���𒆎~���܂��B"
        }
        $updateProgram = Get-YesNoDialog $updateProgramMsg
        if ($updateProgram -eq $true) {
            Write-Output ("�X�V�v���O�����̃C���X�g�[�������܂��B")
            if ($ImageVer[0] -eq "Windows 11") {
                $updateOSDirectory = Join-Path $updateDirectory "Win11"
            }
            elseif ($ImageVer[0] -eq "Windows Server 2022") {
                $updateOSDirectory = Join-Path $updateDirectory "WS2022"
            }
            elseif ($ImageVer[0] -eq "Windows 10") {
                $updateOSDirectory = Join-Path $updateDirectory "Win10"
            }
            # KB �K�p
            Dism /Image:$offlineDirectory /Add-Package /PackagePath:$updateOSDirectory
            Write-Output ("�X�V�v���O�����̃C���X�g�[�����������܂����B")
        }
        else {
            Write-Output ("�X�V�v���O�����̃C���X�g�[���𒆎~���܂����B")
        }
    }
    else {
        Write-Output ("�X�V�v���O�����̃C���X�g�[����Windows 10�ȏ�łȂ��ƍs���܂���B")
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
    $ImageVer = Get-MountImageVer
    while ($offlineEscape -eq $false) {
        Clear-Host
        Write-Output ***********************************************
        Write-Output ("Windows �C���X�g�[���C���[�W �J�X�^���X�N���v�g")
        Write-Output ("�I�t���C���C���[�W�ҏW���j���[")
        Write-Output ***********************************************
        Write-Output ($ImageVer[0] + " Build " + $ImageVer[1])
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
            
            Dism /Export-Image /SourceImageFile:$editTerget /SourceIndex:$index /DestinationImageFile:$DestTarget /Compress:max
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
