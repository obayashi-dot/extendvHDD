# ----------------------
# ���O�o�͗p�֐�
# ----------------------
# �y�@�\�T�v�z
#      ���b�Z�[�W���A��ʂƃe�L�X�g�t�@�C���ɒǉ������݂��܂��B
#      �o�͎��ɓ�����ǉ��\�����邱�Ƃ��\�ł��B
# �y�����z
#      msg  �\�����郁�b�Z�[�W
#      file �o�͂���e�L�X�g�t�@�C���̊��S�p�X
#      dspTimeFlag ���b�Z�[�W�ւ̓����ɒǉ��o�́i0:�ǉ����Ȃ�, 1:�ǉ�����
function msgOutput([String]$msg, [String]$file, [boolean]$dspTimeFlag) {
�@if ($dspTimeFlag) {
�@�@$WkLogTimeMsg = Get-Date -format yyyy/MM/dd-HH:mm:ss
�@�@Write-Host ($WkLogTimeMsg + '  : ' + $msg)
�@�@Write-Output ($WkLogTimeMsg + '  : ' + $msg) | out-file $file Default -append
�@} else {
�@�@Write-Host $msg
�@�@Write-Output  $msg | out-file $file Default -append
�@}
}

############################################
#�ϐ�
############################################
#ViServer�ڑ��p
$ViServer = ""
$ViUser = ""
$ViPass = ""

#VD�̃��[�J���A�h�~�j�X�g���[�^�[
$guestUser="Administrator"
$guestPass=""

#���̑���{�ݒ�
$dt = (get-date -format yyyyMMddHHmm)
$scriptfile = $MyInvocation.MyCommand.Name.Tostring()
$scriptDir = (Split-Path $MyInvocation.MyCommand.Path -parent).Tostring() + "\"
$scriptFullpath = $MyInvocation.MyCommand.Path.Tostring()
$logDir = $scriptDir + "\logs\$dt\"
$logFile = $logDir + $scriptFullpath.split("\")[-1] + "_" + $dt +".log"
$resultFile = $logDir + "Result" + "_" + $dt +".csv"
$vdlist = $scriptDir  + "vdlist.csv"

#�g���e��
[int]$extendHDD = 5

############################################
#�����Ώ�VD�Ŏ��s����R�}���h
############################################

#�f�B�X�N�p�[�e�B�V�����g��&�s�v�t�H���_�폜
$subScript = "
`#try{Remove-Item -path C:\Windows\SoftwareDistribution\Download -recurse -force -ErrorAction Stop}catch{[int]100;continue}
try{set-content -path C:\diskpart.txt -value 'rescan`r`nselect vol c:`r`nextend`r`nrescan'}catch{[int]110;continue};
diskpart /s C:\diskpart.txt
if (`$LastExitCode -ne 0)
{
	[int]120
	continue
}
try{Remove-Item C:\diskpart.txt -ErrorAction Stop}catch{[int]130;continue}
try{Remove-Item -path C:\Windows\SoftwareDistribution\Download -recurse -force -ErrorAction Stop}catch{[int]140;continue}
"


############################################
#CSV��ǂݍ���
#�w���vdlist.csv����A�����Ώۂ�vd�f�[�^���C���|�[�g����B
#����-vd�Ɏw�肪����΁A�P��VD�������[�h�ɂȂ�
############################################
[array]$vdlistData=(Get-Content $vdlist | %{$_.split()} | Select-String -Pattern "\S" | %{$_.line})

#�d�����[�U��������x�����I��
if($vdlistData | where-object {if(($vdlistData -eq $_).count -ge 2){ $_ }} | sort | gu ){
	msgOutput "�d�����[�U�����܂��B�I�����܂��B" $logFile 1
	exit 100
}

############################################
#���O�p�t�H���_�쐬
############################################
if ( ( Test-Path -LiteralPath $logDir -PathType Container) -eq $False ) {
	New-Item $logDir -ItemType Directory |out-null
}

############################################
#���s�O�m�F
############################################
#�I�����̍쐬
$typename = "System.Management.Automation.Host.ChoiceDescription"
$yes = new-object $typename("&Yes","���s����")
$no  = new-object $typename("&No","���s���Ȃ�")

#�I�����R���N�V�����̍쐬
$assembly= $yes.getType().AssemblyQualifiedName
$choice = new-object "System.Collections.ObjectModel.Collection``1[[$assembly]]"
$choice.add($yes)
$choice.add($no)

#�I���v�����v�g�̕\��
$answer = $host.ui.PromptForChoice("���{�O�m�F","���L���[�U��HDD��${extendHDD}GB�g�����܂��B��낵���ł���?" + ($vdlistdata|foreach{"`n";$_}),$choice,1)

if ($answer -ne 0) {
	msgOutput "�������L�����Z�����܂����B" $logFile 1
	exit 0
}

############################################
#PowerCLI�ǂݍ���
############################################
msgOutput "PowerCLI��ǂݍ��݂܂��B" $logFile 1
Add-PSSnapin vmware.vimautomation.core

############################################
#VCenter�ڑ�
############################################
msgOutput "vCenter�ɐڑ����܂��B" $logFile 1
connect-viserver $ViServer -User $ViUser -Password $ViPass 

foreach( $vd in $vdlistdata ){
	#VD��HDD���擾�p�J�E���^�[������
	[int]$count = 0	

	#���ʎ擾�p�J�X�^���I�u�W�F�N�g�쐬
	$result = new-object psobject
	$result | Add-member noteproperty VD $vd
	
	#VM�̏����擾����
	$VMinfo = get-vm $vd

	#�g���O���zHDD�e�ʂ��擾�B
	$result | Add-member noteproperty Before_vHDD ($VMinfo | Get-HardDisk).CapacityGB
	
	#�g���OHDD�e�ʂ��擾�B
	$result | Add-member noteproperty Before_HDD $([int]((get-VMguest $vd).disks | where-object {$_.path -eq "C:\"}).CapacityGB)

	#�g���Ώۂ̃n�[�h�f�B�X�N���擾
	$targetHDD = Get-harddisk $vd | where-object{$_.name -eq "�n�[�h �f�B�X�N 1"} 

	#���zHDD�ǉ�
	msgOutput "${vd}:�n�[�h�f�B�X�N��ǉ����܂��B" $logFile 1
	Set-harddisk -HardDisk $targetHDD -CapacityKB (($result.Before_vHDD + $extendHDD ) *1MB) -GuestUser $GuestUser -GuestPassword $guestPass -confirm:$false -ErrorAction:SilentlyContinue

	#�p�[�e�B�V�����g��&�s�v�t�H���_�폜���܂��B
	msgOutput "${vd}:�p�[�e�B�V�����g�����܂��B" $logFile 1
	msgOutput "${vd}:�s�v�t�H���_�폜���܂��B" $logFile 1

	$buff = invoke-vmscript -vm $vd -scripttext $subScript -Guestuser $GuestUser -GuestPassword $guestPass
	
	#VD���s�R�}���h���ʂ����O�ɋL�^
	msgOutput $buff $logFile 1

	#VD���s�R�}���h�`�F�b�N
	switch (($buff -split '\s+')[-2])
    	{ 
        	110 {msgOutput "diskpart�p�e�L�X�g�t�@�C���쐬���s" $logFile 1} 
        	120 {msgOutput "diskpart�R�}���h���s" $logFile 1} 
        	130 {msgOutput "diskpart�p�e�L�X�g�t�@�C���폜���s" $logFile 1} 
		140 {msgOutput "�uC:Windows\SoftWareDistribution\Download�v�t�H���_�폜���s" $logFile 1} 
        	default {msgOutput "�����s���̃��O�m�F" $logFile 1}
    	}

	#�g���㉼�zHDD�e�ʂ��擾�B
	$result | Add-member noteproperty After_vHDD ($VMinfo | Get-HardDisk).CapacityGB

	#�g����HDD�e�ʂ��擾�B
	$result | Add-member noteproperty After_HDD $([int]((get-VMguest $vd).disks | where-object {$_.path -eq "C:\"}).CapacityGB)

	#�g����HDD�c�e�ʂ��擾�B
	$result | Add-member noteproperty After_FreeHDD $([int]((get-VMguest $vd).disks | where-object {$_.path -eq "C:\"}).FreeSpaceGB).ToString(".00")

	#�p�[�e�B�V�������쒼��͒l������Ɏ��Ȃ����߁A���񂩃��g���C
	while ($result.After_HDD -ne $result.After_vHDD){
		$result.After_HDD = $([int]((get-VMguest $vd).disks | where-object {$_.path -eq "C:\"}).CapacityGB)
		$count +=1
		sleep 5
		#10��ڂ̎��s��NG
		if ($count -ge 10){break}
	}

	#�N���[���A�b�v��Ɏ蓮�ōċN�����邽�߁A�ċN���̓R�����g�A�E�g
	#VD�ċN��
	#msgOutput "${vd}:VD���ċN�����܂��B" $logFile 1
	#Restart-VMGuest -vm $vd

	#���ʎ擾�p�J�X�^���I�u�W�F�N�g���v���X����B
	[PSObject[]]$results += $result
}
############################################
#��������
############################################
msgOutput "�������ʂł��B" $logFile 1
$results
$results | out-file $logFile Default -append
