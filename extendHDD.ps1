# ----------------------
# ログ出力用関数
# ----------------------
# 【機能概要】
#      メッセージを、画面とテキストファイルに追加書込みします。
#      出力時に日時を追加表示することも可能です。
# 【引数】
#      msg  表示するメッセージ
#      file 出力するテキストファイルの完全パス
#      dspTimeFlag メッセージへの日時に追加出力（0:追加しない, 1:追加する
function msgOutput([String]$msg, [String]$file, [boolean]$dspTimeFlag) {
　if ($dspTimeFlag) {
　　$WkLogTimeMsg = Get-Date -format yyyy/MM/dd-HH:mm:ss
　　Write-Host ($WkLogTimeMsg + '  : ' + $msg)
　　Write-Output ($WkLogTimeMsg + '  : ' + $msg) | out-file $file Default -append
　} else {
　　Write-Host $msg
　　Write-Output  $msg | out-file $file Default -append
　}
}

############################################
#変数
############################################
#ViServer接続用
$ViServer = ""
$ViUser = ""
$ViPass = ""

#VDのローカルアドミニストレーター
$guestUser="Administrator"
$guestPass=""

#その他基本設定
$dt = (get-date -format yyyyMMddHHmm)
$scriptfile = $MyInvocation.MyCommand.Name.Tostring()
$scriptDir = (Split-Path $MyInvocation.MyCommand.Path -parent).Tostring() + "\"
$scriptFullpath = $MyInvocation.MyCommand.Path.Tostring()
$logDir = $scriptDir + "\logs\$dt\"
$logFile = $logDir + $scriptFullpath.split("\")[-1] + "_" + $dt +".log"
$resultFile = $logDir + "Result" + "_" + $dt +".csv"
$vdlist = $scriptDir  + "vdlist.csv"

#拡張容量
[int]$extendHDD = 5

############################################
#処理対象VDで実行するコマンド
############################################

#ディスクパーティション拡張&不要フォルダ削除
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
#CSVを読み込み
#指定のvdlist.csvから、処理対象のvdデータをインポートする。
#引数-vdに指定があれば、単体VD処理モードになる
############################################
[array]$vdlistData=(Get-Content $vdlist | %{$_.split()} | Select-String -Pattern "\S" | %{$_.line})

#重複ユーザがいたら警告し終了
if($vdlistData | where-object {if(($vdlistData -eq $_).count -ge 2){ $_ }} | sort | gu ){
	msgOutput "重複ユーザがいます。終了します。" $logFile 1
	exit 100
}

############################################
#ログ用フォルダ作成
############################################
if ( ( Test-Path -LiteralPath $logDir -PathType Container) -eq $False ) {
	New-Item $logDir -ItemType Directory |out-null
}

############################################
#実行前確認
############################################
#選択肢の作成
$typename = "System.Management.Automation.Host.ChoiceDescription"
$yes = new-object $typename("&Yes","実行する")
$no  = new-object $typename("&No","実行しない")

#選択肢コレクションの作成
$assembly= $yes.getType().AssemblyQualifiedName
$choice = new-object "System.Collections.ObjectModel.Collection``1[[$assembly]]"
$choice.add($yes)
$choice.add($no)

#選択プロンプトの表示
$answer = $host.ui.PromptForChoice("実施前確認","下記ユーザのHDDを${extendHDD}GB拡張します。よろしいですか?" + ($vdlistdata|foreach{"`n";$_}),$choice,1)

if ($answer -ne 0) {
	msgOutput "処理をキャンセルしました。" $logFile 1
	exit 0
}

############################################
#PowerCLI読み込み
############################################
msgOutput "PowerCLIを読み込みます。" $logFile 1
Add-PSSnapin vmware.vimautomation.core

############################################
#VCenter接続
############################################
msgOutput "vCenterに接続します。" $logFile 1
connect-viserver $ViServer -User $ViUser -Password $ViPass 

foreach( $vd in $vdlistdata ){
	#VDのHDD情報取得用カウンター初期化
	[int]$count = 0	

	#結果取得用カスタムオブジェクト作成
	$result = new-object psobject
	$result | Add-member noteproperty VD $vd
	
	#VMの情報を取得する
	$VMinfo = get-vm $vd

	#拡張前仮想HDD容量を取得。
	$result | Add-member noteproperty Before_vHDD ($VMinfo | Get-HardDisk).CapacityGB
	
	#拡張前HDD容量を取得。
	$result | Add-member noteproperty Before_HDD $([int]((get-VMguest $vd).disks | where-object {$_.path -eq "C:\"}).CapacityGB)

	#拡張対象のハードディスクを取得
	$targetHDD = Get-harddisk $vd | where-object{$_.name -eq "ハード ディスク 1"} 

	#仮想HDD追加
	msgOutput "${vd}:ハードディスクを追加します。" $logFile 1
	Set-harddisk -HardDisk $targetHDD -CapacityKB (($result.Before_vHDD + $extendHDD ) *1MB) -GuestUser $GuestUser -GuestPassword $guestPass -confirm:$false -ErrorAction:SilentlyContinue

	#パーティション拡張&不要フォルダ削除します。
	msgOutput "${vd}:パーティション拡張します。" $logFile 1
	msgOutput "${vd}:不要フォルダ削除します。" $logFile 1

	$buff = invoke-vmscript -vm $vd -scripttext $subScript -Guestuser $GuestUser -GuestPassword $guestPass
	
	#VD実行コマンド結果をログに記録
	msgOutput $buff $logFile 1

	#VD実行コマンドチェック
	switch (($buff -split '\s+')[-2])
    	{ 
        	110 {msgOutput "diskpart用テキストファイル作成失敗" $logFile 1} 
        	120 {msgOutput "diskpartコマンド失敗" $logFile 1} 
        	130 {msgOutput "diskpart用テキストファイル削除失敗" $logFile 1} 
		140 {msgOutput "「C:Windows\SoftWareDistribution\Download」フォルダ削除失敗" $logFile 1} 
        	default {msgOutput "原因不明のログ確認" $logFile 1}
    	}

	#拡張後仮想HDD容量を取得。
	$result | Add-member noteproperty After_vHDD ($VMinfo | Get-HardDisk).CapacityGB

	#拡張後HDD容量を取得。
	$result | Add-member noteproperty After_HDD $([int]((get-VMguest $vd).disks | where-object {$_.path -eq "C:\"}).CapacityGB)

	#拡張後HDD残容量を取得。
	$result | Add-member noteproperty After_FreeHDD $([int]((get-VMguest $vd).disks | where-object {$_.path -eq "C:\"}).FreeSpaceGB).ToString(".00")

	#パーティション操作直後は値が正常に取れないため、何回かリトライ
	while ($result.After_HDD -ne $result.After_vHDD){
		$result.After_HDD = $([int]((get-VMguest $vd).disks | where-object {$_.path -eq "C:\"}).CapacityGB)
		$count +=1
		sleep 5
		#10回目の失敗でNG
		if ($count -ge 10){break}
	}

	#クリーンアップ後に手動で再起動するため、再起動はコメントアウト
	#VD再起動
	#msgOutput "${vd}:VDを再起動します。" $logFile 1
	#Restart-VMGuest -vm $vd

	#結果取得用カスタムオブジェクトをプラスする。
	[PSObject[]]$results += $result
}
############################################
#処理結果
############################################
msgOutput "処理結果です。" $logFile 1
$results
$results | out-file $logFile Default -append
