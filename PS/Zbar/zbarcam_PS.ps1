<#
.SYNOPSIS
This script makes your webcam a barcode reader device.

Pasting Barcode data to top window using zbarcam.exe.
(You need to install ZBar before using this.)
Double-byte character QRcodes (ex. Japanese) are supported.

"zbarcam path"
 - set path of zbarcam.exe.
"symbology"
 - "ALL": Scan all symbology type.
 - "SELECT": Scan selected symbology type only.
"option"
 - "Symbology type": Scanning data with symbology type information.
 - "No sound": No sound when scanned.
 - "Add{Enter}": Add Enter action after each scanning. 
 - "WEB": If the data includs a WWW URL, open it on default browser.
 - "Confirmation": MessageBox to apply data for each scan. 
 - "Scale": When camera error occurred, change this option.
 - "Density": When the code cannot be scaned, change this option.
These settings above are saved in %USERPROFILE% folder as "zbarcam_PS.ini".

"START":Start scanning by webcam.

Note:
This script uses clipboard operations, so the clipboard will be overwrote.
Tested with ZBar 0.10 in Windows 10 (1909) JA-JP only.
Use this at your own risk. 

.DESCRIPTION
It is recommended that the webcam supports close-up photography.
Some webcams can be adjusted for close-up photography by turning the lens anticlockwise.

"Add {Enter}" option is useful to input in excel worksheets, etc.

.PARAMETER None
<CommonParameters> are not supported.

.NOTES
Version:        1.0 
Author:         ZZZ-tkatoh 

.LINK
http://zbar.sourceforge.net/

#>
using namespace System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# SendKeysを使えるようにする
# 前面にあるウィンドウにキーボード入力を送る
function send([string]$key)
{
	[System.Windows.Forms.SendKeys]::SendWait($key)
	sleep 0.5
}


[Application]::EnableVisualStyles()

$zhelp =""
	$ScriptFullName = $MyInvocation.MyCommand.Path
	Help $ScriptFullName -Examples | ForEach-Object {$zhelp = $zhelp + $_.toString() + "`r`n"}
	
$UserProfile = $env:UserProfile
$file = $UserProfile + "\zbarcam_PS.ini"

$strUAry = New-Object System.Collections.Generic.List[string]


#Barcodeを読取り、クリップボード経由で貼り付ける
#関数を先に定義しないと実行時にエラーになる
function zCap(){
	$zOption=" --xml"
	if($Checkbox_QUIET.Checked){$zOption=$zOption + " --quiet"}
	#if(!$Checkbox_SYMB.Checked){$zOption=$zOption + " --raw"}
	#if($Checkbox_SCALE.Checked){$zOption=$zOption + " --prescale=640x360"}else{$zOption=$zOption + " --prescale=640x480"}
	$zOption=$zOption + " --prescale=" + $Combobox_SCALE.text
	if($RadioButton_SEL.Checked){
		$zOption=$zOption + " --set disable"
		if($Checkbox_EAN_13.Checked){$zOption=$zOption + " --set ean13.enable"}
		if($Checkbox_EAN_8.Checked){$zOption=$zOption + " --set ean8.enable"}
		if($Checkbox_UPC_A.Checked){$zOption=$zOption + " --set upca.enable"}
		if($Checkbox_UPC_E.Checked){$zOption=$zOption + " --set upce.enable"}
		if($Checkbox_ITF.Checked){$zOption=$zOption + " --set i25.enable"}
		if($Checkbox_Code128.Checked){$zOption=$zOption + " --set code128.enable"}
		if($Checkbox_Code39.Checked){$zOption=$zOption + " --set code39.enable"}
		### If EAN-13 disable, ISBN does not seems to be scanned.
		if($Checkbox_ISBN.Checked){
			if($Checkbox_EAN_13.Checked){
				$zOption=$zOption + " --set isbn.enable"
			}else{
				$zOption=$zOption + " --set ean13.enable --set isbn.enable"
			}
		}
		#if($Checkbox_ISBN13.Checked){$zOption=$zOption + " --set isbn13.enable"}
		#if($Checkbox_ISBN10.Checked){$zOption=$zOption + " --set isbn10.enable"}
		if($Checkbox_QR.Checked){$zOption=$zOption + " --set qrcode.enable"}
	}
	### If no symbology option, UPC-A/UPC-E/ISBN-13/ISBN-10 barcode seems to be scanned as EAN-13.
	### isbn10.enable --set isbn13" can be set, but ISBN-13 code is scanned as ISBN-10.
	if($RadioButton_ALL.Checked){
		$zOption=$zOption + " --set upce.enable --set upca.enable --set isbn.enable" 
	}
		$zOption=$zOption + " --set x-density=" + $Combobox_DEN_X.text + "," + "y-density=" + $Combobox_DEN_Y.text

	$psi = New-Object Diagnostics.ProcessStartInfo
	$psi.FileName =$tbZbarcam.Text 
	#if(!$Checkbox_DISP.Checked){$psi.WindowStyle="Minimized"} #video window 実行しても最小化しない
	#$psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Minimized #video window 実行しても最小化しない
	$psi.Arguments  = $zOption
	$psi.UseShellExecute = $false   # プロセス起動時にシェルを使わず、標準出力をリダイレクト可能にする
	$psi.StandardOutputEncoding = [Text.Encoding]::UTF8 # 文字コードを指定
	$psi.RedirectStandardOutput = $true
	#$OutputEncoding が Default だと US-ASCII になっているので
	# S-JIS に変更
	$OutputEncoding = [Console]::OutputEncoding
	#US-ASCII に戻す
	#$OutputEncoding = New-Object System.Text.ASCIIEncoding

	sleep 1
	
	$p = [Diagnostics.Process]::Start($psi)
	$strData=""
	while (!$p.HasExited) {
		$strLine = $p.StandardOutput.ReadLine()
		if($strLine){
			Write-Host $strLine
			if($strData -eq ""){
				$strData = $strLine
			}else{
				#$strData = $strData + "`r`n" + $strLine
				$strData = $strData + "`n" + $strLine
			}
			#if($strLine -eq "</index>"){
			if($strLine.Contains("</symbol>")){
				#Write-Host $strData
				$idx1=$strData.IndexOf("<symbol type='") + 14
				$idx2=$strData.IndexOf("quality='")
				$symbol_type=$strData.Substring($idx1,$idx2-$idx1-2)
				
				
				### If EAN-13 disable, ISBN does not seems to be scanned.
				if(($symbol_type -ne "EAN-13") -or ($Checkbox_EAN_13.Checked) -or ($RadioButton_ALL.Checked)){
					$idx1=$strData.IndexOf("![CDATA[") + 8
					$idx2=$strData.IndexOf("]]></data>")
					$cData=$strData.Substring($idx1,$idx2-$idx1)
					#Write-Host $symbol_type
					#Write-Host $cData
					
					if($Checkbox_SYMB.Checked){
						$dData = $symbol_type + ":" + $cData
					}else{
						$dData =  $cData
					}
					
					$bConf=$true
					if($Checkbox_Conf.Checked){
						if([System.Windows.Forms.MessageBox]::Show($dData, "Confirmation", "OKCancel","None","button1",[System.Windows.Forms.MessageBoxOptions]::ServiceNotification) -eq "Cancel"){$bConf=$false}
					}
					if($bConf){
						if($Checkbox_WEB.Checked){
							$arrcData=$cData.split("`n")
							foreach ($lcData in $arrcData){
								$strP=$lcData.IndexOf("https://",[System.StringComparison]::CurrentCultureIgnoreCase) #:大文字小文字を区別しない
								if($strP -lt 0){
									$strP=$lcData.IndexOf("http://",[System.StringComparison]::CurrentCultureIgnoreCase)
								}
								
								#[Math]::Min($strP, $strPs)
								
								#if($lcData.Length -ge 8){
								#	if(($lcData.Substring(0, 7) -eq "http://") -or ($lcData.Substring(0, 8) -eq "https://")){
							 	#		$strUAry.Add($lcData)
							 	#	}
							 	#}
								if($strP -ge 0){
									$strUAry.Add($lcData.Substring($strP))
								}
							 }
						}
						if(($Checkbox_WEB.Checked) -and ($strUAry.Count -ge 1)){
							foreach ($strU in $strUAry){
								#Write-Host $strU
								Start $strU
								sleep 1
							}

						#$strP=$dData.IndexOf("http://",[System.StringComparison]::CurrentCultureIgnoreCase) #:大文字小文字を区別しない
						#$strPs=$dData.IndexOf("https://",[System.StringComparison]::CurrentCultureIgnoreCase)
						#※なお比較演算子の場合、"-eq"でも大文字小文字を区別しない。区別には"-ceq"を使用する。
						#if(($Checkbox_WEB.Checked) -and (($strP -ge 0) -Or ($strPs -ge 0))){
							#if($strP -ge 0){
							#	$strU=$cData.Substring($strP)
							#}elseif($cData -ge 0){
							#	$strU=$cData.Substring($strPs)
							#}	
							##Write-Host $strU
							#Start $strU
							
						}else{
							$dData | clip
							if($Checkbox_Enter.Checked){send "^v{Enter}"}else{send "^v"}
						
							#Set-Clipboard -Value $p.StandardOutput.ReadLine()
							#Get-Clipboard -Format FileDropList
						}
					}
				}
				
				$strData=""
				$strUAry.Clear()
        	}

		
		
		
		
		
		<###################
			$bConf=$true
			if($Checkbox_Conf.Checked){
				if([System.Windows.Forms.MessageBox]::Show($strLine, "Confirmation", "OKCancel","None","button1",[System.Windows.Forms.MessageBoxOptions]::ServiceNotification) -eq "Cancel"){$bConf=$false}
			}
			if($bConf){
			
				$strP=$strLine.IndexOf("http://",[System.StringComparison]::CurrentCultureIgnoreCase) #:大文字小文字を区別しない
				$strPs=$strLine.IndexOf("https://",[System.StringComparison]::CurrentCultureIgnoreCase)
				#※なお比較演算子の場合、"-eq"でも大文字小文字を区別しない。区別には"-ceq"を使用する。
				if(($Checkbox_WEB.Checked) -and (($strP -ge 0) -Or ($strPs -ge 0))){
					if($strP -ge 0){
						$strU=$strLine.Substring($strP)
					}elseif($strPs -ge 0){
						$strU=$strLine.Substring($strPs)
					}	
					Start $strU
				}else{
					$strLine | clip
					if($Checkbox_Enter.Checked){send "^v{Enter}"}else{send "^v"}
				
					#Set-Clipboard -Value $p.StandardOutput.ReadLine()
					#Get-Clipboard -Format FileDropList
				}
			}
		###################>
		
		
		
		}
	}
	#$s = $p.StandardOutput.ReadToEnd()
	
	$p.WaitForExit()
}

 
 
# ラベル
$label1 = New-Object Label
$label1.Text = "zbarcam PS"
$label1.Name = "Label1"
$label1.Font = New-Object Drawing.Font("Ariel",24)
$label1.Location = "260, 10"
$label1.AutoSize = $True
 

$label2 = New-Object Label
$label2.Text = "zbarcam path"
$label2.Name = "Label2"
$label2.Font = New-Object Drawing.Font("Ariel",10)
$label2.Location = "20, 60"
$label2.AutoSize = $True

$tbZbarcam = New-Object Textbox
$tbZbarcam.Text = "C:\Program Files (x86)\ZBar\bin\zbarcam.exe"
$tbZbarcam.Font = New-Object Drawing.Font("Ariel",10)
$tbZbarcam.ReadOnly = $True
$tbZbarcam.Size = "500, 20"
$tbZbarcam.Location = "110, 60"

$btSel = New-Object Button
$btSel.Text = "Browse.."
$btSel.Font = New-Object Drawing.Font("Ariel",10)
$btSel.Size = "60, 24"
$btSel.Location = "620, 60"
# ボタンイベント
$btSel_Click = {
	$dialog = New-Object System.Windows.Forms.OpenFileDialog
	$dialog.Filter = "zbarcam.exe|zbarcam.exe"
	$dialog.InitialDirectory = "C:\Program Files (x86)\ZBar\bin"
	$dialog.Title = "select zbarcam.exe"
	# 複数選択を許可したい時は Multiselect を設定する
	#$dialog.Multiselect = $true
	
	if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
		# 複数選択を許可している時は $dialog.FileNames を利用する
		$tbZbarcam.Text = $dialog.FileName
	}

}
$btSel.Add_Click($btSel_Click)


# ボタン
$btn = New-Object Button
$btn.Text = "START"
$btn.Font = New-Object Drawing.Font("Ariel",18)
$btn.Size = "120, 40"
$btn.Location = "500, 360"
 
# ボタンイベント
$btn_Click = {
	$fEx = (Test-Path $tbZbarcam.Text)
	if(!$fEx){
		[System.Windows.Forms.MessageBox]::Show($tbZbarcam.Text + " is not exist.", "No zbarcam", [System.Windows.Forms.MessageBoxButtons]::OK)
		#exit
	}else{
		
		### If EAN-13 disable, ISBN does not seems to be scanned.
		#if($Checkbox_ISBN.Checked){
		#	$Checkbox_EAN_13.Checked=$True
		#}
	
		$frame.Controls |  ForEach-Object {$_.Enabled = $false}
		$frame.WindowState = "Minimized"
		Write-Host "Start scanning..."
		#PowerShellではreturnなしでも関数内で出力された全ての値を返すようになっているため、
		#呼び出しの際に「 | Out-Null」を付与するなどの方法で戻り値を破棄する必要がある。
		ZCap | Out-Null
		$frame.Controls |  ForEach-Object {$_.Enabled = $true}
		$frame.WindowState = "Normal"
		Write-Host "Ready to start."

	}
	
  #($sender, $e) = $this, $_
  #$parent = ($sender -as [Button]).Parent -as [Form]
  #$label = [Label]$parent.Controls["Label1"];
  #$label1.Text =$btn.GetType()
}
$btn.Add_Click($btn_Click)


$btnh = New-Object Button
$btnh.Text = "help"
$btnh.Font = New-Object Drawing.Font("Ariel",18)
$btnh.Size = "120, 40"
$btnh.Location = "100, 360"
 
# ボタンイベント
$btnh_Click = {
	[System.Windows.Forms.MessageBox]::Show($zhelp, "Help", [System.Windows.Forms.MessageBoxButtons]::OK)
}
$btnh.Add_Click($btnh_Click)
 
# symbology グループを作る
$GroupBox_SYMB = New-Object System.Windows.Forms.GroupBox
$GroupBox_SYMB.Location = New-Object System.Drawing.Point(20,100)
$GroupBox_SYMB.size = New-Object System.Drawing.Size(660,100)
$GroupBox_SYMB.Font = New-Object Drawing.Font("Ariel",10)
$GroupBox_SYMB.text = "symbology "

# symbology グループの中のラジオボタンを作る
$RadioButton_ALL = New-Object System.Windows.Forms.RadioButton
$RadioButton_ALL.Location = New-Object System.Drawing.Point(20,40)
$RadioButton_ALL.size = New-Object System.Drawing.Size(60,30)
$RadioButton_ALL.Checked = $True
$RadioButton_ALL.Text = "ALL"
$RadioButton_ALL.Name = "RadioButton_ALL"

# RadioButton_ALLイベント
$RadioButton_ALL_Click = {
  $GroupBox_SYMB.Controls | Where-Object {$_.GetType().ToString() -eq "System.Windows.Forms.Checkbox"} | ForEach-Object {$_.Enabled = $false}
}
$RadioButton_ALL.Add_Click($RadioButton_ALL_Click)

$RadioButton_SEL = New-Object System.Windows.Forms.RadioButton
$RadioButton_SEL.Location = New-Object System.Drawing.Point(100,40)
$RadioButton_SEL.size = New-Object System.Drawing.Size(80,30)
$RadioButton_SEL.Text = "SELECT" 
$RadioButton_SEL.Name = "RadioButton_SEL"

Function RB_SEL_Click(){
  $GroupBox_SYMB.Controls | Where-Object {$_.GetType().ToString() -eq "System.Windows.Forms.Checkbox"} | ForEach-Object {$_.Enabled = $true}
}
# RadioButton_SELイベント
$RadioButton_SEL_Click = {
	RB_SEL_Click | Out-Null
  #$GroupBox_SYMB.Controls | Where-Object {$_.Name -like "$Checkbox*"} | ForEach-Object {$_.Enabled = $true}
}
$RadioButton_SEL.Add_Click($RadioButton_SEL_Click)
 
# Checkbox
$Checkbox_EAN_13 = New-Object Checkbox 
$Checkbox_EAN_13.Text = "EAN-13"
$Checkbox_EAN_13.size = "80,20"
$Checkbox_EAN_13.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_EAN_13.Location = "200, 20"
$Checkbox_EAN_13.Enabled = $false
$Checkbox_EAN_13.Name = "Checkbox_EAN_13"
### If EAN-13 disable, ISBN does not seems to be scanned.
#$Checkbox_EAN_13_Click = {
#	if(!$Checkbox_EAN_13.Checked){
#		$Checkbox_ISBN.Checked=$False
#	}
#}
#$Checkbox_EAN_13.Add_Click($Checkbox_EAN_13_Click)

$Checkbox_EAN_8 = New-Object Checkbox 
$Checkbox_EAN_8.Text = "EAN-8"
$Checkbox_EAN_8.size = "80,20"
$Checkbox_EAN_8.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_EAN_8.Location = "300, 20"
$Checkbox_EAN_8.Enabled = $false
$Checkbox_EAN_8.Name = "Checkbox_EAN_8"

$Checkbox_UPC_A = New-Object Checkbox 
$Checkbox_UPC_A.Text = "UPC-A"
$Checkbox_UPC_A.size = "70,20"
$Checkbox_UPC_A.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_UPC_A.Location = "400, 20"
$Checkbox_UPC_A.Enabled = $false
$Checkbox_UPC_A.Name = "Checkbox_UPC_A"

$Checkbox_UPC_E = New-Object Checkbox 
$Checkbox_UPC_E.Text = "UPC-E"
$Checkbox_UPC_E.size = "70,20"
$Checkbox_UPC_E.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_UPC_E.Location = "500, 20"
$Checkbox_UPC_E.Enabled = $false
$Checkbox_UPC_E.Name = "Checkbox_UPC_E"

$Checkbox_ITF = New-Object Checkbox 
$Checkbox_ITF.Text = "ITF"
$Checkbox_ITF.size = "80,20"
$Checkbox_ITF.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_ITF.Location = "600, 20"
$Checkbox_ITF.Enabled = $false
$Checkbox_ITF.Name = "Checkbox_ITF"

$Checkbox_Code128 = New-Object Checkbox 
$Checkbox_Code128.Text = "Code128"
$Checkbox_Code128.size = "90,20"
$Checkbox_Code128.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_Code128.Location = "200, 60"
$Checkbox_Code128.Enabled = $false
$Checkbox_Code128.Name = "Checkbox_Code128"

$Checkbox_Code39 = New-Object Checkbox 
$Checkbox_Code39.Text = "Code39"
$Checkbox_Code39.size = "80,20"
$Checkbox_Code39.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_Code39.Location = "300, 60"
$Checkbox_Code39.Enabled = $false
$Checkbox_Code39.Name = "Checkbox_Code39"

$Checkbox_ISBN = New-Object Checkbox 
$Checkbox_ISBN.Text = "ISBN"
$Checkbox_ISBN.size = "90,20"
$Checkbox_ISBN.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_ISBN.Location = "400, 60"
$Checkbox_ISBN.Enabled = $false
$Checkbox_ISBN.Name = "Checkbox_ISBN"
### If EAN-13 disable, ISBN does not seems to be scanned.
#$Checkbox_ISBN_Click = {
#	if($Checkbox_ISBN.Checked){
#		$Checkbox_EAN_13.Checked=$True
#	}
#}
#$Checkbox_ISBN.Add_Click($Checkbox_ISBN_Click)
	
<#
$Checkbox_ISBN13 = New-Object Checkbox 
$Checkbox_ISBN13.Text = "ISBN13"
$Checkbox_ISBN13.size = "90,20"
$Checkbox_ISBN13.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_ISBN13.Location = "400, 60"
$Checkbox_ISBN13.Enabled = $false
$Checkbox_ISBN13.Name = "Checkbox_ISBN13"

$Checkbox_ISBN10 = New-Object Checkbox 
$Checkbox_ISBN10.Text = "ISBN10"
$Checkbox_ISBN10.size = "80,20"
$Checkbox_ISBN10.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_ISBN10.Location = "500, 60"
$Checkbox_ISBN10.Enabled = $false
$Checkbox_ISBN10.Name = "Checkbox_ISBN10"
#>

$Checkbox_QR = New-Object Checkbox 
$Checkbox_QR.Text = "QR"
$Checkbox_QR.size = "60,20"
$Checkbox_QR.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_QR.Location = "500, 60"
$Checkbox_QR.Enabled = $false
$Checkbox_QR.Name = "Checkbox_QR"

 
# グループにラジオボタンを入れる
$GroupBox_SYMB.Controls.AddRange(@($RadioButton_ALL, $RadioButton_SEL, $Checkbox_EAN_13, $Checkbox_EAN_8, $Checkbox_UPC_A, $Checkbox_UPC_E, `
									$Checkbox_ITF, $Checkbox_Code128, $Checkbox_Code39, $Checkbox_ISBN, $Checkbox_QR))


# option グループを作る
$GroupBox_OPT = New-Object System.Windows.Forms.GroupBox
$GroupBox_OPT.Location = New-Object System.Drawing.Point(20,200)
$GroupBox_OPT.size = New-Object System.Drawing.Size(660,100)
$GroupBox_OPT.Font = New-Object Drawing.Font("Ariel",10)
$GroupBox_OPT.text = "option "

# --raw:if unchecked →xml出力してparseする(--rawオプション無効)
$Checkbox_SYMB = New-Object Checkbox 
$Checkbox_SYMB.Text = "Symbology type"
$Checkbox_SYMB.size = "130,20"
$Checkbox_SYMB.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_SYMB.Location = "60, 20"
$Checkbox_SYMB.Name = "Checkbox_SYMB"

# Checkbox #'--nodisplay'は閉じられなくなるので設定しない

# --quiet:if checked 
$Checkbox_QUIET = New-Object Checkbox 
$Checkbox_QUIET.Text = "No sound"
$Checkbox_QUIET.size = "90,20"
$Checkbox_QUIET.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_QUIET.Location = "200, 20"
$Checkbox_QUIET.Name = "Checkbox_QUIET"

#Add {Enter}
$Checkbox_Enter = New-Object Checkbox 
$Checkbox_Enter.Text = "Add{Enter}"
$Checkbox_Enter.size = "100,20"
$Checkbox_Enter.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_Enter.Location = "300, 20"
$Checkbox_Enter.Name = "Checkbox_Enter"

# --prescale=640x360:if checked / 640x480:if unchcked→Combobox_SCALEで設定
#$Checkbox_SCALE = New-Object Checkbox 
#$Checkbox_SCALE.Text = "16:9"
#$Checkbox_SCALE.size = "80,20"
#$Checkbox_SCALE.Font = New-Object Drawing.Font("Ariel",10)
#$Checkbox_SCALE.Location = "400, 20"
#$Checkbox_SCALE.Name = "Checkbox_SCALE"

#Browser
$Checkbox_WEB = New-Object Checkbox 
$Checkbox_WEB.Text = "WEB"
$Checkbox_WEB.size = "60,20"
$Checkbox_WEB.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_WEB.Location = "400, 20"
$Checkbox_WEB.Name = "Checkbox_WEB"

#Cconfirmation 
$Checkbox_Conf = New-Object Checkbox 
$Checkbox_Conf.Text = "Confirmation"
$Checkbox_Conf.size = "120,20"
$Checkbox_Conf.Font = New-Object Drawing.Font("Ariel",10)
$Checkbox_Conf.Location = "500, 20"
$Checkbox_Conf.Name = "Checkbox_Conf"

#prescale
$Combobox_SCALE = New-Object System.Windows.Forms.Combobox
$Combobox_SCALE.Location = "140, 56"
$Combobox_SCALE.size = "88,24"
$Combobox_SCALE.DropDownStyle = "DropDown"
$Combobox_SCALE.FlatStyle = "standard"
$Combobox_SCALE.font = New-Object Drawing.Font("Ariel",10)
#$Combobox_SCALE.BackColor = "#005050"
#$Combobox_SCALE.ForeColor = "white"
$Combobox_SCALE.Text="640x480"
$Combobox_SCALE.Name = "Combobox_PS"
[void] $Combobox_SCALE.Items.Add("320x180")
[void] $Combobox_SCALE.Items.Add("320x240")
[void] $Combobox_SCALE.Items.Add("352x288")
[void] $Combobox_SCALE.Items.Add("424x240")
[void] $Combobox_SCALE.Items.Add("640x360")
[void] $Combobox_SCALE.Items.Add("640x480")
[void] $Combobox_SCALE.Items.Add("720x480")
[void] $Combobox_SCALE.Items.Add("800x450")
[void] $Combobox_SCALE.Items.Add("800x600")
[void] $Combobox_SCALE.Items.Add("848x480")
[void] $Combobox_SCALE.Items.Add("960x540")
[void] $Combobox_SCALE.Items.Add("1280x720")
[void] $Combobox_SCALE.Items.Add("1280x960")
[void] $Combobox_SCALE.Items.Add("1280x1024")
[void] $Combobox_SCALE.Items.Add("1440x1080")
[void] $Combobox_SCALE.Items.Add("1920x1080")

$label_SCALE = New-Object Label
$label_SCALE.Text = "Scale"
$label_SCALE.Name = "label_SCALE"
$label_SCALE.Font = New-Object Drawing.Font("Ariel",10)
$label_SCALE.Location = "100, 60"
$label_SCALE.AutoSize = $True

#density
$Combobox_DEN_X = New-Object System.Windows.Forms.Combobox
$Combobox_DEN_X.Location = "330, 56"
$Combobox_DEN_X.size = "40,24"
$Combobox_DEN_X.DropDownStyle = "DropDown"
$Combobox_DEN_X.FlatStyle = "standard"
$Combobox_DEN_X.font = New-Object Drawing.Font("Ariel",10)
#$Combobox_DEN_X.BackColor = "#005050"
#$Combobox_DEN_X.ForeColor = "white"
$Combobox_DEN_X.Text="1"
$Combobox_DEN_X.Name = "Combobox_DEN_X"
[void] $Combobox_DEN_X.Items.Add("0")
[void] $Combobox_DEN_X.Items.Add("1")
[void] $Combobox_DEN_X.Items.Add("2")
[void] $Combobox_DEN_X.Items.Add("3")

$label_DEN_X = New-Object Label
$label_DEN_X.Text = "Density  X:"
$label_DEN_X.Name = "label_DEN"
$label_DEN_X.Font = New-Object Drawing.Font("Ariel",10)
$label_DEN_X.Location = "260, 60"
$label_DEN_X.AutoSize = $True

$Combobox_DEN_Y = New-Object System.Windows.Forms.Combobox
$Combobox_DEN_Y.Location = "400, 56"
$Combobox_DEN_Y.size = "40,24"
$Combobox_DEN_Y.DropDownStyle = "DropDown"
$Combobox_DEN_Y.FlatStyle = "standard"
$Combobox_DEN_Y.font = New-Object Drawing.Font("Ariel",10)
#$Combobox_DEN_Y.BackColor = "#005050"
#$Combobox_DEN_Y.ForeColor = "white"
$Combobox_DEN_Y.Text="1"
$Combobox_DEN_Y.Name = "Combobox_DEN_Y"
[void] $Combobox_DEN_Y.Items.Add("0")
[void] $Combobox_DEN_Y.Items.Add("1")
[void] $Combobox_DEN_Y.Items.Add("2")
[void] $Combobox_DEN_Y.Items.Add("3")

$label_DEN_Y = New-Object Label
$label_DEN_Y.Text = "Y:"
$label_DEN_Y.Name = "label_DEN"
$label_DEN_Y.Font = New-Object Drawing.Font("Ariel",10)
$label_DEN_Y.Location = "380, 60"
$label_DEN_X.AutoSize = $True


# Adding the GroupBox_OPT
#$GroupBox_OPT.Controls.AddRange(@($Checkbox_SYMB,$Checkbox_QUIET,$Checkbox_SCALE,$Checkbox_Enter,$Checkbox_WEB,$Checkbox_Conf,$Combobox_SCALE,$label_SCALE))
$GroupBox_OPT.Controls.AddRange(@($Checkbox_SYMB, $Checkbox_QUIET, $Checkbox_Enter, $Checkbox_WEB, $Checkbox_Conf, `
			$Combobox_SCALE, $label_SCALE, $Combobox_DEN_X, $label_DEN_X, $Combobox_DEN_Y, $label_DEN_Y))



function ZLoad(){
	#Get-Variable |ForEach-Object {$_.Name}| Write-Host 
	
	$fEx = (Test-Path $file)
	if($fEx){
		$inifile = New-Object System.IO.StreamReader($file, [System.Text.Encoding]::GetEncoding("US-ASCII"))
		while (($line = $inifile.ReadLine()) -ne $null)
		{
		    #[System.Windows.Forms.MessageBox]::Show($line, "$file") 
		    if($line.SubString(0,6) -eq "zPath="){
		    	$tbZbarcam.Text = $line.SubString(6)
		    }
		    if(($line.Contains("Checkbox")) -Or ($line.Contains("RadioButton"))){
		    	$frame.Controls.Find($line,$true) | ForEach-Object {$_.Checked = $true}
		    	
		    	#$frame.Controls | Where-Object {$_.Name -eq  $line} | ForEach-Object {$_.Checked = $true} #Children Control Only
		    	#Get-Variable | Where-Object {$_.Name -eq $line} | ForEach-Object {$_.Checked = $true}
		    	
		    	#$frame.Controls.Find($line,$true)[0].Checked = $true #OK
		    	
		    	#$Ctls=$frame.Controls.Find($line,$true) #To prevent  no object error
		    	#if($Ctls.length>0){
		    	#	$Ctl=$Ctls[0]
		    	#	$Ctl.Checked = $true
		    	#}
		    	
		    	#Foreach ($x in (Get-variable | Where-Object {$_.Name -eq  $line})){([Checkbox]$x).Checked = $true} #Cast Error
		    }
		    if($line.Contains("Combobox")){
		    	$frame.Controls.Find($line.SubString(0,$line.IndexOf("=")),$true) | ForEach-Object {$_.Text = $line.SubString($line.IndexOf("=")+1)}
		    	#Write-Host $line.SubString($line.IndexOf("=")+1)
		    }
		}
		$inifile.Close()
		if ($RadioButton_SEL.Checked){RB_SEL_Click | Out-Null}
		Write-Host "Ready to start."
	}
}

function ZSave(){
	$txtSave="zPath=" + $tbZbarcam.Text 
	
	#$Ctls=$frame.Controls.Find("Checkbox*",$true) #No wildcard support
	
	$Ctls=$GroupBox_SYMB.Controls
	foreach ($Ctl in $Ctls){
		if (($Ctl.Name.Contains("Checkbox")) -Or ($Ctl.Name.Contains("RadioButton"))){
			if($Ctl.Checked){
				$txtSave=$txtSave + "`r`n" + $Ctl.Name
			}
		}
	}
	$Ctls=$GroupBox_OPT.Controls
	foreach ($Ctl in $Ctls){
		if (($Ctl.Name.Contains("Checkbox")) -Or ($Ctl.Name.Contains("RadioButton"))){
			if($Ctl.Checked){
				$txtSave=$txtSave + "`r`n" + $Ctl.Name
			}
		}
		if ($Ctl.Name.Contains("Combobox")){
			$txtSave=$txtSave + "`r`n" + $Ctl.Name + "=" + $Ctl.Text
		}
	}
	
	#New-Item -Path $UserProfile -Name zbarcam_PS.ini -Type File -Force
	New-Item -Path $file -Type File -Force
	Write-Output $txtSave| Out-File -FilePath $file -Append
}


# フォーム
$frame = New-Object Form
$frame.Text = "zbarcam_PS"
$frame.Size = "720, 480"
$frame.Controls.AddRange(@($label1,$label2, $btn, $btnh))
$frame.Controls.Add($GroupBox_SYMB)
$frame.Controls.Add($GroupBox_OPT)
$frame.Controls.AddRange(@($tbZbarcam,$btSel))
$frame.add_Load({
	ZLoad | Out-Null
})
$frame.add_Closing({
	ZSave | Out-Null
	
	#{param($sender,$e)
    #$result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to exit?", "Close", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
    #if ($result -ne [System.Windows.Forms.DialogResult]::Yes)
    #{
    #   $e.Cancel= $true
    #}
})

$frame.ShowDialog()


#Get-PnpDevice -Class camera| ?{ $_.Status -eq "OK" } | ft FriendlyName -AutoSize
#Get-PnpDevice -Class camera| ?{ $_.Status -eq "OK" } | fl FriendlyName
