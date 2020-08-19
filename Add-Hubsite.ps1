#######################################################
#
# Title: Add-HubSite.ps1
# Author: Shomi_Nijo
# Version: 1.0
# LastUpdate: 2020.08.20
# Description:
# CSVで設定した構成のハブに関連付ける
# 
#######################################################

# パラメータ定義
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True)]
    [ValidateNotNullOrEmpty()]
	[string]$CSVPath,
	[Parameter(Mandatory=$True)]
    [ValidateNotNullOrEmpty()]
	[string]$adminCenterUrl
)

# 初期処理
$Error.Clear()
[String]$myPath = Split-Path $MyInvocation.MyCommand.Path -parent
Set-Location($myPath)
$ErrorActionPreference = "Stop"

# ********************************************
# Settings
# ********************************************
# 定数

# 出力ファイル
$outdir = $myPath + "\out"
$datestring = Get-Date -format "yyyyMMddHHmmss"

# ログファイル
$logdir = $myPath + "\log"
$logFile = $logdir + "\log_sethub_" + $datestring + ".log" 
# ログステータス
$logStatus_Inf = "情報"
$logStatus_War = "警告"
$logStatus_Err = "エラー"

# エンコード文字
$EncodingStr = "UTF8"

# 変数
$credential = $NULL
$SetCount = 0
$SPOConnected = $false
$catchError = $false

# ********************************************
# Function
# ********************************************

# ログファイル出力関数
Function WriteLog($status, $message)
{
    $logmessage = [DateTime]::Now.ToString() + "`t" + $status + "`t" + $message
    $logmessage | Out-File $logFile -Encoding $EncodingStr -append
    Write-Host $logmessage
}

# ********************************************
# MAIN
# ********************************************

try 
{
    # 出力フォルダ作成
    if(![System.IO.Directory]::Exists($outdir))
    {
        [System.IO.Directory]::CreateDirectory($outdir)
    }
    if(![System.IO.Directory]::Exists($logdir))
    {
        [System.IO.Directory]::CreateDirectory($logdir)
    }

    WriteLog $logStatus_Inf "スクリプトを開始します。"

    # 接続アカウント/パスワード入力（管理者アカウント）
    WriteLog $logStatus_Inf "管理者アカウント/パスワードを入力してください。"
    $credential = Get-Credential

    #--------------------------------------
    # 接続
    #--------------------------------------
    # SharepointOnline接続
    try
    {
        WriteLog $logStatus_Inf "SharepointOnlineへの接続を開始します。"
        Import-Module Microsoft.Online.SharePoint.Powershell
        Connect-SPOService -Url $adminCenterUrl -Credential $credential
        $SPOConnected = $true
    }
    catch
    {
        WriteLog $logStatus_err ("SharePointOnlineに接続できませんでした。" + $error[0].Exception.Message)
        break
    }

    # ハブサイト設定一覧の読み込み
    if((Get-ChildItem $CSVPath).Extension -eq ".csv")
    {
        $CSV = Import-Csv -Path $CSVPath -Encoding $EncodingStr
        WriteLog $logStatus_Inf "ファイル $CSVPath を読み込みます。"
    }
    else
    {
        WriteLog $logStatus_Err "ファイル $CSVPath の拡張子が不適切です。CSVファイルを指定してください。"
        break
    }
    
    # ハブサイト設定
    WriteLog $logStatus_Inf "ハブサイトの設定を開始します。"

        # ファイルの行毎にハブサイトを設定
        foreach ($line in $CSV)
        {
            try
            {
                # ハブサイトに追加
                Add-SPOHubSiteAssociation -Site $line.SiteURL -HubSite $line.HubSiteURL
                $SetCount++
                WriteLog $logStatus_Inf ("ハブサイト " + $line.HubSiteURL + "に" + $line.SiteURL + "を追加しました。")
             }
             catch
             {
                WriteLog $logStatus_War ("ハブサイトの関連付けができませんでした。URL:" + $line.SiteURL)
                WriteLog $logStatus_Err ($error[0].Exception.Message + "`n" + $error[0].ScriptStackTrace)
                $line.SiteURL + "," + $line.HubSiteURL | Out-File $errListFile -Encoding $EncodingStr -append
                $errorCount++
                $catchError = $true
                continue
             }
        }
        if($catchError -eq $true)
        {   
            WriteLog $logStatus_War "ハブサイトを設定できなかった行があります。"
        }
        else
        {
            if($SetCount -eq 0)
            {
                WriteLog $logStatus_War "設定できませんでした。"
            }
            elseif($SetCount -ge 1)
            {
                WriteLog $logStatus_Inf "設定一覧ファイルの全ての行を設定しました。" 
            }
        }

}
catch
{
    WriteLog $logStatus_Err ($error[0].Exception.Message + "`n" + $error[0].ScriptStackTrace)
}
finally
{
    # サイトコレクション数の合計出力
    WriteLog $logStatus_Inf ("設定サイト数合計：" + $SetCount.ToString())

    # 接続終了
    if($SPOConnected -eq $true)
    {
        Disconnect-SPOService
        WriteLog $logStatus_Inf "SharepointOnlineの接続を終了します。"
    }　
    WriteLog $logStatus_Inf "スクリプトを終了します。"
}