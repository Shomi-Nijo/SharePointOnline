#######################################################
#
# Title: Get-SiteCollections.ps1
# Author: Shomi_Nijo
# Version: 1.0
# LastUpdate: 2019.12.24
# Description:
# $adminCenterUrlで指定した管理サイトのサイトコレクションを取得・出力する
# 
#######################################################

# パラメータ定義
[CmdletBinding()]
Param(
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

$SiteColListFile = $outdir + "\SiteCollectionList_" + $datestring + ".csv"
$ErrURLListFile = $outdir + "\ErrorURLList_" + $datestring + ".csv"

# ログファイル
$logdir = $myPath + "\log"
$logFile = $logdir + "\log_sitecol_" + $datestring + ".log" 
# ログステータス
$logStatus_Inf = "情報"
$logStatus_War = "警告"
$logStatus_Err = "エラー"

# サイトコレクションの種類
$SiteCategory_SPO = "SharePoint Online"
$SiteCategory_365 = "Office365 グループ"
$SiteCategory_ODB = "OneDrive for Business"

# エンコード文字
$EncodingStr = "UTF8"

# 変数
$credential = $NULL
$exchangeSession = $NULL
$SPOSitesCount = [int]0
$365SitesCount = [int]0
$ODBSitesCount = [int]0
$SitesCount = [int]0
$SPOConnected = $false
$SiteAdminsArray = @()
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

# ヘッダー出力関数
Function WriteHeaderSiteCollection($filepath)
{
    [String]$outStr = 
                        "SiteCategory" `
                        + "," + "Url" `
                        + "," + "ResourceUsageCurrent" `
                        + "," + "ResourceQuotaWarningLevel" `
                        + "," + "ResourceQuota" `
                        + "," + "StorageUsageCurrent" `
                        + "," + "StorageQuotaWarningLevel" `
                        + "," + "StorageQuota" `
                        + "," + "PrimaryOwner" `
                        + "," + "Owners" `
                        + "," + "LastContentModifiedDate" `
                        + "," + "SharingCapability" `
                        + "," + "Title"
                        
    $outStr | Out-File $filePath -Encoding $EncodingStr
}

# サイトコレクション情報出力関数
function WriteRecordSiteCollection($filepath,$SiteCategory,$SiteCollection,$PrimaryOwner,$Owners,$SharingCapability)
{
    WriteLog $logStatus_Inf ("サイトコレクションURL：" + $SiteCollection.Url)
    $outStr = New-Object System.Collections.ArrayList
    $outStr.Add($SiteCategory `
                        + "," + $SiteCollection.url `
                        + "," + $SiteCollection.ResourceUsageCurrent `
                        + "," + $SiteCollection.ResourceQuotaWarningLevel `
                        + "," + $SiteCollection.ResourceQuota `
                        + "," + $SiteCollection.StorageUsageCurrent `
                        + "," + $SiteCollection.StorageQuotaWarningLevel `
                        + "," + $SiteCollection.StorageQuota `
                        + "," + $PrimaryOwner `
                        + "," + $Owners`
                        + "," + $SiteCollection.LastContentModifiedDate `
                        + "," + $SharingCapability`
                        + "," + $SiteCollection.Title)
    $outStr | Out-File $filePath -Encoding $EncodingStr -Append
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

    # Office365接続
    try
    {
        WriteLog $logStatus_Inf "Office365への接続を開始します。"
        Import-Module MSOnline
        Connect-MsolService -Credential $credential
    }
    catch
    {
        WriteLog $logStatus_err ("Office365に接続できませんでした。" + $error[0].Exception.Message)
        break
    }

    # Exchange Online接続
    try
    {
        Get-PSSession | where ComputerName -eq 'outlook.office365.com'
        WriteLog $logStatus_Inf "Exchange Onlineへの接続を開始します。"
        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
        Import-PSSession $exchangeSession -AllowClobber
    }
    catch
    {
        WriteLog $logStatus_err ("Exchange Onlineに接続できませんでした。" + $error[0].Exception.Message)
        break
    }

    # サイトコレクション一覧のヘッダー出力
    WriteHeaderSiteCollection $sitecolListfile

    #--------------------------------------------------------------
    # SharePoint Online サイトコレクション一覧出力
    #--------------------------------------------------------------
    WriteLog $logStatus_Inf "================================================================================================"
    WriteLog $logStatus_Inf "$SiteCategory_SPO サイトコレクション一覧出力を開始します。"

    $sitecollectionList = Get-SPOSite -Limit All | Where-Object {$_.Template -ne "GROUP#0"} | Sort-Object URL
    if(($sitecollectionList -eq $NULL) -or ($sitecollectionList.count -lt 0))
    {
        WriteLog $logStatus_War "$SiteCategory_SPO サイトコレクションがありません。"
    }
    else
    {
        # 各サイトコレクションの取得
        foreach ($sitecollection in $sitecollectionList)
        {
            try
            {
                # OneDrive for Businessのアドレスを格納
                if($sitecollection.Url.Contains("-my.sharepoint.com"))
                {
                    $SPOUserSiteUrl = $sitecollection.Url
                }

        　　  　# サイトコレクション管理者の配列への格納
                try
                {
                    $siteColAdministrators =  Get-SPOUser -site $SiteCollection.Url | Where-Object {$_.IsSiteAdmin -eq $TRUE}
                    if($siteColAdministrators -ne $null)
                    {
                        foreach($siteColAdmin in $siteColAdministrators)
                        {   
                            if($siteColAdmin.IsGroup -eq $true)
                            {
                                $SiteAdminsArray += $siteColAdmin.DisplayName + ";"
                            }
                            else
                            {
                                $SiteAdminsArray += $siteColAdmin.LoginName + ";"
                            }
                        }
                    }
                    else
                    {
                        # サイトコレクション管理者を取得できない場合の出力
                        $SiteAdminsArray = "Error when getting"
                    }

                    # サイトコレクション管理者の出力用処理
                    $Owners = $SiteAdminsArray
                }
                catch
                {
                    # エラーが発生した場合の出力
                    $Owners = "Error when getting"
                }
                finally
                {
                    # サイトコレクション情報再取得
                    $sitecollection = Get-SPOSite -Identity $siteCollection.URL -Detailed # ★変更

                    # プライマリ管理者の出力用処理
                    $PrimaryOwner = $SiteCollection.Owner

                    # 外部共有の出力用処理
                    $SharingCapability = $SiteCollection.SharingCapability.ToString()
                    if($SharingCapability -eq "3")
                    {
                        $SharingCapability = "ExsistingExternalUserSharingOnly"
                    }

                    # プロパティのCSV出力
                    WriteRecordSiteCollection $sitecolListfile $SiteCategory_SPO $SiteCollection $PrimaryOwner $Owners $SharingCapability
                    $SPOSitesCount++
                    $SitesCount++

                    # サイトコレクション管理者の配列の初期化
                    $SiteAdminsArray = @()
                }
            }
            catch
            {
                WriteLog $logStatus_War ("$SiteCategory_SPO サイトコレクションを取得できません。URL：" + $sitecollection.Url)
                WriteLog $logStatus_Err ($error[0].Exception.Message + "`n" + $error[0].ScriptStackTrace)
                $sitecollection.Url | Out-File $errURLListFile -Encoding $EncodingStr -append
                $catchError = $true
                continue
            }

        }
        if($catchError -eq $true)
        {
            WriteLog $logStatus_War "$SiteCategory_SPO に取得できなかったサイトコレクションがあります。"
        }
        else
        {
            WriteLog $logStatus_Inf "$SiteCategory_SPO の全てのサイトコレクションを正常に取得しました。" 
        }
    }

    # SharePoint Online サイトコレクション一覧出力終了
    WriteLog $logStatus_Inf ($SiteCategory_SPO + "出力サイトコレクション数：" + $SPOSitesCount.ToString())
    WriteLog $logStatus_Inf "$SiteCategory_SPO サイトコレクション一覧出力を終了します。"
    $catchError = $false

    
    ### Office365 グループ サイトコレクション一覧出力
    WriteLog $logStatus_Inf "================================================================================================"
    WriteLog $logStatus_Inf "$SiteCategory_365 サイトコレクション一覧出力を開始します。"

    #--------------------------------------------------------------
    # Office365グループ サイトコレクション一覧出力
    #--------------------------------------------------------------
    # Office365 グループの取得
    $o365GroupsInfo = Get-UnifiedGroup
    if(($o365GroupsInfo -eq $NULL) -or ($o365GroupsInfo.count -lt 0))
    {
        WriteLog $logStatus_War "$SiteCategory_365 がありません。"
    }
    else
    {
        # 各サイトコレクションの取得
        for ($i=0; $i -lt $o365GroupsInfo.count; $i++)
        {
            try
            {   
                # 各サイトコレクションの取得
                $SiteCollection = Get-SPOSite -Identity $o365GroupsInfo[$i].SharePointSiteUrl -Detailed

                # プライマリ管理者の取得
                $PrimaryOwner = ""
                if(($sitecollection.Owner).Contains("_o") -eq $True)
                {
                    $GUID = ($sitecollection.Owner).Remove(36,2)
                    $PrimaryOwner = (Get-MsolGroup -ObjectId $GUID).DisplayName
                }
                else
                {
                    $PrimaryOwner = $sitecollection.Owner
                }

                # サイトコレクション管理者の取得
                $siteColAdministrators =  Get-SPOUser -site $SiteCollection.Url | Where-Object {$_.IsSiteAdmin -eq $TRUE}
                if($siteColAdministrators -ne $null)
                {
                    foreach($siteColAdmin in $siteColAdministrators)
                    {   
                        if($siteColAdmin.IsGroup -eq $true)
                        {
                            $SiteAdminsArray += $siteColAdmin.DisplayName + ";"
                        }
                        else
                        {
                            $SiteAdminsArray += $siteColAdmin.LoginName + ";"
                        }
                    }
                }
                else
                {
                    $SiteAdminsArray = "Error when getting"
                }
                
                # サイトコレクション情報再取得
                $sitecollection = Get-SPOSite -Identity $siteCollection.URL -Detailed
                
                # サイトコレクション管理者の出力用処理
                $Owners = $SiteAdminsArray
                
                # 外部共有の出力用処理
                $SharingCapability = $SiteCollection.SharingCapability.ToString()
                if($SharingCapability -eq "3")
                {
                    $SharingCapability = "ExsistingExternalUserSharingOnly"
                }
                # プロパティのCSV出力
                WriteRecordSiteCollection $sitecolListfile $SiteCategory_365 $SiteCollection $PrimaryOwner $Owners $SharingCapability
                $365SitesCount++
                $SitesCount++

                # サイトコレクション管理者の配列の初期化
                $SiteAdminsArray = @()    
            }
            catch
            {
                WriteLog $logStatus_War ("$SiteCategory_365 サイトコレクションを取得できません。URL:" + $o365GroupsInfo[$i].SharePointSiteUrl)
                WriteLog $logStatus_Err ($error[0].Exception.Message + "`n" + $error[0].ScriptStackTrace)
                $o365GroupsInfo[$i].SharePointSiteUrl | Out-File $errURLListFile -Encoding $EncodingStr -append
                $catchError = $true
                continue
            }
        }
        if($catchError -eq $true)
        {
            WriteLog $logStatus_War "$SiteCategory_365 に取得できなかったサイトコレクションがあります。"
        }
        else
        {
            WriteLog $logStatus_Inf "$SiteCategory_365 の全てのサイトコレクションを正常に取得しました。" 
        }
    }

    # Office365 グループ サイトコレクション一覧出力終了
    WriteLog $logStatus_Inf ($SiteCategory_365 + "出力サイトコレクション数：" + $365SitesCount.ToString())
    WriteLog $logStatus_Inf "$SiteCategory_365 サイトコレクション一覧出力を終了します。"
    $catchError = $false
    
    #--------------------------------------------------------------
    # OneDrive for Business サイトコレクション一覧出力
    #--------------------------------------------------------------
    WriteLog $logStatus_Inf "================================================================================================"
    WriteLog $logStatus_Inf "$SiteCategory_ODB サイトコレクション一覧出力を開始します。"

    # ユーザーの取得
    $Logins = Get-MsolUser -All | Where-Object  {($_.IsLicensed -eq $True) -and ($_.Usertype -ne "Guest")} | Sort-Object UserPrincipalName

    if(($Logins -eq $NULL) -or ($Logins.count -lt 0))
    {
        WriteLog $logStatus_War "$SiteCategory_ODB サイトコレクションがありません。"
    }
    else
    {
        # 各ユーザーのOneDriveforBusinessの取得
        foreach ($loginObj in $logins) 
        {
            $login = $loginObj.UserPrincipalName #★追加
            try
            {   
                # ユーザー名の置換
                if($login.Contains('@'))
                {
                    $login=$login.Replace('@','_')
                    $login=$login.Replace('.','_')
                    $loginURL= $SPOUserSiteUrl + "personal/"+$login
                }
                else
                {
                    $loginURL= $SPOUserSiteUrl + "personal/"+$login
                }
                # 各サイトコレクションの取得
                $SiteCollection = Get-SPOSite -Identity $loginURL #| Where-Object {$_.Template -eq "SPSPERS#10"}
                
                # サイトコレクション管理者の配列への格納　
                $siteColAdministrators =  Get-SPOUser -site $SiteCollection.Url | Where-Object {$_.IsSiteAdmin -eq $TRUE}
                if($siteColAdministrators -ne $null)
                {
                    foreach($siteColAdmin in $siteColAdministrators)
                    {
                        if($siteColAdmin.IsGroup -eq $true)
                        {
                            $SiteAdminsArray += $siteColAdmin.DisplayName + ";"
                        }
                        else
                        {
                            $SiteAdminsArray += $siteColAdmin.LoginName + ";"
                        }
                    }
                }
                else
                {
                    $SiteAdminsArray = "Error when getting"
                }
                # プライマリ管理者・サイトコレクション管理者の出力用処理
                $PrimaryOwner = $SiteCollection.Owner
                $Owners = $SiteAdminsArray
            
                # 外部共有の出力用処理
                $SharingCapability = $SiteCollection.SharingCapability.ToString()
                if($SharingCapability -eq "3")
                {
                    $SharingCapability = "ExsistingExternalUserSharingOnly"
                }

                # プロパティのCSV出力
                WriteRecordSiteCollection $sitecolListfile $SiteCategory_ODB $SiteCollection $PrimaryOwner $Owners $SharingCapability
                $ODBSitesCount++
                $SitesCount++

                # サイトコレクション管理者の配列の初期化
                $SiteAdminsArray = @()
            }
            catch
            {
                WriteLog $logStatus_War "$SiteCategory_ODB サイトコレクションを取得できません。　URL: $loginURL"
                WriteLog $logStatus_Err ($error[0].Exception.Message + "`n" + $error[0].ScriptStackTrace)
                $loginURL | Out-File $errURLListFile -Encoding $EncodingStr -append
                $catchError = $true
                continue
            }
        }
        if($catchError -eq $true)
        {
            WriteLog $logStatus_War "$SiteCategory_ODB に取得できなかったサイトコレクションがあります。"
        }
        elseif($SitesCount -eq 0)
        {
            WriteLog $logStatus_War "$SiteCategory_ODB サイトコレクションがありません。"
        }
        elseif($SitesCount -ge 1)
        {
            WriteLog $logStatus_Inf "$SiteCategory_ODB の全てのサイトコレクションを正常に取得しました。" 
        }
    }

    # OneDrive for Business サイトコレクション一覧出力終了
    WriteLog $logStatus_Inf ($SiteCategory_ODB + "出力サイトコレクション数：" + $ODBSitesCount.ToString())
    WriteLog $logStatus_Inf "$SiteCategory_ODB サイトコレクション一覧出力を終了します。"
    WriteLog $logStatus_Inf "================================================================================================"
}
catch
{
    WriteLog $logStatus_Err ($error[0].Exception.Message + "`n" + $error[0].ScriptStackTrace)
}
finally
{
    # サイトコレクション数の合計出力
    WriteLog $logStatus_Inf ("出力サイトコレクション数合計：" + $SitesCount.ToString())

    # 接続終了
    $GetS = Get-PSSession
    if($GetS -ne $Null)
    {
        Get-PSSession | Remove-PSSession
        WriteLog $logStatus_Inf "Exchange Onlineの接続を終了します。"
        WriteLog $logStatus_Inf "Office365の接続を終了します。"
    }
    if($SPOConnected -eq $true)
    {
        Disconnect-SPOService
        WriteLog $logStatus_Inf "SharepointOnlineの接続を終了します。"
    }　
    WriteLog $logStatus_Inf "スクリプトを終了します。"
}