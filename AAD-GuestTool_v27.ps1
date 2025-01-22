<#
  スクエニ AAD ゲスト登録用スクリプト
  作成日 : 2022/08/15
  作成者 : 土井

  ※ 本スクリプトの変更履歴は、フッターへ記載
  ※ このスクリプトは Teams や WDG のゲスト登録と招待メール送付用途です。スクエニ踏み台で実行します。
  ※ guestuserinvitation@square-enix.com で接続・実行する前提です。(偽名アカウントで招待メールを送付しないように…)
#>

function StartProc {}
# 終了処理
function EndProc {
    # 接続解除
    Disconnect-MgGraph
    Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear(); Clear-Host
    Write-Host ""
    Write-Host ""
    Write-Host -ForegroundColor Cyan "   終了処理が完了しました。"
    Write-Host ""
    Write-Host ""
}


# Graph 接続
function ConnectGraph {
    Write-Host ""
    Write-Host ""
	Write-Host -ForegroundColor Cyan "   Graph 接続します。"
    Write-Host ""
    Write-Host ""
	$rtn = $true
    $ErrorActionPreference = "silentlycontinue"

    # エラー処理でなぜか try catch がうまくいかず、$? 処理を採用
    Connect-MgGraph -Scopes "User.ReadWrite.All","Sites.ReadWrite.All","Group.Read.All"
    if ($? -eq $False) {
    Write-Host ""
    Write-Host ""
	Write-Host -ForegroundColor Red "   認証に失敗しました。"
    Write-Host ""
    Write-Host ""
	$rtn = $false
    
    }
}

# Graph モジュールインポート
function ImportGraph {
    Write-Host ""
    Write-Host ""
    Write-Host -ForegroundColor Cyan "   認証に成功しました。続けて Graph モジュールをインポートします。"
    Import-Module -Name Microsoft.Graph.Authentication
    Import-Module -Name Microsoft.Graph.Users
    Import-Module -Name Microsoft.Graph.Groups
    Import-Module -Name Microsoft.Graph.Identity.SignIns
    }

# 常時表示されるタイトルバナー
function Disp_Title {
    Clear-Host
    Write-Output " ****************************************"
    Write-Output " * AAD ゲスト管理ツール [SQEX 1.1.0 版] *"
    Write-Output " ****************************************"
    Write-Output ""
    Write-Output ""
}


function Main {
    Disp_Title
    Write-Output "----------------------------------------------------------------------------------"
    Write-Output "< Guest >"
    Write-Output "   1 Azure AD ゲスト追加 (Teams 等)          2 Azure AD ゲスト追加 (WDC外注)"
    Write-Output "   3 ゲストユーザー存在確認                  4 全ゲストユーザー表示"
    Write-Output "   5「外注WDCユーザー」全メンバー表示        6「外注WDCユーザー」全員を CSV 出力 "
    Write-Output "   7 ゲストに招待状再送 (Teams 等)           8 ゲストに招待状再送 (WDC外注)"
    Write-Output "   9 全ゲストユーザー CSV 出力"
    Write-Output ""
    Write-Output "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
    Write-Output "< Others >"
    Write-Output "   XX"
    Write-Output "---------------------------------------------------------------------------------"
    Write-Output "                                            99 終了"
    Write-Output "---------------------------------------------------------------------------------"
    Write-Output ""
    $ret1 = read-host " 処理する番号を入力して下さい。  "
    switch($ret1) {
         "1" {
                Invite-AADGuest
                break
             }
         "2" {
                Invite-WDCGuest
                break
             }
         "3" {
                Check-AADUser 
                break
             }
         "4" {
                Get-AADGuestUser
                break
             }
         "5" {
                Get-WDCGroupMember
                break
             }
         "6" {
                Export-WDCGroupMemberCsv
                break
             }
         "7" {
                ResendAADInvitation
                break
             }
         "8" {
                ResendWDCInvitation
                break
             }
         "9" {
                Export-AADGuestUserCsv
                break
             }
        "99" {
                cls
                EndProc
                Exit
             }
        default {
                Write-Output " 有効な入力ではありません"
                }
    }
    Write-Output " "
    Write-Output " 何か押すとメニューに戻ります。"
    $ret2 = Read-Host " "
    
    main
}

#######################################
#  1 AzureAD ゲスト追加　(Teams 等)   #
#######################################
function Invite-AADGuest {
    Disp_Title

# ゲストのメールアドレス入力要求。if チェックでも使用
Write-Host "   Azure AD に招待するゲストのメールアドレスを入力してください。(用途は Teams 等です。こちらは WDC ゲスト招待ではありません。)" -ForegroundColor Yellow
$AADGuestMail = Read-Host "末尾に改行やスペースが入らないように注意して入力してください。"

# 入力から改行をなくす試行錯誤…
# $WDCGuestMail = $WDCGuestMailInput -replace "`n",""
# $WDCGuestMail = $WDCGuestMailInput.Replace("<br>","")
# $WDCGuestMail = $WDCGuestMailInput | Out-String -Stream | ?{$_ -ne ""}

# if でチェックするために変数格納
$UserCheck = Get-MgUser -Filter "Mail eq '$AADGuestMail'" `
             -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName

# アドレスがすでに AAD に存在しているかをチェック。存在しない場合のみゲスト登録に進む。SEA / SEE もオミットする。
If ($UserCheck.UserType -eq "Member"){

      Write-Host ""
      Write-Host "　このメールアドレスのユーザーは下記の通り、スクエニ組織の内部ユーザーです。本作業は不要です。" -ForegroundColor Yellow
      Write-Host "　※下記に結果が返らない場合は SEE の可能性がございます。" -ForegroundColor Yellow

      $UserCheck | fl DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType

}ElseIf($UserCheck.UserType -eq "Guest"){

      Write-Host ""
      Write-Host "このメールアドレスのユーザーはスクエニ組織の Azure AD に既にゲストとして存在します。(下記参照) " -ForegroundColor Yellow

      $UserCheck | fl DisplayName, Mail, UserPrincipalName, CreatedDateTime, UserType, ExternalUserState, ExternalUserStateChangeDateTime

}Else{

# ゲストの DisplayName 入力
Write-Host ""
Write-Host ""
Write-Host "ゲストの DisplayName を入力してください。"
$GuestDisplayName = Read-Host "※日本名だと性 名 例 : 山田 太郎"

# 招待メールの言語選択
Write-Host ""
Write-Host ""
Write-Host "招待メールの言語を選択します。"
Write-Host "日本語で送付する場合は j 、英語で送付する場合は e を入力して Enter 押下してください。(j / e)"
$Prompt = Read-Host "※ この操作をキャンセルする場合はその他の文字を押下してください。"


# ユーザーの入力結果から、処理を分岐
switch ($Prompt) {
    "j" {
        # j の場合の処理 (日本語招待メール)

        $InvitedUserMessageInfo1 = @{
        CustomizedMessageBody = "スクウェア・エニックス組織のゲストの皆様へ

        いつもお世話になっております。
        スクウェア・エニックス 情報システム部です。

        本メールは、弊社社員経由でスクウェア・エニックスの O365 リソース (Teams 等) 利用依頼があった方にお送りしております。
        弊社のリソースを利用開始使用するにあたり、本メールの下にある [招待の承諾] をクリックし、
        二要素認証の設定とゲストユーザー登録を進めて下さい。

        登録に関して、不明点がございましたら弊社の社員経由でご連絡ください。
        ※本メール受領後 10 日間承諾されない場合は招待が破棄されますのでご注意ください。
        "
        }

        New-MgInvitation -InvitedUserDisplayName $GuestDisplayName -InvitedUserEmailAddress $AADGuestMail `
        -InviteRedirectUrl "https://account.activedirectory.windowsazure.com/?tenantid=0e371789-fac2-4e0e-b7ac-30c4834d6b4e" `
        -InvitedUserMessageInfo $InvitedUserMessageInfo1 `
        -SendInvitationMessage:$true

        # CreationType が Invitation になるまでループ
        do {
        $status = Get-MgUser -Filter "CreationType eq 'Invitation' and Mail eq '$AADGuestMail'" `
        -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName
        }
        until ($status.CreationType -eq "Invitation")

        Write-Host ""
        Write-Host ""
        Write-Host " 新規ゲストユーザーの登録が以下の通り完了しました。招待メールも送付済です。" -ForegroundColor Green
        $status | fl DisplayName, Mail, UserPrincipalName, UserType, CreationType, CreatedDateTime
        }

    "e" {
        # e の場合の処理 (英語招待メール)

        $InvitedUserMessageInfo1e = @{
        CustomizedMessageBody = "To SEJ Guest Users.

        This is Office 365 Admin from SQUARE-ENIX Japan (SEJ) Information Technology Division.

        Please click the following button, and complete to setup 2 factor authentication and register as a guest user.
        After registered, please check you can sign-in to SEJ's Office 365 resouces with your email address.
        If you can’t access it, please let SEJ staffs know.
        
        Regards,
        "
        }

        New-MgInvitation -InvitedUserDisplayName $GuestDisplayName -InvitedUserEmailAddress $AADGuestMail `
        -InviteRedirectUrl "https://account.activedirectory.windowsazure.com/?tenantid=0e371789-fac2-4e0e-b7ac-30c4834d6b4e" `
        -InvitedUserMessageInfo $InvitedUserMessageInfo1e `
        -SendInvitationMessage:$true

        # CreationType が Invitation になるまでループ

        do {
        $status = Get-MgUser -Filter "CreationType eq 'Invitation' and Mail eq '$AADGuestMail'" `
        -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName
        }
        until ($status.CreationType -eq "Invitation")

        Write-Host ""
        Write-Host ""
        Write-Host " 新規ゲストユーザーの登録が以下の通り完了しました。招待メールも送付済です。" -ForegroundColor Green

        Write-Host " 新規ゲストユーザーの登録が以下の通り完了しました。招待メールも送付済です。" -ForegroundColor Green
        $status | fl DisplayName, Mail, UserPrincipalName, UserType, CreationType, CreatedDateTime
        }

    default {
            # j でも e でもない場合の処理
            Write "j でも e でもない文字が入力されました。処理を中止します"
            }
  }
  }
}

#######################################
#  2 WDC ゲスト追加　(WDC外注)        #
#######################################
function Invite-WDCGuest {
    Disp_Title

# ゲストのメールアドレス入力要求。if チェックでも使用
Write-Host "WDC に招待するゲストのメールアドレスを入力してください。" -ForegroundColor Yellow
$WDCGuestMail = Read-Host "末尾に改行やスペースが入らないように注意して入力してください。"

# if でチェックするために変数格納
$UserCheck = Get-MgUser -Filter "Mail eq '$WDCGuestMail'" `
-Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName
             

# アドレスがすでに AAD に存在しているかをチェック。存在しない場合のみゲスト登録に進む。SEA / SEE もオミットする。
If ($UserCheck.UserType -eq "Member"){

      Write-Host ""
      Write-Host "　このメールアドレスのユーザーは下記の通り、スクエニ組織の内部ユーザーです。本作業は不要です。" -ForegroundColor Yellow
      Write-Host "　※下記に結果が返らない場合は SEE の可能性がございます。" -ForegroundColor Yellow

      $UserCheck | fl DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType

      Write-Host ""
      Write-Host ""
      Write-Host "※まれに、「サイトアクセス申請(外注会社)」と記載と書いてあるにもかかわらず、" -ForegroundColor Green
      Write-Host "　「申請メールアドレス」に、スクエニメール (SEA/SEE含む) が記載されていることがあります。" -ForegroundColor Green
      Write-Host "　この場合、ゲストユーザー登録をする必要はないので、作業は不要です。" -ForegroundColor Green
      Write-Host "　その旨、申請者と業務部 岡本さんにメール返信をお願いいたします。" -ForegroundColor Green
      Write-Host ""

}ElseIf($WDCGuestMail -Like "*@us.square-enix.com"){

      Write-Host ""
      Write-Host "このメールアドレスのユーザーは SEA です。本作業は不要です。" -ForegroundColor Yellow
      Write-Host ""
      Write-Host ""
      Write-Host "※まれに、「サイトアクセス申請(外注会社)」と記載と書いてあるにもかかわらず、" -ForegroundColor Green
      Write-Host "　「申請メールアドレス」に、スクエニメール (SEA/SEE含む) が記載されていることがあります。" -ForegroundColor Green
      Write-Host "　この場合、ゲストユーザー登録をする必要はないので、作業は不要です。" -ForegroundColor Green
      Write-Host "　その旨、申請者と業務部 岡本さんにメール返信をお願いいたします。" -ForegroundColor Green
      Write-Host ""
      Write-Host ""

}ElseIf($WDCGuestMail -Like "*@eu.square-enix.com"){

      Write-Host ""
      Write-Host "このメールアドレスのユーザーは SEE です。本作業は不要です。" -ForegroundColor Yellow
      Write-Host ""
      Write-Host ""
      Write-Host "※まれに、「サイトアクセス申請(外注会社)」と記載と書いてあるにもかかわらず、" -ForegroundColor Green
      Write-Host "　「申請メールアドレス」に、スクエニメール (SEA/SEE含む) が記載されていることがあります。" -ForegroundColor Green
      Write-Host "　この場合、ゲストユーザー登録をする必要はないので、作業は不要です。" -ForegroundColor Green
      Write-Host "　その旨、申請者と業務部 岡本さんにメール返信をお願いいたします。" -ForegroundColor Green
      Write-Host ""
      Write-Host ""


}ElseIf($UserCheck.UserType -eq "Guest"){

      Write-Host ""
      Write-Host "このメールアドレスのユーザーはスクエニ組織の Azure AD に既にゲストとして存在します。(下記参照)" -ForegroundColor Yellow
      Write-Host "依頼者と業務部 (岡本さん) に連絡してください。" -ForegroundColor Yellow

      $UserCheck | fl DisplayName, Mail, UserPrincipalName, CreatedDateTime, UserType, ExternalUserState, ExternalUserStateChangeDateTime
      
}Else{

# ゲストの DisplayName 入力
Write-Host ""
Write-Host ""
Write-Host "ゲストの DisplayName を入力してください。"
$WDCGuestDisplayName = Read-Host "※日本名だと性 名 例 : 山田 太郎"

# 招待メールの言語選択
Write-Host ""
Write-Host ""
Write-Host "招待メールの言語を選択します。"
Write-Host "日本語で送付する場合は j 、英語で送付する場合は e を入力して Enter 押下してください。(j / e)"
$Prompt2 = Read-Host "※ この操作をキャンセルする場合はその他の文字を押下してください。"

# ユーザーの入力結果から、処理を分岐
switch ($Prompt2) {
    "j" {
        # j の場合の処理 (日本語招待メール)

        $InvitedUserMessageInfo2 = @{
        CustomizedMessageBody = "Windows Developer Center/Partner Centerをご利用の皆様へ

        いつもお世話になっております。
        スクウェア・エニックス 情報システム部です。

        弊社の Windows Developer Center/Partner Center を使用するにあたり、
        こちらのメールの下にある、[招待の承諾] をクリックし、二要素認証の設定とゲストユーザー登録を進めて下さい。
        登録完了後、Windows Developer Center/Partner Center にサインインできない際は、弊社の社員経由でご連絡ください。

        よろしくお願いいたします。
        "
        }

        New-MgInvitation -InvitedUserDisplayName $WDCGuestDisplayName -InvitedUserEmailAddress $WDCGuestMail `
        -InviteRedirectUrl "https://account.activedirectory.windowsazure.com/?tenantid=0e371789-fac2-4e0e-b7ac-30c4834d6b4e" `
        -InvitedUserMessageInfo $InvitedUserMessageInfo2 `
        -SendInvitationMessage:$true

        # CreationType が Invitation になるまでループ
        do {
        $status = Get-MgUser -Filter "Mail eq '$WDCGuestMail'" `
        -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName,Id
        }
        until ($status.CreationType -eq "Invitation")

        # SPOリストアイテム登録 - PowerAutomate で 外注WDCグループに登録させる目的
        $SiteId = '773c66a7-0841-40f8-b397-a38905e05ae8' # o365project
        $ListId = '66473670-63B5-4C8D-B5CF-AA94A992D85C' # WDCGuestList
        $operatorMailAddress = $env:USERNAME + "@square-enix.com" # 作業実行者のメールアドレス
        $body = @{
        fields = @{
            Title = $status.Mail
            AADUserId = $status.Id
            Operator = $operatorMailAddress
            }
        }

        Invoke-MgGraphRequest -Method POST -Uri "/v1.0/sites/$SiteId/lists/$ListId/items" -Body $body -ContentType "application/json; chrset=utf-8"

        Write-Host ""
        Write-Host ""
        Write-Host " 新規ゲストユーザーの登録が以下の通り完了しました。招待メールも送付済です。" -ForegroundColor Green

        $status | fl DisplayName, Mail, UserPrincipalName, CreatedDateTime, UserType, ExternalUserState

        Write-Host " この後、自動処理でゲストユーザーを「外注WDCユーザー」グループに参加させます。" -ForegroundColor Yellow
        Write-Host " 処理が完了するとメールで通知されます。" -ForegroundColor Yellow
        Write-Host " ※ 15 分以上経過しても完了通知が届かない場合は、GUI で Azure AD 管理センターから状態確認し、手動で手順を完了させてください。" -ForegroundColor Yellow

        }

    "e" {
        # e の場合の処理 (英語招待メール)

        $InvitedUserMessageInfo2e = @{
        CustomizedMessageBody = "To Windows Developer Center/Partner Center Users.

        This is Office 365 Admin from SEJ Information Technology Division.

        Please click the following link, and complete to setup 2 factor authentication and register as a guest user.
        After registered, please check you can sign-in to Windows Developer Center/Partner Center with your email address.
        If you can’t access it, please let SEJ staffs know.
        
        Regards,
        "
        }

        New-MgInvitation -InvitedUserDisplayName $WDCGuestDisplayName -InvitedUserEmailAddress $WDCGuestMail `
        -InviteRedirectUrl "https://account.activedirectory.windowsazure.com/?tenantid=0e371789-fac2-4e0e-b7ac-30c4834d6b4e" `
        -InvitedUserMessageInfo $InvitedUserMessageInfo2e `
        -SendInvitationMessage:$true

        # CreationType が Invitation になるまでループ
        do {
        $status = Get-MgUser -Filter "Mail eq '$WDCGuestMail'" `
        -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName, Id
        }
        until ($status.CreationType -eq "Invitation")

        # SPOリストアイテム登録 - PowerAutomate で 外注WDCグループに登録させる目的
        $SiteId = '773c66a7-0841-40f8-b397-a38905e05ae8' # o365project
        $ListId = '66473670-63B5-4C8D-B5CF-AA94A992D85C' # WDCGuestList
        $operatorMailAddress = $env:USERNAME + "@square-enix.com" # 作業実行者のメールアドレス
        $body = @{
        fields = @{
            Title = $status.Mail
            AADUserId = $status.Id
            Operator = $operatorMailAddress
            }
        }

        Invoke-MgGraphRequest -Method POST -Uri "/v1.0/sites/$SiteId/lists/$ListId/items" -Body $body -ContentType "application/json; chrset=utf-8"

        Write-Host ""
        Write-Host ""
        Write-Host " 新規ゲストユーザーの登録が以下の通り完了しました。招待メールも送付済です。" -ForegroundColor Green

        $status | fl DisplayName, Mail, UserPrincipalName, CreatedDateTime, UserType, ExternalUserState

        Write-Host " この後、自動処理でゲストユーザーを「外注WDCユーザー」グループに参加させます。" -ForegroundColor Yellow
        Write-Host " 処理が完了するとメールで通知されます。" -ForegroundColor Yellow
        Write-Host " ※ 15 分以上経過しても完了通知が届かない場合は、GUI で Azure AD 管理センターから状態確認し、手動で手順を完了させてください。" -ForegroundColor Yellow

        }

    default {
            # j でも e でもない場合の処理
            Write "j でも e でもない文字が入力されました。処理を中止します"
            }
  }
  }
}

#####################################
#  3 AAD ゲストユーザー存在確認     #
#####################################
function Check-AADUser {
    Disp_Title

    # ゲストのメールアドレス入力要求。if チェックでも使用
    Write-Host "スクエニ組織の Azure AD に存在するかを確認したいメールアドレスを入力してください。"
    $verifyMail = Read-Host "末尾に改行やスペースが入らないように注意してください。"

    # if でチェックするために変数格納
    $AADUserCheck = Get-MgUser -Filter "Mail eq '$verifyMail'" `
    -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName

    # アドレスがすでに AAD に存在しているかをチェック。SEJ 内部ユーザーはオミットする。
    If ($AADUserCheck.UserType -eq "Member"){

      Write-Host ""
      Write-Host "　このメールアドレスのユーザーは下記の通り、スクエニ組織の内部ユーザーです。" -ForegroundColor Yellow
      Write-Host "　※下記に結果が返らない場合は SEE の可能性がございます。" -ForegroundColor Yellow

      $AADUserCheck | fl DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType

    }ElseIf($AADUserCheck.UserType -eq "Guest"){

      Write-Host ""
      Write-Host "このメールアドレスのユーザーはスクエニ組織の Azure AD に既にゲストとして存在します。(下記参照) " -ForegroundColor Yellow

      $AADUserCheck | fl DisplayName, Mail, UserPrincipalName, CreatedDateTime, UserType, ExternalUserState, ExternalUserStateChangeDateTime

    }Else{

      Write-Host "このメールアドレスのユーザーはスクエニ組織の Azure AD には存在しません。 " -ForegroundColor Yellow
  }
}

########################################
#  4 スクエニ全ゲストユーザー表示      #
########################################
function Get-AADGuestUser
{
    Disp_Title
    Write-Output " スクエニ組織の Azure AD に登録されている全ゲストユーザを表示します。"
    Write-Output " ウィンドウが開くまでお待ちください..."

    Get-MgUser -Filter "UserType eq 'Guest'" -All `
    -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName |`
    Select Mail, DisplayName, UserPrincipalName, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime | sort -Property CreatedDateTime -Descending | `
    ogv -Title "スクエニ AAD ゲストユーザー (閲覧専用)"
}

#################################################
#  5「外注WDCユーザー」グループの全メンバー表示 #
#################################################
function Get-WDCGroupMember
{
    Disp_Title
    Write-Output " 「外注WDCユーザー」グループの全メンバーを表示します。"
    Write-Output " ウィンドウが開くまでお待ちください..."

    Get-MgGroupMember -GroupId 812e9035-5b45-49ea-b9e4-6f49224599f7 -All | % {@{UserId=$_.Id}} |`
    Get-MgUser -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName |`
    select Mail, DisplayName, UserPrincipalName, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime | sort -Property CreatedDateTime -Descending | `
    ogv -Title "「外注WDCユーザー」グループのメンバー (閲覧専用)"
}

########################################
#  6「外注WDCユーザー」全員を CSV 出力 #
########################################
function Export-WDCGroupMemberCsv
{
    Disp_Title
    Write-Output " 「外注WDCユーザー」グループの全メンバーを CSV で出力します。"

    $Date = Get-Date -Format "yyyy-MMdd-HHmmss"
    Get-MgGroupMember -GroupId 812e9035-5b45-49ea-b9e4-6f49224599f7 -All | % {@{UserId=$_.Id}} |`
    Get-MgUser -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName |`
    select Mail, DisplayName, UserPrincipalName, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime | sort -Property CreatedDateTime -Descending | `
    epcsv -NotypeInformation -Encoding utf8 -Path "$env:USERPROFILE\Desktop\WDCGroupMember_$Date.csv"

    Write-Host ""
    Write-Host "デスクトップに WDCGroupMember_$Date.csv を出力しました。" -ForegroundColor Green
    Write-Host ""
}

#################################################
#  7 Azure AD ゲストに招待状再送 (Teams 等)     #
#################################################
function ResendAADInvitation {
    Disp_Title

    # ゲストのメールアドレス入力要求。if チェックでも使用
    Write-Host "Azure AD 招待状を再送するゲストのメールアドレスを入力してください。(こちらは WDC ゲスト再送ではありません。)" -ForegroundColor Yellow
    $ResendAADGuest = Read-Host "末尾に改行やスペースが入らないように注意して入力してください。"

    # if でチェックするために変数格納
    $AADGuestPendingCheck = Get-MgUser -Filter "Mail eq '$ResendAADGuest'" `
    -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName

    # アドレスがすでに AAD に存在しているかをチェック。存在しない場合のみゲスト登録に進む。
    If ($AADGuestPendingCheck.UserType -eq "Member") {
      
      Write-Host ""
      Write-Host "　このメールアドレスのユーザーは下記の通り、スクエニ組織の内部ユーザーです。本作業は不要です。" -ForegroundColor Yellow
      Write-Host "　※下記に結果が返らない場合は SEE の可能性がございます。" -ForegroundColor Yellow

      $AADGuestPendingCheck | fl DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType

      }ElseIf($AADGuestPendingCheck.ExternalUserState -eq "Accepted"){

      Write-Host ""
      Write-Host "  このユーザーは下記の通り、スクエニ組織の Azure AD ゲスト招待状を受諾済です。本作業は不要です。" -ForegroundColor Yellow

      $AADGuestPendingCheck | fl Mail, DisplayName, UserType, CreatedDateTime, ExternalUserState 

      }ElseIf([string]::IsNullOrEmpty($AADGuestPendingCheck)){

      Write-Host ""
      Write-Host "  このユーザーはスクエニ組織の Azure AD には存在しません。" -ForegroundColor Yellow
      Write-Host "  過去に登録したゲストの場合、招待期限切れ等の理由で既に削除されています。" -ForegroundColor Yellow

      }Elseif($AADGuestPendingCheck.ExternalUserState -eq "PendingAcceptance"){

      # 招待メールの言語選択
      Write-Host ""
      Write-Host "   再招待の対象となるゲストユーザーです。" -ForegroundColor Yellow
      Write-Host "   再招待メールの言語を選択します。"  -ForegroundColor Yellow
      Write-Host ""
      Write-Host ""
      Write-Host "日本語で再送付する場合は j 、英語で再送付する場合は e を入力して Enter 押下してください。(j / e)"
      $Prompt7 = Read-Host "※ この操作をキャンセルする場合はその他の文字を押下してください。"

      # ユーザーの入力結果から、処理を分岐
      switch ($Prompt7) {

          "j" {
           # j の場合の処理 (日本語招待メール)

           $InvitedUserMessageInfo7 = @{

           CustomizedMessageBody = "※※※ 本招待メールは再送です。※※※
           
           スクウェア・エニックス組織のゲストの皆様へ

           いつもお世話になっております。
           スクウェア・エニックス 情報システム部です。

           本メールは、弊社社員経由でスクウェア・エニックスの O365 リソース (Teams 等) 利用依頼があった方にお送りしております。
           弊社のリソースを利用開始使用するにあたり、本メールの下にある [招待の承諾] をクリックし、
           二要素認証の設定とゲストユーザー登録を進めて下さい。

           登録に関して、不明点がございましたら弊社の社員経由でご連絡ください。
           ※本メール受領後 10 日間承諾されない場合は招待が破棄されますのでご注意ください。
           "
           }

           New-MgInvitation -InvitedUserEmailAddress $ResendAADGuest `
           -InviteRedirectUrl "https://account.activedirectory.windowsazure.com/?tenantid=0e371789-fac2-4e0e-b7ac-30c4834d6b4e" `
           -InvitedUserMessageInfo $InvitedUserMessageInfo7 `
           -SendInvitationMessage:$true

           # CreationType が Invitation になるまでループ
           do {
           $status = Get-MgUser -Filter "Mail eq '$ResendAADGuest'" `
           -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName
           }
           until ($status.CreationType -eq "Invitation")

           Write-Host ""
           Write-Host ""
           Write-Host " 以下のゲストユーザーに日本語で招待状を再送しました。" -ForegroundColor Green

           $status | fl Mail, DisplayName, UserType, CreatedDateTime, ExternalUserState, UserPrincipalName
           }

           "e" {
           # e の場合の処理 (英語招待メール)

           $InvitedUserMessageInfo7e = @{
           CustomizedMessageBody = "***** THIS EMAIL IS RESEND *****
           
           To SEJ Guest Users.

           This is Office 365 Admin from SQUARE-ENIX Japan (SEJ) IT Division.

           Please click the following link, and complete to setup 2 factor authentication and register as a guest user.
           After registered, please check you can sign-in to SEJ's Office 365 resouces with your email address.
           If you can’t access it, please let SEJ staffs know.

           Regards,
           "
           }

           New-MgInvitation -InvitedUserEmailAddress $ResendAADGuest `
           -InviteRedirectUrl "https://account.activedirectory.windowsazure.com/?tenantid=0e371789-fac2-4e0e-b7ac-30c4834d6b4e" `
           -InvitedUserMessageInfo $InvitedUserMessageInfo7e `
           -SendInvitationMessage:$true

           # CreationType が Invitation になるまでループ
           do {
           $status = Get-MgUser -Filter "Mail eq '$ResendAADGuest'" `
           -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName
           }
           until ($status.CreationType -eq "Invitation")

           Write-Host ""
           Write-Host ""
           Write-Host " 以下のゲストユーザーに英文で招待状を再送しました。" -ForegroundColor Green
           $status | fl Mail, DisplayName, UserType, CreatedDateTime, ExternalUserState, UserPrincipalName
           }

           default {
           # j でも e でもない場合の処理
            Write "j でも e でもない文字が入力されました。処理を中止します"
            }
                        }
                                                                               }
      }

#################################################
#  8 Azure AD ゲストに招待状再送 (WDC 外注)     #
#################################################
function ResendWDCInvitation {
    Disp_Title

    # ゲストのメールアドレス入力要求。if チェックでも使用
    Write-Host "WDC の招待状を再送するゲストのメールアドレスを入力してください。" -ForegroundColor Yellow
    $ResendWDCGuest = Read-Host "末尾に改行やスペースが入らないように注意して入力してください。"

    # if でチェックするために変数格納
    $WDCGuestPendingCheck = Get-MgUser -Filter "Mail eq '$ResendWDCGuest'" `
    -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName

    # アドレスがすでに AAD に存在しているかをチェック。存在しない場合のみゲスト登録に進む。
    If ($WDCGuestPendingCheck.UserType -eq "Member") {
      
      Write-Host ""
      Write-Host "　このメールアドレスのユーザーは下記の通り、スクエニ組織の内部ユーザーです。本作業は不要です。" -ForegroundColor Yellow
      Write-Host "　※下記に結果が返らない場合は SEE の可能性がございます。" -ForegroundColor Yellow

      $WDCGuestPendingCheck | fl DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType

      }ElseIf($WDCGuestPendingCheck.ExternalUserState -eq "Accepted"){

      Write-Host ""
      Write-Host "  このユーザーは下記の通り、スクエニ組織の Azure AD ゲスト招待状を受諾済です。本作業は不要です。" -ForegroundColor Yellow

      $WDCGuestPendingCheck | fl Mail, DisplayName, UserType, CreatedDateTime, ExternalUserState

      }ElseIf([string]::IsNullOrEmpty($WDCGuestPendingCheck)){

      Write-Host ""
      Write-Host "  このユーザーはスクエニ組織の Azure AD には存在しません。" -ForegroundColor Yellow
      Write-Host "  過去に登録したゲストの場合、招待期限切れ等の理由で既に削除されています。" -ForegroundColor Yellow

      }Elseif($WDCGuestPendingCheck.ExternalUserState -eq "PendingAcceptance"){

      # 招待メールの言語選択
      Write-Host ""
      Write-Host "   再招待の対象となるゲストユーザーです。" -ForegroundColor Yellow
      Write-Host "   再招待メールの言語を選択します。"  -ForegroundColor Yellow
      Write-Host ""
      Write-Host "日本語で再送付する場合は j 、英語で再送付する場合は e を入力して Enter 押下してください。(j / e)"
      $Prompt8 = Read-Host "※ この操作をキャンセルする場合はその他の文字を押下してください。"

      # ユーザーの入力結果から、処理を分岐
      switch ($Prompt8) {

          "j" {
           # j の場合の処理 (日本語招待メール)

           $InvitedUserMessageInfo8 = @{

           CustomizedMessageBody = "※※※ 本招待メールは再送です。※※※
           
           Windows Developer Center/Partner Centerをご利用の皆様へ

           いつもお世話になっております。
           スクウェア・エニックス 情報システム部です。

           弊社の Windows Developer Center/Partner Center を使用するにあたり、
           本メール下部、[招待の承諾] をクリックし、二要素認証の設定とゲストユーザー登録を進めて下さい。
           登録完了後、Windows Developer Center/Partner Center にサインインできない際は、弊社の社員経由でご連絡ください。

           よろしくお願いいたします。
           "
           }

           New-MgInvitation -InvitedUserEmailAddress $ResendWDCGuest `
           -InviteRedirectUrl "https://account.activedirectory.windowsazure.com/?tenantid=0e371789-fac2-4e0e-b7ac-30c4834d6b4e" `
           -InvitedUserMessageInfo $InvitedUserMessageInfo8 `
           -SendInvitationMessage:$true

           # CreationType が Invitation になるまでループ
           do {
           $status = Get-MgUser -Filter "Mail eq '$ResendWDCGuest'" `
           -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName
           }
           until ($status.CreationType -eq "Invitation")

           Write-Host ""
           Write-Host ""
           Write-Host " 以下のゲストユーザーに日本語で招待状を再送しました。" -ForegroundColor Green
           $status | fl Mail, DisplayName, UserType, CreatedDateTime, ExternalUserState
           }

           "e" {
           # e の場合の処理 (英語招待メール)

           $InvitedUserMessageInfo8e = @{
           CustomizedMessageBody = "***** THIS EMAIL IS RESEND *****
           
           To Windows Developer Center/Partner Center Users.

           This is Office 365 Admin from SEJ IT Division.

           Please click the following link, 
           and complete to setup 2 factor authentication and register as a guest user.
           After registered, please check you can sign-in to Windows Developer Center/Partner Center with your email address.
           If you can’t access it, please let SEJ staffs know.

           Regards,
           "
           }

           New-MgInvitation -InvitedUserEmailAddress $ResendWDCGuest `
           -InviteRedirectUrl "https://account.activedirectory.windowsazure.com/?tenantid=0e371789-fac2-4e0e-b7ac-30c4834d6b4e" `
           -InvitedUserMessageInfo $InvitedUserMessageInfo8e `
           -SendInvitationMessage:$true

           # CreationType が Invitation になるまでループ
           do {
           $status = Get-MgUser -Filter "Mail eq '$ResendWDCGuest'" `
           -Property DisplayName, Mail, BusinessPhones, CompanyName, Department, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName
           }
           until ($status.CreationType -eq "Invitation")

           Write-Host ""
           Write-Host ""
           Write-Host " 以下のゲストユーザーに英文で招待状を再送しました。" -ForegroundColor Green

           $status | fl Mail, DisplayName, UserType, CreatedDateTime, ExternalUserState

           }

           default {
           # j でも e でもない場合の処理
            Write "j でも e でもない文字が入力されました。処理を中止します"
            }
                        }
                                                                               }
      }

##########################################
#  9 スクエニ全ゲストユーザー CSV 出力   #
##########################################
function Export-AADGuestUserCsv
{
    Disp_Title
    Write-Output " スクエニ組織の Azure AD に登録されている全ゲストユーザを CSV で出力します。"

    # CSV 出力先パス指定
    $Date = Get-Date -Format "yyyy-MMdd-HHmmss"

    Get-MgUser -Filter "UserType eq 'Guest'" -All `
    -Property DisplayName, Mail, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime, UserPrincipalName | `
    Select Mail, DisplayName, UserPrincipalName, UserType, CreationType, CreatedDateTime, ExternalUserState, ExternalUserStateChangeDateTime | sort -Property CreatedDateTime -Descending | `
    epcsv -NotypeInformation -Encoding utf8 -Path "$env:USERPROFILE\Desktop\AADGuestUsers_$Date.csv"

    Write-Host ""
    Write-Host "デスクトップに AADGuestUsers_$Date.csv を出力しました。" -ForegroundColor Green
    Write-Host ""
}

#------------------------------------------------------------------- Main Section Start --------------------------------------------------------------------------------------------

. StartProc
#---------------------------
#  Graph 接続
#---------------------------
Write-Host ""
Write-Host "   guestuserinvitation@square-enix.com で接続します。" -ForegroundColor Yellow
Write-Host ""
Write-Host "   偽名アカウントで接続すると、ゲスト招待状を偽名メールアドレスから送信してしまうため要注意。" -ForegroundColor Cyan

. ConnectGraph
if ($rtn -eq $false){

    Write-Output "処理を終了します。何かキーを押してください..."
    $host.UI.RawUI.ReadKey()
    . EndProc
    exit
}
else
{
    . ImportGraph
    . main
}

#------------------------------------------------------------------- Footer Section Start --------------------------------------------------------------------------------------------

<#

バージョン版数 X.Y.X はそれぞれ以下のルールとします。

　X = メジャーバージョン ... マイナーバージョン 9 の次は繰り上がる。他の繰上り条件は現時点で明確な定義無し。
　Y = マイナーバージョン ... function 追加など
　Z = 不具合修正・既存 function の機能修正で版数追加。メジャー or マイナーバージョンが上がれば 0 にリセットする。

  更新日付   : 更新者    更新内容
  ---------- : ------    --------
  2022/08/16 : 土井      初版 1.0.0 ゲストユーザーを登録する用途で作成
  2022/08/23 : 土井      更新 1.1.0 接続時の Scope と Import-Module を実行最低限にすることで接続高速化を図る

【欲しい機能】
・動作の軽量化 … 接続時の scope と Import-Module を最小限にすればより高速化する見込み

#>