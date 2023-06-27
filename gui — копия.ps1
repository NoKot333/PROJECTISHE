Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

$StartForm                    = New-Object system.Windows.Forms.Form

$StartForm.ClientSize         = '420,430'
$StartForm.text               = "AdminHelperCenter"
$StartForm.BackColor          = "#ffffff"

$BlockUserBtn = New-Object system.Windows.Forms.Button
$BlockUserBtn.BackColor         = "#DCDCDC"
$BlockUserBtn.text              = "Заблокировать пользователя"
$BlockUserBtn.width             = 400
$BlockUserBtn.height            = 30
$BlockUserBtn.location          = New-Object System.Drawing.Point(10,10)
$BlockUserBtn.Font              = 'Microsoft Sans Serif,10'
$BlockUserBtn.ForeColor         = "#000000"


$UnbanUserBtn = New-Object system.Windows.Forms.Button
$UnbanUserBtn.BackColor         = "#DCDCDC"
$UnbanUserBtn.text              = "Разблокировать пользователя"
$UnbanUserBtn.width             = 400
$UnbanUserBtn.height            = 30
$UnbanUserBtn.location          = New-Object System.Drawing.Point(10,45)
$UnbanUserBtn.Font              = 'Microsoft Sans Serif,10'
$UnbanUserBtn.ForeColor         = "#000000"

$ShowMailBoxBtn = New-Object system.Windows.Forms.Button
$ShowMailBoxBtn.BackColor         = "#DCDCDC"
$ShowMailBoxBtn.text              = "Показать почтовый ящик"
$ShowMailBoxBtn.width             = 400
$ShowMailBoxBtn.height            = 30
$ShowMailBoxBtn.location          = New-Object System.Drawing.Point(10,80)
$ShowMailBoxBtn.Font              = 'Microsoft Sans Serif,10'
$ShowMailBoxBtn.ForeColor         = "#000000"

$HideMailBoxBtn = New-Object system.Windows.Forms.Button
$HideMailBoxBtn.BackColor         = "#DCDCDC"
$HideMailBoxBtn.text              = "Скрыть почтовый ящик"
$HideMailBoxBtn.width             = 400
$HideMailBoxBtn.height            = 30
$HideMailBoxBtn.location          = New-Object System.Drawing.Point(10,115)
$HideMailBoxBtn.Font              = 'Microsoft Sans Serif,10'
$HideMailBoxBtn.ForeColor         = "#000000"

$ChangePasswBtn = New-Object system.Windows.Forms.Button
$ChangePasswBtn.BackColor         = "#DCDCDC"
$ChangePasswBtn.text              = "Сменить пароль"
$ChangePasswBtn.width             = 400
$ChangePasswBtn.height            = 30
$ChangePasswBtn.location          = New-Object System.Drawing.Point(10,150)
$ChangePasswBtn.Font              = 'Microsoft Sans Serif,10'
$ChangePasswBtn.ForeColor         = "#000000"

$AddToGroupBtn = New-Object system.Windows.Forms.Button
$AddToGroupBtn.BackColor         = "#DCDCDC"
$AddToGroupBtn.text              = "Добавить к группе"
$AddToGroupBtn.width             = 400
$AddToGroupBtn.height            = 30
$AddToGroupBtn.location          = New-Object System.Drawing.Point(10,185)
$AddToGroupBtn.Font              = 'Microsoft Sans Serif,10'
$AddToGroupBtn.ForeColor         = "#000000"

$RemoveFromGroupBtn = New-Object system.Windows.Forms.Button
$RemoveFromGroupBtn.BackColor         = "#DCDCDC"
$RemoveFromGroupBtn.text              = "Исключить из группы"
$RemoveFromGroupBtn.width             = 400
$RemoveFromGroupBtn.height            = 30
$RemoveFromGroupBtn.location          = New-Object System.Drawing.Point(10,220)
$RemoveFromGroupBtn.Font              = 'Microsoft Sans Serif,10'
$RemoveFromGroupBtn.ForeColor         = "#000000"

$GetUserInfoBtn = New-Object system.Windows.Forms.Button
$GetUserInfoBtn.BackColor         = "#DCDCDC"
$GetUserInfoBtn.text              = "Получить информацию о пользователе"
$GetUserInfoBtn.width             = 400
$GetUserInfoBtn.height            = 30
$GetUserInfoBtn.location          = New-Object System.Drawing.Point(10,255)
$GetUserInfoBtn.Font              = 'Microsoft Sans Serif,10'
$GetUserInfoBtn.ForeColor         = "#000000"

$GetGroupInfoBtn = New-Object system.Windows.Forms.Button
$GetGroupInfoBtn.BackColor         = "#DCDCDC"
$GetGroupInfoBtn.text              = "Получить информацию о группе"
$GetGroupInfoBtn.width             = 400
$GetGroupInfoBtn.height            = 30
$GetGroupInfoBtn.location          = New-Object System.Drawing.Point(10,290)
$GetGroupInfoBtn.Font              = 'Microsoft Sans Serif,10'
$GetGroupInfoBtn.ForeColor         = "#000000"

$CreateUserBtn = New-Object system.Windows.Forms.Button
$CreateUserBtn.BackColor         = "#DCDCDC"
$CreateUserBtn.text              = "Создать нового пользователя"
$CreateUserBtn.width             = 400
$CreateUserBtn.height            = 30
$CreateUserBtn.location          = New-Object System.Drawing.Point(10,325)
$CreateUserBtn.Font              = 'Microsoft Sans Serif,10'
$CreateUserBtn.ForeColor         = "#000000"

$ConnectToDBBtn = New-Object system.Windows.Forms.Button
$ConnectToDBBtn.BackColor         = "#DCDCDC"
$ConnectToDBBtn.text              = "Подключиться к Microsoft Exchange"
$ConnectToDBBtn.width             = 400
$ConnectToDBBtn.height            = 30
$ConnectToDBBtn.location          = New-Object System.Drawing.Point(10,360)
$ConnectToDBBtn.Font              = 'Microsoft Sans Serif,10'
$ConnectToDBBtn.ForeColor         = "#000000"

$AddToDBBtn = New-Object system.Windows.Forms.Button
$AddToDBBtn.BackColor         = "#DCDCDC"
$AddToDBBtn.text              = "Добавить пользователя к почтовой БД"
$AddToDBBtn.width             = 400
$AddToDBBtn.height            = 30
$AddToDBBtn.location          = New-Object System.Drawing.Point(10,360)
$AddToDBBtn.Font              = 'Microsoft Sans Serif,10'
$AddToDBBtn.ForeColor         = "#000000"

$DisconnectFromDBBtn = New-Object system.Windows.Forms.Button
$DisconnectFromDBBtn.BackColor         = "#DCDCDC"
$DisconnectFromDBBtn.text              = "Отключиться от Microsoft Exchange"
$DisconnectFromDBBtn.width             = 400
$DisconnectFromDBBtn.height            = 30
$DisconnectFromDBBtn.location          = New-Object System.Drawing.Point(10,395)
$DisconnectFromDBBtn.Font              = 'Microsoft Sans Serif,10'
$DisconnectFromDBBtn.ForeColor         = "#000000"

function BlockUserFunc { 
    function CheckUser     {
        $userInfo = Get-ADUser -identity $UserLogin.Text -ErrorAction Ignore
            if ( $null -ne $userInfo ) {
                if ( ($userInfo  | Select-Object -expand Enabled) -eq $true) {
                    if([System.Windows.MessageBox]::Show('Вы уверены, что хотите заблокировать пользователя','Подтверждение', 'YesNo','Question') -eq 'Yes'){
                        try {Disable-ADAccount -identity $UserLogin.Text
                            [System.Windows.MessageBox]::Show("Пользователь заблокирован","Готово","OK","Information")}
                        catch {[System.Windows.MessageBox]::Show("Произошла неизвестная ошибка","Ошибка","OK","Error")} }
                    else {[System.Windows.MessageBox]::Show("Блокировка отменена","Отмена","OK","Information")}
            } else {
            [System.Windows.MessageBox]::Show("Этот пользователь уже заблокирован","Ошибка","OK","Error")
            }
      } 
            else {
        [System.Windows.MessageBox]::Show("Невозможно найти пользователя $(UserLogin.Text)","Ошибка","OK","Error")
           }
    }
    $CurrForm                    = New-Object system.Windows.Forms.Form
    $CurrForm.ClientSize         = '500,300'
    $CurrForm.text               = "Блокировка пользователя"
    
    $LabelL                           = New-Object system.Windows.Forms.Label
    $LabelL.text                      = "Введите логин"
    $LabelL.AutoSize                  = $true
    $LabelL.width                     = 25
    $LabelL.height                    = 10
    $LabelL.location                  = New-Object System.Drawing.Point(10,10)
    $LabelL.Font                      = 'Microsoft Sans Serif,13'
    
    $UserLogin = New-Object System.Windows.Forms.ComboBox
    $users = Get-ADUser -Filter * | Select -ExpandProperty "Name"
    $UserLogin.Items.AddRange($users)
    $UserLogin.Width = 400
    $UserLogin.Height = 30
    $UserLogin.location = New-Object System.Drawing.Point(10,45)
    $UserLogin.Font = 'Microsoft Sans Serif,10'

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.BackColor         = "#DCDCDC"
    $Submit.text              = "Заблокировать пользователя"
    $Submit.width             = 300
    $Submit.height            = 30
    $Submit.location          = New-Object System.Drawing.Point(10,80)
    $Submit.Font              = 'Microsoft Sans Serif,10'
    $Submit.ForeColor         = "#000000"

    $Submit.Add_Click({ CheckUser })

    $CurrForm.Controls.AddRange(@($LabelL,$UserLogin,$Submit))

    [void]$CurrForm.ShowDialog()
}
function UnBanUserFunc { 
    function CheckUser     {
        $userInfo = Get-ADUser -identity $UserLogin.Text -ErrorAction SilentlyContinue
            if ( $null -ne $userInfo ) {
                if ( ($userInfo  | Select-Object -expand Enabled) -eq $false) {
                    if([System.Windows.MessageBox]::Show('Точно?', 'Уверены?', 'YesNo','Question') -eq 'Yes'){
                        try {Enable-ADAccount -identity $UserLogin.Text
                            [System.Windows.MessageBox]::Show("Пользователь разблокирован","Готово","OK","Information")}
                        catch {[System.Windows.MessageBox]::Show("Произошла неизвестная ошибка","Ошибка","OK","Error")} }
                    else {[System.Windows.MessageBox]::Show("Разблокировка отменена","Отмена","OK","Information")}
            } else {
            [System.Windows.MessageBox]::Show("Этот пользователь уже разблокирован","Ошибка","OK","Error")
            }
      } 
            else {
        [System.Windows.MessageBox]::Show("Невозможно найти пользователя $($UserLogin.Text)","Ошибка","OK","Error")
           }
    }
    $CurrForm                    = New-Object system.Windows.Forms.Form
    $CurrForm.ClientSize         = '500,300'
    $CurrForm.text               = "Разблокировка пользователя"

    $LabelL                           = New-Object system.Windows.Forms.Label
    $LabelL.text                      = "Введите логин"
    $LabelL.AutoSize                  = $true
    $LabelL.width                     = 25
    $LabelL.height                    = 10
    $LabelL.location                  = New-Object System.Drawing.Point(10,10)
    $LabelL.Font                      = 'Microsoft Sans Serif,13'

    $UserLogin = New-Object System.Windows.Forms.ComboBox
    $users = Get-ADUser -Filter * | Select -ExpandProperty "Name"
    $UserLogin.Items.AddRange($users)
    $UserLogin.Width = 400
    $UserLogin.Height = 30
    $UserLogin.location = New-Object System.Drawing.Point(10,45)
    $UserLogin.Font = 'Microsoft Sans Serif,10'

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.BackColor         = "#DCDCDC"
    $Submit.text              = "Разблокировать пользователя"
    $Submit.width             = 300
    $Submit.height            = 30
    $Submit.location          = New-Object System.Drawing.Point(10,80)
    $Submit.Font              = 'Microsoft Sans Serif,10'
    $Submit.ForeColor         = "#000000"

    $Submit.Add_Click({ CheckUser })

    $CurrForm.Controls.AddRange(@($LabelL,$UserLogin,$Submit))

    [void]$CurrForm.ShowDialog()
}
function ShowMailBoxFunc { 
    function CheckUser  
   {
        $userInfo = Get-ADUser -identity $UserLogin.Text -ErrorAction Ignore
        if ($null -ne $userInfo ) {
            if([System.Windows.MessageBox]::Show('Точно?', 'Уверены?', 'YesNo','Question') -eq 'Yes'){
                Set-ADUser -Identity $UserLogin.Text -replace @{msExchHideFromAddressLists=$false}
                [System.Windows.MessageBox]::Show("Ящик показывается в списке адресов","Готово","OK","Information")
                } else {
                [System.Windows.MessageBox]::Show("Произошла отмена","Отмена","OK","Information")
      }} else {
        [System.Windows.MessageBox]::Show("Пользователь не найден","Ошибка","OK","Error")
        }
    }
    $CurrForm                    = New-Object system.Windows.Forms.Form
    $CurrForm.ClientSize         = '500,300'
    $CurrForm.text               = "Раскрыть почтовый ящик пользователя"

    $LabelL                           = New-Object system.Windows.Forms.Label
    $LabelL.text                      = "Введите логин пользователя"
    $LabelL.AutoSize                  = $true
    $LabelL.width                     = 25
    $LabelL.height                    = 10
    $LabelL.location                  = New-Object System.Drawing.Point(10,10)
    $LabelL.Font                      = 'Microsoft Sans Serif,13'

    $UserLogin = New-Object System.Windows.Forms.ComboBox
    $users = Get-ADUser -Filter * | Select -ExpandProperty "Name"
    $UserLogin.Items.AddRange($users)
    $UserLogin.Width = 400
    $UserLogin.Height = 30
    $UserLogin.location = New-Object System.Drawing.Point(10,45)
    $UserLogin.Font = 'Microsoft Sans Serif,10'

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.BackColor         = "#DCDCDC"
    $Submit.text              = "Сделать ящик видимым"
    $Submit.width             = 300
    $Submit.height            = 30
    $Submit.location          = New-Object System.Drawing.Point(10,80)
    $Submit.Font              = 'Microsoft Sans Serif,10'
    $Submit.ForeColor         = "#000000"

    $Submit.Add_Click({ CheckUser })

    $CurrForm.Controls.AddRange(@($LabelL,$UserLogin,$Submit))

    [void]$CurrForm.ShowDialog()
}
function HideMailBoxFunc { 
    function CheckUser     {
        $userInfo = Get-ADUser -identity $UserLogin.Text -ErrorAction Ignore
        if ($null -ne $userInfo ) {
            if([System.Windows.MessageBox]::Show('Точно?', 'Уверены?', 'YesNo','Question') -eq 'Yes'){
                Set-ADUser -Identity $UserLogin.Text -replace @{msExchHideFromAddressLists=$true}
                [System.Windows.MessageBox]::Show("Ящик не показывается в списке адресов","Готово","OK","Information")
                } else {
                [System.Windows.MessageBox]::Show("Произошла отмена","Отмена","OK","Information")
      }} else {
        [System.Windows.MessageBox]::Show("Пользователь не найден","Ошибка","OK","Error")
        }
    }
    $CurrForm                    = New-Object system.Windows.Forms.Form
    $CurrForm.ClientSize         = '500,300'
    $CurrForm.text               = "Скрыть почтовый ящик пользователя"

    $LabelL                           = New-Object system.Windows.Forms.Label
    $LabelL.text                      = "Введите логин"
    $LabelL.AutoSize                  = $true
    $LabelL.width                     = 25
    $LabelL.height                    = 10
    $LabelL.location                  = New-Object System.Drawing.Point(10,10)
    $LabelL.Font                      = 'Microsoft Sans Serif,13'

    $UserLogin = New-Object System.Windows.Forms.ComboBox
    $users = Get-ADUser -Filter * | Select -ExpandProperty "Name"
    $UserLogin.Items.AddRange($users)
    $UserLogin.Width = 400
    $UserLogin.Height = 30
    $UserLogin.location = New-Object System.Drawing.Point(10,10)
    $UserLogin.Font = 'Microsoft Sans Serif,10'

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.BackColor         = "#DCDCDC"
    $Submit.text              = "Скрыть ящик"
    $Submit.width             = 300
    $Submit.height            = 30
    $Submit.location          = New-Object System.Drawing.Point(10,45)
    $Submit.Font              = 'Microsoft Sans Serif,10'
    $Submit.ForeColor         = "#000000"

    $Submit.Add_Click({ CheckUser })

    $CurrForm.Controls.AddRange(@($LabelL,$UserLogin,$Submit))

    [void]$CurrForm.ShowDialog()
}
function ChangePasswFunc { 
    function CheckUser     {
    $userInfo = Get-ADUser -identity $UserLogin.Text -ErrorAction Ignore
        if ( $null -ne $userInfo ) {
            if ($FNewPassw.Text -eq $SNewPassw.Text) {
                if ([System.Windows.MessageBox]::Show('Точно?', 'Уверены?', 'YesNo','Question') -eq 'Yes'){
                    try {@( Set-ADAccountPassword $UserLogin.Text -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $FNewPassw.Text -Force -Verbose) -ErrorAction SilentlyContinue | Set-ADuser -ChangePasswordAtLogon $True )
                        [System.Windows.MessageBox]::Show("Пароль пользователя обновлён","Готово","OK","Information")}
                    catch {[System.Windows.MessageBox]::Show("Произошла неизвестная ошибка","Ошибка","OK","Error")} }
                else {[System.Windows.MessageBox]::Show("Смена пароля отменена","Отмена","OK","Information")}
         } else {[System.Windows.MessageBox]::Show("Пароли не совпадают","Ошибка","OK","Error")}
      }
       else {
        [System.Windows.MessageBox]::Show("Пользователь не найден")
      }
    }
    $CurrForm                    = New-Object system.Windows.Forms.Form
    $CurrForm.ClientSize         = '500,300'
    $CurrForm.text               = "Сменить пароль пользователя"

    $LabelL                           = New-Object system.Windows.Forms.Label
    $LabelL.text                      = "Введите логин"
    $LabelL.AutoSize                  = $true
    $LabelL.width                     = 25
    $LabelL.height                    = 10
    $LabelL.location                  = New-Object System.Drawing.Point(10,10)
    $LabelL.Font                      = 'Microsoft Sans Serif,13'
    
    $UserLogin = New-Object System.Windows.Forms.ComboBox
    $users = Get-ADUser -Filter * | Select -ExpandProperty "Name"
    $UserLogin.Items.AddRange($users)
    $UserLogin.Width = 400
    $UserLogin.Height = 30
    $UserLogin.location = New-Object System.Drawing.Point(10,45)
    $UserLogin.Font = 'Microsoft Sans Serif,10'

    $LabelFP                           = New-Object system.Windows.Forms.Label
    $LabelFP.text                      = "Введите пароль"
    $LabelFP.AutoSize                  = $true
    $LabelFP.width                     = 25
    $LabelFP.height                    = 10
    $LabelFP.location                  = New-Object System.Drawing.Point(10,80)
    $LabelFP.Font                      = 'Microsoft Sans Serif,13'

    $FNewPassw = New-Object System.Windows.Forms.TextBox
    $FNewPassw.Width = 400
    $FNewPassw.Height = 30
    $FNewPassw.location = New-Object System.Drawing.Point(10,115)
    $FNewPassw.Font = 'Microsoft Sans Serif,10'

    $LabelSP                           = New-Object system.Windows.Forms.Label
    $LabelSP.text                      = "Введите пароль повторно"
    $LabelSP.AutoSize                  = $true
    $LabelSP.width                     = 25
    $LabelSP.height                    = 10
    $LabelSP.location                  = New-Object System.Drawing.Point(10,150)
    $LabelSP.Font                      = 'Microsoft Sans Serif,13'

    $SNewPassw = New-Object System.Windows.Forms.TextBox
    $SNewPassw.Width = 400
    $SNewPassw.Height = 30
    $SNewPassw.location = New-Object System.Drawing.Point(10,185)
    $SNewPassw.Font = 'Microsoft Sans Serif,10'

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.BackColor         = "#DCDCDC"
    $Submit.text              = "Сменить пароль пользователя"
    $Submit.width             = 300
    $Submit.height            = 30
    $Submit.location          = New-Object System.Drawing.Point(10,220)
    $Submit.Font              = 'Microsoft Sans Serif,10'
    $Submit.ForeColor         = "#000000"

    $Submit.Add_Click({ CheckUser })

    $CurrForm.Controls.AddRange(@($LabelL,$UserLogin,$LabelFP,$FNewPassw,$LabelSP,$SNewPassw,$Submit))

    [void]$CurrForm.ShowDialog()
}
function AddToGroupFunc { 
    function CheckUser     {
    try {
            $groupInfo = Get-ADGroup -Identity $GroupLogin.Text -ErrorAction Ignore 
        } catch { $groupInfo = $null}
        if ($null -ne $groupInfo) {
        try {
            $userInfo = Get-ADUser -identity $UserLogin.Text -ErrorAction Ignore
        } catch { $userInfo = $null}
            if ( $null -ne $userInfo) {
                $members = Get-ADGroupMember -Identity $GroupLogin.Text -Recursive | Select-Object -ExpandProperty SamAccountName
                if ( $members -notcontains $UserLogin.Text) {

                    if([System.Windows.MessageBox]::Show('Точно?', 'Уверены?', 'YesNo','Question') -eq 'Yes')
                    {
                        try {Add-ADGroupMember -identity $GroupLogin.Text -Members $UserLogin.Text
                            [System.Windows.MessageBox]::Show("Пользователь  добавлен к группе")}
                        catch {[System.Windows.MessageBox]::Show("Произошла неизвестная ошибка")}
                    }
                    else {[System.Windows.MessageBox]::Show("Добавление пользователя было отменено")}
              } else {
                [System.Windows.MessageBox]::Show("Пользователь уже есть в группе")
              }
            } 
          else {
            [System.Windows.MessageBox]::Show("Пользователь не может быть найден")
           } 
         }
      else {
         [System.Windows.MessageBox]::Show("Группа не найдена")
       }
    }
    $CurrForm                    = New-Object system.Windows.Forms.Form
    $CurrForm.ClientSize         = '500,300'
    $CurrForm.text               = "Добавить пользователя в группу"

    $LabelL                           = New-Object system.Windows.Forms.Label
    $LabelL.text                      = "Введите логин пользователя"
    $LabelL.AutoSize                  = $true
    $LabelL.width                     = 25
    $LabelL.height                    = 10
    $LabelL.location                  = New-Object System.Drawing.Point(10,10)
    $LabelL.Font                      = 'Microsoft Sans Serif,13'

    $UserLogin = New-Object System.Windows.Forms.ComboBox
    $users = Get-ADUser -Filter * | Select -ExpandProperty "Name"
    $UserLogin.Items.AddRange($users)
    $UserLogin.Width = 400
    $UserLogin.Height = 30
    $UserLogin.location = New-Object System.Drawing.Point(10,45)
    $UserLogin.Font = 'Microsoft Sans Serif,10'

    $LabelG                           = New-Object system.Windows.Forms.Label
    $LabelG.text                      = "Введите логин группы"
    $LabelG.AutoSize                  = $true
    $LabelG.width                     = 25
    $LabelG.height                    = 10
    $LabelG.location                  = New-Object System.Drawing.Point(10,80)
    $LabelG.Font                      = 'Microsoft Sans Serif,13'

    $GroupLogin = New-Object System.Windows.Forms.ComboBox
    $groups = Get-ADGroup -Filter * | Select -ExpandProperty "Name"
    $GroupLogin.Items.AddRange($groups)
    $GroupLogin.Width = 400
    $GroupLogin.Height = 30
    $GroupLogin.location = New-Object System.Drawing.Point(10,115)
    $GroupLogin.Font = 'Microsoft Sans Serif,10'

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.BackColor         = "#DCDCDC"
    $Submit.text              = "Добавить"
    $Submit.width             = 300
    $Submit.height            = 30
    $Submit.location          = New-Object System.Drawing.Point(10,150)
    $Submit.Font              = 'Microsoft Sans Serif,10'
    $Submit.ForeColor         = "#000000"

    $Submit.Add_Click({ CheckUser })

    $CurrForm.Controls.AddRange(@($LabelL,$UserLogin,$LabelG,$GroupLogin,$Submit))

    [void]$CurrForm.ShowDialog()
}
function RemoveFromGroupFunc { 
    function CheckUser     {
    try {
            $groupInfo = Get-ADGroup -Identity $GroupLogin.Text -ErrorAction Ignore 
        } catch { $groupInfo = $null}
        if ($null -ne $groupInfo) {
        try {
            $userInfo = Get-ADUser -identity $UserLogin.Text -ErrorAction Ignore
        } catch { $userInfo = $null}
            if ( $null -ne $userInfo) {
                if ( $null -ne ((Get-ADUser $UserLogin.Text -Properties MemberOf).memberof -like "*$GroupLogin.Text*")) {

                    if([System.Windows.MessageBox]::Show('Точно?', 'Уверены?', 'YesNo','Question') -eq 'Yes')
                    {
                        try {Remove-ADGroupMember -identity $GroupLogin.Text -Members $UserLogin.Text -Confirm:$false
                            [System.Windows.MessageBox]::Show("Пользователь исключён из группы")}
                        catch {[System.Windows.MessageBox]::Show("Произошла неизвестная ошибка")}
                    }
                    else {[System.Windows.MessageBox]::Show("Исключение пользователя было отменено")}
              } else {
                [System.Windows.MessageBox]::Show("Пользователь нет в группе")
              }
            } 
          else {
            [System.Windows.MessageBox]::Show("Пользователь не может быть найден")
           } 
         }
      else {
         [System.Windows.MessageBox]::Show("Группа не найдена")
       }
    }
    $CurrForm                    = New-Object system.Windows.Forms.Form
    $CurrForm.ClientSize         = '500,300'
    $CurrForm.text               = "Исключить пользователя из группы"

    $LabelL                           = New-Object system.Windows.Forms.Label
    $LabelL.text                      = "Введите логин пользователя"
    $LabelL.AutoSize                  = $true
    $LabelL.width                     = 25
    $LabelL.height                    = 10
    $LabelL.location                  = New-Object System.Drawing.Point(10,10)
    $LabelL.Font                      = 'Microsoft Sans Serif,13'

    $UserLogin = New-Object System.Windows.Forms.ComboBox
    $users = Get-ADUser -Filter * | Select -ExpandProperty "Name"
    $UserLogin.Items.AddRange($users)
    $UserLogin.Width = 400
    $UserLogin.Height = 30
    $UserLogin.location = New-Object System.Drawing.Point(10,45)
    $UserLogin.Font = 'Microsoft Sans Serif,10'

    $LabelG                           = New-Object system.Windows.Forms.Label
    $LabelG.text                      = "Введите логин группы"
    $LabelG.AutoSize                  = $true
    $LabelG.width                     = 25
    $LabelG.height                    = 10
    $LabelG.location                  = New-Object System.Drawing.Point(10,80)
    $LabelG.Font                      = 'Microsoft Sans Serif,13'

    $GroupLogin = New-Object System.Windows.Forms.ComboBox
    $groups = Get-ADGroup -Filter * | Select -ExpandProperty "Name"
    $GroupLogin.Items.AddRange($groups)
    $GroupLogin.Width = 400
    $GroupLogin.Height = 30
    $GroupLogin.location = New-Object System.Drawing.Point(10,115)
    $GroupLogin.Font = 'Microsoft Sans Serif,10'

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.BackColor         = "#DCDCDC"
    $Submit.text              = "Исключить"
    $Submit.width             = 300
    $Submit.height            = 30
    $Submit.location          = New-Object System.Drawing.Point(10,150)
    $Submit.Font              = 'Microsoft Sans Serif,10'
    $Submit.ForeColor         = "#000000"

    $Submit.Add_Click({ CheckUser })

    $CurrForm.Controls.AddRange(@($LabelL,$UserLogin,$LabelG,$GroupLogin,$Submit))

    [void]$CurrForm.ShowDialog()
}
function GetUserInfoFunc { 
    function CheckUser     {
        $userInfo = Get-ADUser -identity $UserLogin.Text -ErrorAction Ignore
            if ( $null -ne $userInfo ) {
            $info = Get-ADUser -identity $UserLogin.Text -properties SamAccountName,EmailAddress,Enabled,PasswordExpired,PasswordLastSet,MemberOf | select-object DistinguishedName,SamAccountName,EmailAddress,Enabled,PasswordExpired,PasswordLastSet,MemberOf | Out-String
    [System.Windows.MessageBox]::Show($info)
  }
  else {
    [System.Windows.MessageBox]::Show("Невозможно найти пользователя")
       } 
    }
    $CurrForm                    = New-Object system.Windows.Forms.Form
    $CurrForm.ClientSize         = '500,300'
    $CurrForm.text               = "Получить информацию о пользователе"

    $LabelL                           = New-Object system.Windows.Forms.Label
    $LabelL.text                      = "Введите логин пользователя"
    $LabelL.AutoSize                  = $true
    $LabelL.width                     = 25
    $LabelL.height                    = 10
    $LabelL.location                  = New-Object System.Drawing.Point(10,10)
    $LabelL.Font                      = 'Microsoft Sans Serif,13'

    

    $UserLogin = New-Object System.Windows.Forms.ComboBox
    $users = Get-ADUser -Filter * | Select -ExpandProperty "Name"
    $UserLogin.Items.AddRange($users)
    $UserLogin.Width = 400
    $UserLogin.Height = 30
    $UserLogin.location = New-Object System.Drawing.Point(10,45)
    $UserLogin.Font = 'Microsoft Sans Serif,10'
    

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.BackColor         = "#DCDCDC"
    $Submit.text              = "Получить информацию"
    $Submit.width             = 300
    $Submit.height            = 30
    $Submit.location          = New-Object System.Drawing.Point(10,80)
    $Submit.Font              = 'Microsoft Sans Serif,10'
    $Submit.ForeColor         = "#000000"

    $Submit.Add_Click({ CheckUser })

    $CurrForm.Controls.AddRange(@($LabelL,$UserLogin,$Submit))

    [void]$CurrForm.ShowDialog()
}
function GetGroupInfoFunc { 
    function CheckUser     {
        $groupInfo = Get-ADGroup -identity $GroupLogin.Text -ErrorAction ignore
            if ( $null -ne $groupInfo ) {
            $info = "
            Name: $($groupInfo  | select -expand SamAccountName) 
            Container: $($groupInfo  | select -expand DistinguishedName)  
            Type: $($groupInfo  | select -expand GroupCategory)  
            Group Scope: $($groupInfo  | select -expand GroupScope)"
            $members = "
            Members: " + (Get-ADGroupMember $GroupLogin.Text | Select-Object Name| ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1)
            $info =  $info + $members
    [System.Windows.MessageBox]::Show($info)
  }
  else {
    [System.Windows.MessageBox]::Show("Невозможно найти группу")
       } 
    }
    $CurrForm                    = New-Object system.Windows.Forms.Form
    $CurrForm.ClientSize         = '500,300'
    $CurrForm.text               = "Получить информацию о пользователе"

    $LabelG                           = New-Object system.Windows.Forms.Label
    $LabelG.text                      = "Введите логин группы"
    $LabelG.AutoSize                  = $true
    $LabelG.width                     = 25
    $LabelG.height                    = 10
    $LabelG.location                  = New-Object System.Drawing.Point(10,10)
    $LabelG.Font                      = 'Microsoft Sans Serif,13'

    $GroupLogin = New-Object System.Windows.Forms.ComboBox
    $groups = Get-ADGroup -Filter * | Select -ExpandProperty "Name"
    $GroupLogin.Items.AddRange($groups)
    $GroupLogin.Width = 400
    $GroupLogin.Height = 30
    $GroupLogin.location = New-Object System.Drawing.Point(10,45)
    $GroupLogin.Font = 'Microsoft Sans Serif,10'
    

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.BackColor         = "#DCDCDC"
    $Submit.text              = "Получить информацию"
    $Submit.width             = 300
    $Submit.height            = 30
    $Submit.location          = New-Object System.Drawing.Point(10,80)
    $Submit.Font              = 'Microsoft Sans Serif,10'
    $Submit.ForeColor         = "#000000"

    $Submit.Add_Click({ CheckUser })

    $CurrForm.Controls.AddRange(@($LabelG,$GroupLogin,$Submit))

    [void]$CurrForm.ShowDialog()
}
function CreateUserFunc { 
    function CheckUser     {
        if (($FName.Text -match "[^А-Яа-яеЁ-]+") -or -not($FName.Text) ) {
            [System.Windows.MessageBox]::Show('Введите поле имя значение, содержащие буквы кириллицы и - ','Ошибка', 'OK','Error')
            return
        }
        if ($SName.Text -match "[^А-Яа-яеЁ-]+"-or -not($SName.Text)) {
            [System.Windows.MessageBox]::Show('Введите поле фамилия значение, содержащие буквы кириллицы и -','Ошибка', 'OK','Error')
            return
        }
        if ($MName.Text -match "[^А-Яа-яеЁ-]+"-or -not($MName.Text)) {
            [System.Windows.MessageBox]::Show('Введите поле отчество значение, содержащие буквы кириллицы и -','Ошибка', 'OK','Error')
            return
        }
        if ($CityBox.Text -eq "") {
            [System.Windows.MessageBox]::Show("Выберите город '$($FName.Text)'",'Ошибка', 'OK','Error')
            return
        }
        if ($CompanyBox.Text -eq "") {
            [System.Windows.MessageBox]::Show('Выберите компанию','Ошибка', 'OK','Error')
            return
        }
        if ($DepartmentsBox.Text -eq "") {
            [System.Windows.MessageBox]::Show('Выберите отдел','Ошибка', 'OK','Error')
            return
        }
        if ($AppointmentBox.Text -eq "") {
            [System.Windows.MessageBox]::Show('Выберите должность','Ошибка', 'OK','Error')
            return
        }
        $Hashcities = Get-Content "$PSScriptRoot\texts\cities.txt" -Raw | ConvertFrom-Json -AsHashtable
        $Hashcompanies = Get-Content ("$PSScriptRoot\texts\companies" + $Hashcities[$CityBox.Text]+".txt")-Raw | ConvertFrom-Json -AsHashtable
        $HashDepartments = Get-Content ("$PSScriptRoot\texts\departments" + $Hashcompanies[$CompanyBox.Text] +$Hashcities[$CityBox.Text]+ ".txt") -Raw | ConvertFrom-Json -AsHashtable
        $HashAppointment = Get-Content "$PSScriptRoot\texts\appointment.txt" -Raw | ConvertFrom-Json -AsHashtable
        $enFName = &"$PSScriptRoot\funcs\translit.ps1" $FName.Text
        $enSName = &"$PSScriptRoot\funcs\translit.ps1" $SName.Text
        $enMName = &"$PSScriptRoot\funcs\translit.ps1" $MName.Text

        $extAttr1=@{}
        $extAttr1.Add('Имя',$FName.Text)
        $extAttr1.Add('Фамилия',$SName.Text)
        $extAttr1.Add('Отчество',$MName.Text)
        $extAttr1.Add('Город', $CityBox.Text)
        $extAttr1.Add('Компания', $CompanyBox.Text)
        $extAttr1.Add('Отдел', $DepartmentsBox.Text)
        $extAttr1.Add('Должность',$AppointmentBox.Text)
        
        $baseUserName = $enFName +'.'+ $enSName
        while ($true) {    
            $userName = $baseUserName     
            if ($index -ne 0) {         
                $userName += $index     
            }     
            $user =  Get-ADUser -identity $userName -ErrorAction SilentlyContinue     
            if ( $null -eq $user ) {                  
                break     
            }     
            $index++ 
       }
        $enCompany = $Hashcompanies[$CompanyBox.Text]
        $enDepartment = $HashDepartments[$DepartmentsBox.Text]
        $enCity = $Hashcities[$CityBox.Text]
        $enAppointment = $HashAppointment[$AppointmentBox.Text]
        $UPName = $userName + '@' + (Get-ADDomain | Select -ExpandProperty DnsRoot)
        $message = "Создать пользователя?
        Пользователь: $userName
        Имя: $enFName
        Фамилия: $enSName
        Отчество: $enMName
        Компания: $enCompany
        Отдел: $enDepartment
        Город: $enCity
        Должность: $enAppointment
        UPName: $UPName

        Имя : $($extAttr1["Имя"])
        Фамилия : $($extAttr1["Фамилия"])
        Отчество : $($extAttr1["Отчество"])
        Город : $($extAttr1["Город"])
        Компания : $($extAttr1["Компания"])
        Отдел : $($extAttr1["Отдел"])
        Должность : $($extAttr1["Должность"])"
        if([System.Windows.MessageBox]::Show($message,'Подтверждение', 'YesNo','Question') -eq 'Yes'){
            try {New-ADUser -Name $userName -UserPrincipalName $UPName -Department $enDepartment -GivenName $enFName  -Surname $enSName -OtherName $enMName -SamAccountName $userName -City $enCity -Company $enCompany -title $enAppointment -Enabled $true
                      
            Set-ADAccountPassword $userName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText “PASSWORD” -Force -Verbose) | Set-ADuser -ChangePasswordAtLogon $True

            Set-ADuser -Identity $userName -Add @{extensionAttribute1 = $extAttr1["Имя"]} 
            Set-ADuser -Identity $userName -Add @{extensionAttribute2 = $extAttr1["Фамилия"]}
            Set-ADuser -Identity $userName -Add @{extensionAttribute3 = $extAttr1["Отчество"]}
            Set-ADuser -Identity $userName -Add @{extensionAttribute4 = $extAttr1["Город"]}
            Set-ADuser -Identity $userName -Add @{extensionAttribute4 = $extAttr1["Компания"]}
            Set-ADuser -Identity $userName -Add @{extensionAttribute5 = $extAttr1["Должность"]}
            Set-ADuser -Identity $userName -Add @{extensionAttribute6 = $extAttr1["Отдел"]}
            [System.Windows.MessageBox]::Show("Создание завершено успешно")
            } catch {
            [System.Windows.MessageBox]::Show("Неизвестная ошибка")
            }
      } else {
        [System.Windows.MessageBox]::Show("Отмена создания пользователя")        
      }
    }
    $CurrForm                    = New-Object system.Windows.Forms.Form
    $CurrForm.ClientSize         = '420,800'
    $CurrForm.text               = "Создать нового пользователя"

    $LabelF                           = New-Object system.Windows.Forms.Label
    $LabelF.text                      = "Введите Имя"
    $LabelF.AutoSize                  = $true
    $LabelF.width                     = 25
    $LabelF.height                    = 10
    $LabelF.location                  = New-Object System.Drawing.Point(10,10)
    $LabelF.Font                      = 'Microsoft Sans Serif,13'

    $FName = New-Object System.Windows.Forms.TextBox
    $FName.Width = 400
    $FName.Height = 30
    $FName.location = New-Object System.Drawing.Point(10,40)
    $FName.Font = 'Microsoft Sans Serif,10'

    $LabelS                           = New-Object system.Windows.Forms.Label
    $LabelS.text                      = "Введите Фамилия"
    $LabelS.AutoSize                  = $true
    $LabelS.width                     = 25
    $LabelS.height                    = 10
    $LabelS.location                  = New-Object System.Drawing.Point(10,75)
    $LabelS.Font                      = 'Microsoft Sans Serif,13'

    $SName = New-Object System.Windows.Forms.TextBox
    $SName.Width = 400
    $SName.Height = 30
    $SName.location = New-Object System.Drawing.Point(10,105)
    $SName.Font = 'Microsoft Sans Serif,10'

    $LabelM                           = New-Object system.Windows.Forms.Label
    $LabelM.text                      = "Введите Отчество"
    $LabelM.AutoSize                  = $true
    $LabelM.width                     = 25
    $LabelM.height                    = 10
    $LabelM.location                  = New-Object System.Drawing.Point(10,140)
    $LabelM.Font                      = 'Microsoft Sans Serif,13'

    $MName = New-Object System.Windows.Forms.TextBox
    $MName.Width = 400
    $MName.Height = 30
    $MName.location = New-Object System.Drawing.Point(10,170)
    $MName.Font = 'Microsoft Sans Serif,10'

    $LabelCity                           = New-Object system.Windows.Forms.Label
    $LabelCity.text                      = "Выберите город"
    $LabelCity.AutoSize                  = $true
    $LabelCity.width                     = 25
    $LabelCity.height                    = 10
    $LabelCity.location                  = New-Object System.Drawing.Point(10,205)
    $LabelCity.Font                      = 'Microsoft Sans Serif,13'

    $CityBox = New-Object System.Windows.Forms.ComboBox
    $CityBox.Width = 400
    $CityBox.Height = 30
    $CityBox.location = New-Object System.Drawing.Point(10,235)
    $CityBox.Font = 'Microsoft Sans Serif,10'

    $LabelCompany                           = New-Object system.Windows.Forms.Label
    $LabelCompany.text                      = "Выберите Компанию"
    $LabelCompany.AutoSize                  = $true
    $LabelCompany.width                     = 25
    $LabelCompany.height                    = 10
    $LabelCompany.location                  = New-Object System.Drawing.Point(10,270)
    $LabelCompany.Font                      = 'Microsoft Sans Serif,13'

    $CompanyBox = New-Object System.Windows.Forms.ComboBox
    $CompanyBox.Width = 400
    $CompanyBox.Height = 30
    $CompanyBox.location = New-Object System.Drawing.Point(10,300)
    $CompanyBox.Font = 'Microsoft Sans Serif,10'

    $LabelDepartments                           = New-Object system.Windows.Forms.Label
    $LabelDepartments.text                      = "Выберите отдел"
    $LabelDepartments.AutoSize                  = $true
    $LabelDepartments.width                     = 25
    $LabelDepartments.height                    = 10
    $LabelDepartments.location                  = New-Object System.Drawing.Point(10,335)
    $LabelDepartments.Font                      = 'Microsoft Sans Serif,13'

    $DepartmentsBox = New-Object System.Windows.Forms.ComboBox
    $DepartmentsBox.Width = 400
    $DepartmentsBox.Height = 30
    $DepartmentsBox.location = New-Object System.Drawing.Point(10,365)
    $DepartmentsBox.Font = 'Microsoft Sans Serif,10'

    $LabelAppointment                           = New-Object system.Windows.Forms.Label
    $LabelAppointment.text                      = "Выберите должность"
    $LabelAppointment.AutoSize                  = $true
    $LabelAppointment.width                     = 25
    $LabelAppointment.height                    = 10
    $LabelAppointment.location                  = New-Object System.Drawing.Point(10,400)
    $LabelAppointment.Font                      = 'Microsoft Sans Serif,13'

    $AppointmentBox = New-Object System.Windows.Forms.ComboBox
    $AppointmentBox.Width = 400
    $AppointmentBox.Height = 30
    $AppointmentBox.location = New-Object System.Drawing.Point(10,430)
    $AppointmentBox.Font = 'Microsoft Sans Serif,10'

    $HashAppointment = Get-Content "C:\Users\Администратор.WIN-HBSQBQLKH1H\Downloads\ПроектITакадемия\texts\appointment.txt" -Raw | ConvertFrom-Json -AsHashtable
    $Appointment = [string[]] $HashAppointment.Keys
    $enappointment = [string[]] $HashAppointment.Values
    $AppointmentBox.Items.AddRange($Appointment)

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.BackColor         = "#DCDCDC"
    $Submit.text              = "Создать пользователя"
    $Submit.width             = 300
    $Submit.height            = 30
    $Submit.location          = New-Object System.Drawing.Point(10,465)
    $Submit.Font              = 'Microsoft Sans Serif,10'
    $Submit.ForeColor         = "#000000"

    $Submit.Add_Click({ CheckUser })

    $Hashcities = Get-Content "C:\Users\Администратор.WIN-HBSQBQLKH1H\Downloads\ПроектITакадемия\texts\cities.txt" -Raw | ConvertFrom-Json -AsHashtable
    $cities = [string[]] $Hashcities.Keys
    $encites = [string[]] $Hashcities.Values
    $City_SelectedIndexChanged= {
        $CompanyBox.Items.Clear()
        $DepartmentsBox.Items.Clear() 
        $DepartmentsBox.Text = $null
        $CompanyBox.Text = $null
        $Hashcompanies = Get-Content ("C:\Users\Администратор.WIN-HBSQBQLKH1H\Downloads\ПроектITакадемия\texts\companies" + $Hashcities[$CityBox.Text]+".txt")-Raw | ConvertFrom-Json -AsHashtable
        $companies = [string[]] $Hashcompanies.Keys
        $encompanies = [string[]] $Hashcompanies.Values
        $CompanyBox.Items.AddRange($companies)
    }
    $CityBox.items.AddRange($cities)
    $CityBox.add_SelectedIndexChanged($City_SelectedIndexChanged)

    $Company_SelectedIndexChanged= {
        $DepartmentsBox.Items.Clear() 
        $DepartmentsBox.Text = $null
        $Hashcompanies = Get-Content ("C:\Users\Администратор.WIN-HBSQBQLKH1H\Downloads\ПроектITакадемия\texts\companies" + $Hashcities[$CityBox.Text]+".txt")-Raw | ConvertFrom-Json -AsHashtable
        $path = "C:\Users\Администратор.WIN-HBSQBQLKH1H\Downloads\ПроектITакадемия\texts\departments" + $Hashcompanies[$CompanyBox.Text] +$Hashcities[$CityBox.Text]+ ".txt"
        $HashDepartments = Get-Content ($path) -Raw | ConvertFrom-Json -AsHashtable
        Write-Output $path
        $Departments = [string[]] $HashDepartments.Keys
        $enDepartments = [string[]] $HashDepartments.Values
        $DepartmentsBox.Items.AddRange($Departments)
    }
    $CompanyBox.add_SelectedIndexChanged($Company_SelectedIndexChanged)

    $CurrForm.Controls.AddRange(@($LabelF,$FName,$LabelS,$SName,$LabelM,$MName,$LabelCity,$CityBox,$LabelCompany,$CompanyBox,$LabelDepartments,$DepartmentsBox,$LabelAppointment,$AppointmentBox,$Submit))

    [void]$CurrForm.ShowDialog()
}
function ConnectToDBFunc {
    function CheckUser     {
    try {
        $User = $UserLogin.Text
        $PWord = ConvertTo-SecureString -String $Passw.Text -AsPlainText -Force
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord 
        $exchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Uri.Text -ErrorAction SilentlyContinue -Authentication Kerberos -Credential $Credential
        Import-PSSession $exchSession -DisableNameChecking
        [System.Windows.MessageBox]::Show("Вам удалось авторизоваться в exchange","Поздравляем","OK","Information")
        $pass = $true
        $StartForm.Controls.Remove($ConnectToDBBtn)
        $StartForm.Controls.Add($AddToDBBtn)
        $StartForm.Controls.Add($DisconnectFromDBBtn)
      } catch {
        [System.Windows.MessageBox]::Show("Вы не смогли авторизоваться в Exchange.","Ошибка","OK","Error")
      }
    }
    $CurrForm                    = New-Object system.Windows.Forms.Form
    $CurrForm.ClientSize         = '500,300'
    $CurrForm.text               = "Подключиться к БД"

    $LabelL                           = New-Object system.Windows.Forms.Label
    $LabelL.text                      = "Введите логин"
    $LabelL.AutoSize                  = $true
    $LabelL.width                     = 25
    $LabelL.height                    = 10
    $LabelL.location                  = New-Object System.Drawing.Point(10,10)
    $LabelL.Font                      = 'Microsoft Sans Serif,13'
    
    $UserLogin = New-Object System.Windows.Forms.TextBox
    $UserLogin.Width = 400
    $UserLogin.Height = 30
    $UserLogin.location = New-Object System.Drawing.Point(10,45)
    $UserLogin.Font = 'Microsoft Sans Serif,10'

    $LabelP                           = New-Object system.Windows.Forms.Label
    $LabelP.text                      = "Введите пароль"
    $LabelP.AutoSize                  = $true
    $LabelP.width                     = 25
    $LabelP.height                    = 10
    $LabelP.location                  = New-Object System.Drawing.Point(10,80)
    $LabelP.Font                      = 'Microsoft Sans Serif,13'

    $Passw = New-Object System.Windows.Forms.TextBox
    $Passw.Width = 400
    $Passw.Height = 30
    $Passw.location = New-Object System.Drawing.Point(10,115)
    $Passw.Font = 'Microsoft Sans Serif,10'

    $LabelUri                           = New-Object system.Windows.Forms.Label
    $LabelUri.text                      = "Введите Uri для соединения"
    $LabelUri.AutoSize                  = $true
    $LabelUri.width                     = 25
    $LabelUri.height                    = 10
    $LabelUri.location                  = New-Object System.Drawing.Point(10,150)
    $LabelUri.Font                      = 'Microsoft Sans Serif,13'

    $Uri = New-Object System.Windows.Forms.TextBox
    $Uri.Width = 400
    $Uri.Height = 30
    $Uri.location = New-Object System.Drawing.Point(10,185)
    $Uri.Font = 'Microsoft Sans Serif,10'

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.BackColor         = "#DCDCDC"
    $Submit.text              = "Подключиться к MS Exchange"
    $Submit.width             = 300
    $Submit.height            = 30
    $Submit.location          = New-Object System.Drawing.Point(10,220)
    $Submit.Font              = 'Microsoft Sans Serif,10'
    $Submit.ForeColor         = "#000000"

    $Submit.Add_Click({ CheckUser })

    $CurrForm.Controls.AddRange(@($LabelL, $UserLogin, $LabelP,$Passw,$LabelUri,$Uri,$Submit))

    [void]$CurrForm.ShowDialog()
}
function AddToDBFunc {
    function CheckUser {
    $userInfo = Get-ADUser -identity $UserLogin.Text -ErrorAction Ignore
            if ( $null -ne $userInfo ) {
            try {
                Enable-Mailbox $UserLogin.Text
                [System.Windows.MessageBox]::Show("Почтовый ящик создан","Поздравляем","OK","Information")
                } catch {
                [System.Windows.MessageBox]::Show("Не удалось создать почтовый ящик. Убедить, что пользователь не подключен к БД","Ошибка","OK","Error")
                }
                
            } 
            else {
        [System.Windows.MessageBox]::Show("Невозможно найти пользователя $($UserLogin.Text)","Ошибка","OK","Error")
           }
          
        }
    $CurrForm                    = New-Object system.Windows.Forms.Form
    $CurrForm.ClientSize         = '500,300'
    $CurrForm.text               = "Добавление пользователя к почтовой БД"
    
    $LabelL                           = New-Object system.Windows.Forms.Label
    $LabelL.text                      = "Введите логин"
    $LabelL.AutoSize                  = $true
    $LabelL.width                     = 25
    $LabelL.height                    = 10
    $LabelL.location                  = New-Object System.Drawing.Point(10,10)
    $LabelL.Font                      = 'Microsoft Sans Serif,13'
    
    $UserLogin = New-Object System.Windows.Forms.ComboBox
    $users = Get-ADUser -Filter * | Select -ExpandProperty "Name"
    $UserLogin.Items.AddRange($users)
    $UserLogin.Width = 400
    $UserLogin.Height = 30
    $UserLogin.location = New-Object System.Drawing.Point(10,45)
    $UserLogin.Font = 'Microsoft Sans Serif,10'

    $Submit = New-Object system.Windows.Forms.Button
    $Submit.BackColor         = "#DCDCDC"
    $Submit.text              = "Добавить пользователя в БД"
    $Submit.width             = 300
    $Submit.height            = 30
    $Submit.location          = New-Object System.Drawing.Point(10,80)
    $Submit.Font              = 'Microsoft Sans Serif,10'
    $Submit.ForeColor         = "#000000"

    $Submit.Add_Click({ CheckUser })

    $CurrForm.Controls.AddRange(@($LabelL,$UserLogin,$Submit))

    [void]$CurrForm.ShowDialog()
}
function DisconnectFromDBFunc {
$exchSession = (Get-PSSession |where -property ConfigurationName -eq 'Microsoft.Exchange' | select -expand id)
if ($exchSession) {
    if([System.Windows.MessageBox]::Show('Точно?', 'Уверены?', 'YesNo','Question') -eq 'Yes'){
        try {Remove-PSSession $exchSession
            $pass = $false
            [System.Windows.MessageBox]::Show("Вы отключены от Exchange","Готово","OK","Information")
            $StartForm.Controls.Add($ConnectToDBBtn)
            $StartForm.Controls.Remove($AddToDBBtn)
            $StartForm.Controls.Remove($DisconnectFromDBBtn)}
        catch {[System.Windows.MessageBox]::Show("Произошла неизвестная ошибка","Ошибка","OK","Error")} }
        }
        else {
        [System.Windows.MessageBox]::Show("Немогу найти id. Перезапустите Powershell","Ошибка","OK","Error")
        }
}


$BlockUserBtn.Add_Click({ BlockUserFunc })
$UnbanUserBtn.Add_Click({ UnBanUserFunc })
$ShowMailBoxBtn.Add_Click({ ShowMailBoxFunc })
$HideMailBoxBtn.Add_Click({ HideMailBoxFunc })
$ChangePasswBtn.Add_Click({ ChangePasswFunc })
$AddToGroupBtn.Add_Click({ AddToGroupFunc })
$RemoveFromGroupBtn.Add_Click({ RemoveFromGroupFunc})
$GetUserInfoBtn.Add_Click({GetUserInfoFunc})
$GetGroupInfoBtn.Add_Click({GetGroupInfoFunc})
$CreateUserBtn.Add_Click({CreateUserFunc})
$ConnectToDBBtn.Add_Click({ConnectToDBFunc})
$DisconnectFromDBBtn.Add_Click({DisconnectFromDBFunc})
$AddToDBBtn.Add_Click({AddToDBFunc})

$StartForm.Controls.AddRange(@($BlockUserBtn,$UnbanUserBtn,$ShowMailBoxBtn,$HideMailBoxBtn,$ChangePasswBtn,$AddToGroupBtn,$RemoveFromGroupBtn,$GetUserInfoBtn,$GetGroupInfoBtn,$CreateUserBtn))


if ("Microsoft.Exchange" -in [string[]](Get-PSSession | Select-Object -expand ConfigurationName)) {
  $pass = $true
  $exchSession = (Get-PSSession |where -property ConfigurationName -eq 'Microsoft.Exchange' | select -expand id)
  } else {
  $exchSession = $null
  $pass = $false
}
if ($pass -eq $false) {
$StartForm.Controls.Add($ConnectToDBBtn)
$StartForm.Controls.Remove($AddToDBBtn)
$StartForm.Controls.Remove($DisconnectFromDBBtn)

}
if ($pass -eq $true) {
$StartForm.Controls.Remove($ConnectToDBBtn)
$StartForm.Controls.Add($AddToDBBtn)
$StartForm.Controls.Add($DisconnectFromDBBtn)
}


[void]$StartForm.ShowDialog()

