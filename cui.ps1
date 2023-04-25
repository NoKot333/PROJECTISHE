function usercreation {

  $firstName = Read-Host "Введите имя"
  if ($firstName -notmatch "[^А-Яа-я -]+") {
    $lastName = Read-Host "Введите фамилию"
    if ($lastName -notmatch "[^А-Яа-я -]+") {
        $midName = Read-Host "Введите отчество"
            if ($midName -notmatch "[^А-Яа-я -]+") {
                $extAttr1=@{}
                $extAttr1.Add('Имя',$firstName)
                $extAttr1.Add('Фамилия',$lastName)
                $extAttr1.Add('Отчество',$midName)

                $firstName = &"$PSScriptRoot\funcs\translit.ps1" $firstName
                $lastName = &"$PSScriptRoot\funcs\translit.ps1" $lastName
                $midName = &"$PSScriptRoot\funcs\translit.ps1" $midName

                $Hashcities = Get-Content "$PSScriptRoot\texts\cities.txt" -Raw | ConvertFrom-Json -AsHashtable
                $cities = [string[]] $Hashcities.Keys
                $encites = [string[]] $Hashcities.Values
                $index = 1
                foreach ($city in $cities){
                  Write-host "$index. $city" 
                  $index++
                }

                $ChooseIndex = Read-Host "Выберите город (введите цифру)"
                if (([int]$ChooseIndex -le 0) -or ([int]$ChooseIndex -ge [int]$index)) {
                  Write-Host "Вы ввели неправильный индекс!" 
                } else {
                $cityName = $encites[$ChooseIndex-1]
                $extAttr1.Add('Город',$cities[$ChooseIndex-1])

                $Hashcompanies = Get-Content "$PSScriptRoot\texts\companies$cityName.txt" -Raw | ConvertFrom-Json -AsHashtable
                $companies = [string[]] $Hashcompanies.Keys
                $encompanies = [string[]] $Hashcompanies.Values
                $index = 1
                foreach ($company in $companies){
                  Write-host "$index. $company" 
                  $index++
                }

                $ChooseIndex = Read-Host "Выберите компанию (введите цифру)"
                if (([int]$ChooseIndex -le 0) -or ([int]$ChooseIndex -ge $index)) {
                  Write-Host "Вы ввели неправильный индекс!" 
                }
                else {
                  $companyName = $encompanies[$ChooseIndex-1]
                  $extAttr1.Add('Компания',$companies[$ChooseIndex-1])

                  $HashDepartments = Get-Content "$PSScriptRoot\texts\departments$companyName$cityName.txt" -Raw | ConvertFrom-Json -AsHashtable
                  $Departments = [string[]] $HashDepartments.Keys
                  $enDepartments = [string[]] $HashDepartments.Values
                  $index = 1
                  foreach ($Department in $Departments){
                    Write-host "$index. $Department" 
                    $index++
                }

                $ChooseIndex = Read-Host "Выберите Отдел (введите цифру)"
                if (([int]$ChooseIndex -le 0) -or ([int]$ChooseIndex -ge [int]$index)) {
                  Write-Host "Вы ввели неправильный индекс!" 
                } else {
                  $departmentName = $enDepartments[$ChooseIndex-1]
                  $extAttr1.Add('Отдел',$Departments[$ChooseIndex-1])
                
                  $Hashappointments = Get-Content "$PSScriptRoot\texts\appointment.txt" -Raw | ConvertFrom-Json -AsHashtable
                  $appointments = [string[]] $Hashappointments.Keys
                  $enappointments = [string[]] $Hashappointments.Values
                  $index = 1
                  foreach ($appointment in $appointments){
                    Write-host "$index. $appointment" 
                    $index++
                  }

                  $ChooseIndex = Read-Host "Выберите должность (введите цифру)"
                  if (([int]$ChooseIndex -le 0) -or ([int]$ChooseIndex -ge [int]$index)) {
                    Write-Host "Вы ввели неправильный индекс!" 
                  } else {
                    $appointmentName = $enappointments[$ChooseIndex-1]
                    $extAttr1.Add('Должность',$appointments[$ChooseIndex-1])

                    $index = 0  
                    $baseUserName = "$firstName.$lastName" 
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

                    $UserPName = "$userName@MYDOMAIN.com"

                    Write-Host "Пользователь: $userName"
                    Write-Host "UserPrincipalName: $UserPName"
                    write-host "Имя: $firstName"
                    write-host "Фамилия: $lastName"
                    Write-Host "Компания: $companyName"
                    Write-Host "Отдел: $departmentName"
                    Write-Host "Город: $cityName"
                    Write-Host "Должность: $appointmentName"

                    Write-Host "Имя : $($extAttr1["Имя"])"
                    Write-Host "Фамилия : $($extAttr1["Фамилия"])"
                    Write-Host "Отчество : $($extAttr1["Отчество"])"
                    Write-Host "Город : $($extAttr1["Город"])"
                    Write-Host "Компания : $($extAttr1["Компания"])"
                    Write-Host "Отдел : $($extAttr1["Отдел"])"
                    Write-Host "Должность : $($extAttr1["Должность"])"

                    $accept=read-host "Подтвердить создание пользователя (y/n)"
                    if($accept -eq "y"){

                      New-ADUser -Name $UserName -UserPrincipalName $UserPName -Department $departmentName -GivenName $firstName  -Surname $lastName -OtherName $appointmentName -SamAccountName $userName -City $cityName -Company $companyName -title $appointmentName -Enabled $true
                      
                      Set-ADAccountPassword $UserName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText “PASSWORD” -Force -Verbose) | Set-ADuser -ChangePasswordAtLogon $True

                      Set-ADuser -Identity $userName -Add @{extensionAttribute1 = $extAttr1["Имя"]} 
                      Set-ADuser -Identity $userName -Add @{extensionAttribute2 = $extAttr1["Фамилия"]}
                      Set-ADuser -Identity $userName -Add @{extensionAttribute3 = $extAttr1["Отчество"]}
                      Set-ADuser -Identity $userName -Add @{extensionAttribute4 = $extAttr1["Компания"]}
                      Set-ADuser -Identity $userName -Add @{extensionAttribute5 = $extAttr1["Должность"]}
                      Set-ADuser -Identity $userName -Add @{extensionAttribute6 = $extAttr1["Отдел"]}
                      Write-Host "Пользователь создан" 
                    } else {
                      Write-Host "Создание пользователя отменено"
                    }
                  }
                }
              }
            }
          } else {
    Write-Host "Вы использовали символы помимо букв кириллицы, пробела и -"
          }
        } else {
          Write-Host "Вы использовали символы помимо букв кириллицы, пробела и -"
        }
        } else {
          Write-Host "Вы использовали символы помимо букв кириллицы, пробела и -"
        }
}

function GetUserInfo {
  $userID = Read-Host "Введите имя пользователя"
  if ( $null -ne (Get-ADUser -identity $userID -ErrorAction SilentlyContinue) ) {
    Get-ADUser -identity $userID -properties SamAccountName,EmailAddress,Enabled,PasswordExpired,PasswordLastSet,MemberOf | select-object DistinguishedName,SamAccountName,EmailAddress,Enabled,PasswordExpired,PasswordLastSet,MemberOf
  }
  else {
    Write-Host "Невозможно найти пользователя '$userID'"
       } 
  
}

function GetGroupInfo {
  $GroupID = Read-Host "Введите SamAccountName группы"
  if ( $null -ne (Get-ADGroup -identity $GroupID -ErrorAction SilentlyContinue) ) {
    Get-ADGroup -identity $GroupID  | select-object SamAccountName,DistinguishedName,GroupCategory,GroupScope
    "Members = $(Get-ADGroupMember LazyPeople | Select-Object -expand Name)"
  }
  else {
    Write-Host "Невозможно найти группу '$GroupID'"
       } 
  
}


function ShowMailBox { 
  $userID = Read-Host "Введите имя пользователя"
  if ($null -ne (Get-ADUser -identity $userID -ErrorAction SilentlyContinue) ) {
      Set-ADUser -Identity $userID -replace @{msExchHideFromAddressLists=$false}
      Write-Host "Ящик не скрыт в адресной книге"
  } else {
    Write-Host "Пользователь не найден"
  }
}

function HideMailBox {
  $userID = Read-Host "Введите имя пользователя"
  if ($null -ne (Get-ADUser -identity $userID -ErrorAction SilentlyContinue) ) {
      Set-ADUser -Identity $userID -replace @{msExchHideFromAddressLists=$true}
      Write-Host "Ящик скрыт в адресной книге"
  } else {
    Write-Host "Пользователь не найден"
  }
}

function createMailBox {
  $userID = Read-Host "Введите имя пользователя"
  if ( $null -ne (Get-ADUser -identity $userID -ErrorAction SilentlyContinue) ) {
    $index = 1
    $servers = [string[]] (Get-MailboxDatabase | Select-Object -expand Server)
    foreach ($server in $servers){
      Write-host "$index. $server" 
      $index++
    }
    $ChooseIndex = Read-Host "Выберите БД (введите цифру)"
    if (($ChooseIndex -le 0) -or ($ChooseIndex -ge $index)) {
          Write-Host "Вы ввели неправильный индекс!"
        } 
        else {
          $server = $servers[$index-1]
          Enable-Mailbox $userID
          Write-Host "Почтовый ящик создан"
        }
      }
      else {
        Write-Host "Невозможно найти пользователя '$userID'"
      }    
}

function BanUser {
    $userID = Read-Host "Введите имя пользователя"
      if ( $null -ne (Get-ADUser -identity $userID -ErrorAction SilentlyContinue) ) {
        if ( (Get-ADUser -identity $userID  | Select-Object -expand Enabled) -eq $true) {
            $accept=read-host "Подтвердить блокировку пользователя (y/n)"
  if($accept -eq "y"){
                try {Disable-ADAccount -identity $userID
                Write-Host "Пользователь заблокирован"}
                catch {Write-Host "Произошла неизвестная ошибка"} }
            else {Write-Host "Блокировка отменена"}
            } else {
            Write-Host "Этот пользователь уже заблокирован"
            }
      } 
      else {
        Write-Host "Невозможно найти пользователя '$userID'"
           } 
}

function ChangePassw {
    $userID = Read-Host "Введите имя пользователя"
      if ( $null -ne (Get-ADUser -identity $userID -ErrorAction SilentlyContinue ) ) {
        $fpassw = Read-Host "Введите новый пароль пользователя"
        $spassw = Read-Host "Введите ещё раз новый пароль пользователя"
        if ($fpassw -eq $spassw) {
        $accept=read-host "Подтвердить смену пароля пользователя (y/n)"
        if($accept -eq "y"){
          try {@( Set-ADAccountPassword $UserName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $fpassw -Force -Verbose) -ErrorAction SilentlyContinue | Set-ADuser -ChangePasswordAtLogon $True )
            Write-Host "Пароль пользователя обновлён"}
          catch {Write-Host "Произошла неизвестная ошибка"} }
        else {Write-Host "Смена пароля отменена отменена"}
      } else {Write-Host "Пароли не совпадают"}
      }
      else {
        Write-Host "Невозможно найти пользователя '$userID'"
      } 
}


function UnbanUser {
    $userID = Read-Host "Введите sAMAccount имя пользователя"
      if ( $null -ne (Get-ADUser -identity $userID -ErrorAction SilentlyContinue )) {
        if ( (Get-ADUser -identity $userID  | Select-Object -expand Enabled) -eq $false) {
        $accept=read-host "Подтвердить разблокировку пользователя (y/n)"
        if($accept -eq "y")
            {
            try {Enable-ADAccount -identity $userID
             Write-Host "Пользователь разблокирован"}
            catch {Write-Host "Произошла неизвестная ошибка"}
        }
        else {Write-Host "Разблокировка отменена"}
        } else {
        Write-Host "Это пользователь не заблокирован"
        
        }
      } 
      else {
        Write-Host "Невозможно найти пользователя '$userID'"
           } 
}

function Add-ToGroup {
    $grpID = Read-Host "Введите название группы"
    if ($null -ne (Get-ADGroup -Identity $grpID -ErrorAction SilentlyContinue )) {
      $userID = Read-Host "Введите sAMAccount имя пользователя"
      if ( $null -ne (Get-ADUser -identity $userID -ErrorAction SilentlyContinue )) {
        $members = Get-ADGroupMember -Identity $grpID -Recursive | Select-Object -ExpandProperty SamAccountName
        if ( $members -notcontains $userID) {
        $accept=read-host "Подтвердить добавление пользователя '$userID' в группу '$grpID' (y/n)"
        if($accept -eq "y")
            {
            try {Add-ADGroupMember -identity $grpID -Members $userID
            Write-Host "Пользователь $userID добавлен к $grpID"}
            catch {Write-Host "Произошла неизвестная ошибка"}
        }
        else {Write-Host "Добавление пользователя было отменено"}
        } else {
        Write-Host "Пользователь $userID уже есть в $grpID"
        
        }
      } 
      else {
        Write-Host "Пользователь не может быть найден"
           } 
           }
    else {
           Write-Host "Группа не найдена"
           }
}

function Remove_fromGroup {
    $grpID = Read-Host "Введите название группы"
    if ($null -ne (Get-ADGroup -Identity $grpID -ErrorAction SilentlyContinue)) {
      $userID = Read-Host "Введите sAMAccount имя пользователя"
      if ( $null -ne (Get-ADUser -identity $userID -ErrorAction SilentlyContinue )) {
        if ( $null -ne ((Get-ADUser $userID -Properties MemberOf).memberof -like "*$grpID*")) {
        $accept=read-host "Подтвердить удаление пользователя '$userID' из группы '$grpID' (y/n)"
        if($accept -eq "y")
            {
            try {Remove-ADGroupMember -identity $grpID -Members $userID
            Write-Host "Пользователь $userID исключен из $grpID"}
            catch {Write-Host "Произошла неизвестная ошибка"}
        }
        else {Write-Host "Исключение пользователя было отменено"}
        } else {
        Write-Host "Пользователь $userID не существует в $grpID"
        
        }
      } 
      else {
        Write-Host "Пользователь не может быть найден"
           } 
           }
    else {
           Write-Host "Группа не найдена"
           }

}

if ("Microsoft.Exchange" -in [string[]](Get-PSSession | Select-Object -expand ConfigurationName)) {
  $pass = $true
  } else {
  $pass = $false
}

$ChooseIndex = 1
While (("0" -ne $ChooseIndex) -and ($pass -eq $false)) {
$ChooseIndex = Read-Host "

Выберите опцию:
0)Выход из программы

1)Заблокировать пользователя
2)Разблокировать пользователя
3)Сменить пароль пользователя

4)Добавить к группе
5)Удалить из группы

6)Создать Пользователя

7)Транслитерация текста

8)Получить информацию о пользователе
9)Получить информацию о группе

10) Скрыть почтовый ящик из адресной книги
11) Показать почтовый ящик в адресной книге

12)Авторизоваться в MS Exchange"
switch ($ChooseIndex) {
    0 {break}
    1 {BanUser}
    2 {UnbanUser}
    3 {ChangePassw}
    4 {Add-ToGroup}
    5 {Remove_fromGroup}
    6 {usercreation}
    7 {
    $msg = Read-Host "Введите текст для транслитерации"
    . "$PSScriptRoot\funcs\translit.ps1" $msg}
    8 {GetUserInfo}
    9 {GetGroupInfo}
    10 {HideMailBox}
    11 {ShowMailBox}
    12 {
      Write-Host "Введите данные вашей учётной записи MS Exchange:"
      try {
        $UserCredential = Get-Credential
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://WIN-HBSQBQLKH1H.MYDOMAIN.com/powershell -ErrorAction SilentlyContinue -Authentication Kerberos -Credential $UserCredential
        Import-PSSession $Session -DisableNameChecking
        Write-Host "Вам удалось авторизоваться в exchange, поздравляем"
        $pass = $true
      } catch {
        Write-Host "Произошла ошибка, вы не смогли авторизоваться в Exchange."
      }
    }
    default {"Неизвестная команда."
    break}
}
Read-Host "Нажмите, чтобы продолжить"
} 

While (("0" -ne $ChooseIndex) -and ($pass -eq $true) ){
$ChooseIndex = Read-Host "

Выберите опцию:
0)Выход из программы

1)Заблокировать пользователя
2)Разблокировать пользователя
3)Сменить пароль пользователя

4)Добавить к группе
5)Удалить из группы

6)Создать Пользователя

7)Транслитерация текста

8)Получить информацию о пользователе
9)Получить информацию о группе

10) Скрыть почтовый ящик из адресной книги
11) Показать почтовый ящик в адресной книге

12)Добавить почтовый ящик Exchange"
switch ($ChooseIndex) {
    0 {break}
    1 {BanUser}
    2 {UnbanUser}
    3 {ChangePassw}
    4 {Add-ToGroup}
    5 {Remove_fromGroup}
    6 {usercreation}
    7 {
    $msg = Read-Host "Введите текст для транслитерации"
    . "$PSScriptRoot\funcs\translit.ps1" $msg}
    8 {GetUserInfo}
    9 {GetGroupInfo}
    10 {HideMailBox}
    11 {ShowMailBox}
    12 {createMailBox}
    default {"Неизвестная команда."}
}
Read-Host "Нажмите, чтобы продолжить"
}