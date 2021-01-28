

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#подключаем оснастку AD
Import-Module ActiveDirectory
#подключаем остастку exchange 2016
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn 

###########функция транслитерации###########
function global:TranslitToLAT
{
 	param([string]$inString)
	$Translit_To_LAT = @{ 
	[char]'а' = "a"
	[char]'А' = "A"
	[char]'б' = "b"
	[char]'Б' = "B"
	[char]'в' = "v"
	[char]'В' = "V"
	[char]'г' = "g"
	[char]'Г' = "G"
	[char]'д' = "d"
	[char]'Д' = "D"
	[char]'е' = "e"
	[char]'Е' = "E"
	[char]'ё' = "e"
	[char]'Ё' = "E"
	[char]'ж' = "zh"
	[char]'Ж' = "Zh"
	[char]'з' = "z"
	[char]'З' = "Z"
	[char]'и' = "i"
	[char]'И' = "I"
	[char]'й' = "i"
	[char]'Й' = "I"
	[char]'к' = "k"
	[char]'К' = "K"
	[char]'л' = "l"
	[char]'Л' = "L"
	[char]'м' = "m"
	[char]'М' = "M"
	[char]'н' = "n"
	[char]'Н' = "N"
	[char]'о' = "o"
	[char]'О' = "O"
	[char]'п' = "p"
	[char]'П' = "P"
	[char]'р' = "r"
	[char]'Р' = "R"
	[char]'с' = "s"
	[char]'С' = "S"
	[char]'т' = "t"
	[char]'Т' = "T"
	[char]'у' = "u"
	[char]'У' = "U"
	[char]'ф' = "f"
	[char]'Ф' = "F"
	[char]'х' = "kh"
	[char]'Х' = "Kh"
	[char]'ц' = "ts"
	[char]'Ц' = "Ts"
	[char]'ч' = "ch"
	[char]'Ч' = "Ch"
	[char]'ш' = "sh"
	[char]'Ш' = "Sh"
	[char]'щ' = "shch"
	[char]'Щ' = "Shch"
	[char]'ъ' = "ie"		# "``"
	[char]'Ъ' = "Ie"		# "``"
	[char]'ы' = "y"		# "y`"
	[char]'Ы' = "Y"		# "Y`"
	[char]'ь' = ""		# "`"
	[char]'Ь' = ""		# "`"
	[char]'э' = "e"		# "e`"
	[char]'Э' = "E"		# "E`"
	[char]'ю' = "iu"
	[char]'Ю' = "Iu"
	[char]'я' = "ia"
	[char]'Я' = "Ia"
    #[char]'-' = "-"
	}
	$outChars=""
	foreach ($c in $inChars = $inString.ToCharArray())
		{
		if ($Translit_To_LAT[$c] -cne $Null ) 
			{$outChars += $Translit_To_LAT[$c]}
		else
			{$outChars += $c}
		}
	Write-Output $outChars
 }
function LogAdd($msg)
	{
	$LogBox.text = $LogBox.text + $msg + [char]13
	
 	}


##########определение группы по орнанизации###########
function UserCreate()
{
$LogBox.text = ""
#задаем базу для создания почтового ящика
$mailboxDB = "ExchangeBase"
#задаем пароль
$pass = ConvertTo-SecureString -String "12345678" -AsPlainText -Force
#OU куда будут создаваться учетки
$OU = "OU=OUnit,DC=domain,DC=local"
#место где будет создаваться пользовательский диск
$zdisk = "\\fileerver\users\"
#сплитуем ФИО что бы разделить на отдельные фамилию, имя, отчество
$fioarray = $fio_TextBox.Text.split(" ", 3)
$surname = $fioarray[0]
$name = $fioarray[1]
$mname = $fioarray[2]

#выбираем группу в зависимости от организации
Switch ($company_ComboBox.SelectedItem)
{
    'АО «Рога и копыта»'{$group = "roga_users"}
    'АО «Ромашка»'{$group = "romashka_users"}
    'АО Завод «Бабайкин»'{$group = "babayka_users"}
   
    default {$group = $null
    LogAdd ("[WARNING] Группа не выбрана")
            }
}
#алиас для почтового ящика
$alias = TranslitToLAT ($name + "." + $surname)
#первый перевод логина с поощью функции транслитерации
$login = TranslitToLAT ($surname) 
#проверка логина на существование
 if ((Get-ADUser -f 'sAMAccountName -eq $login' -Server DC.domain.local:3268) -eq $null) 
 {
 #создаем upn на основе логани
    $upn = $login + "@domain.local" 
    #создаем окончательно путь к пользовательскому диску
    $zdisk = $zdisk + $login
    #создаем пользователя
    New-ADUser -Name $fio_TextBox.Text -DisplayName $fio_TextBox.Text -GivenName $name -Surname $surname -SamAccountName $login -UserPrincipalName $upn -Path $OU -Department $dep_textBox.text -Company $company_ComboBox.SelectedItem -Title $jtitle_TextBox.Text -OfficePhone $phone_textBox.Text -Office $company_ComboBox.SelectedItem -Description $jtitle_TextBox.Text -HomeDrive Z: -HomeDirectory $zdisk -AccountPassword $pass -ChangePasswordAtLogon $true -Enabled $true
    LogAdd ("[OK] Присвоен логин:  " + $login)
 }
 else {
 #если логин занят то к нему прицепляем первую букву имени
    LogAdd ("[ERROR] " + $login + " логин занят. Пытаюсь создать новый.")
    $login = TranslitToLAT($surname  + $name.Substring(0,1))

    #проверяем существует ли такой логин
    if ((Get-ADUser -f 'sAMAccountName -eq $login' -Server DC.domain.local:3268) -eq $null)
    {
    $upn = $login + "@domain.local" 
    $zdisk = $zdisk + $login
    New-ADUser -Name $fio_TextBox.Text -DisplayName $fio_TextBox.Text -GivenName $name -Surname $surname -SamAccountName $login -UserPrincipalName $upn -Path $OU -Department $dep_textBox.text -Company $company_ComboBox.SelectedItem -Title $jtitle_TextBox.Text -OfficePhone $phone_textBox.Text -Office $company_ComboBox.SelectedItem -Description $jtitle_TextBox.Text -HomeDrive Z: -HomeDirectory $zdisk -AccountPassword $pass -ChangePasswordAtLogon $true -Enabled $true
    LogAdd ("[OK] Присвоен логин:  " + $login)  
    }
    else{
        
        LogAdd ("[ERROR] " + $login + " логин занят. Пытаюсь создать новый.")
        #если и этот логин занят то прицепляем первую букву отчество
        $login = TranslitToLAT($surname  + $name.Substring(0,1) + $mname.Substring(0,1))
        #проверяем логин на существование
        if ((Get-ADUser -f 'sAMAccountName -eq $login' -Server DC.domain.local:3268) -eq $null)
        {
           $upn = $login + "@domain.local" 
           $zdisk = $zdisk + $login
           New-ADUser -Name $fio_TextBox.Text -DisplayName $fio_TextBox.Text -GivenName $name -Surname $surname -SamAccountName $login -UserPrincipalName $upn -Path $OU -Department $dep_textBox.text -Company $company_ComboBox.SelectedItem -Title $jtitle_TextBox.Text -OfficePhone $phone_textBox.Text -Office $company_ComboBox.SelectedItem -Description $jtitle_TextBox.Text -HomeDrive Z: -HomeDirectory $zdisk -AccountPassword $pass -ChangePasswordAtLogon $true -Enabled $true
           LogAdd ("[OK] Присвоен логин:  " + $login)
        }

        else
            {
            #если и этот логин занят, то завершаем работу и разбираемся с этим
            LogAdd ("[ERROR] " + $login + " логин занят. Требуется помощь кожанного ублюдка. Работа остановлена.")
            return          
            }
        }

 }

#проверяем существование пользовательской шары
sleep 7
if(!(Test-Path -Path $zdisk))
        {
        #если нет, то создаем пользовательский диск
         New-item $zdisk -type directory | out-null
         #получаем разрешения на созданную папку
         $acl = Get-Acl $zdisk -Audit | select
         #возвращаем коллекцию записей списка управления доступом
         $acl.GetAccessRules($true, $true, [System.Security.Principal.NTAccount])
         #указываем, применяются ли наследуемые правила контроля доступа к данному объекту файловой системы
         $acl.SetAccessRuleProtection($false, $true)
         #тоже самое указываем для аудита
         $acl.GetAuditRules($true, $true, [System.Security.Principal.NTAccount])
         $acl.SetAuditRuleProtection($false, $true)
         #создаем новую запись
         $rule = New-Object System.Security.AccessControl.FileSystemAccessRule ("domain\$login","FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
         #добавляем запись к списку контроля доступа
         $acl.addAccessRule($rule)
         #Записываем все на папку
         Set-Acl $zdisk $acl | out-null
         LogAdd ("[OK] Пользовательская шара создана")
        }
        else
        {
            #если папка существует, то надо разбираться
            LogAdd ("[ERROR] Такая пользовательская шара уже существует")
        }

#если была выбрана группа, то добавляем ее пользователю
if ($group -ne $null)
            {
                Add-ADGroupMember -Identity $group -Members $login
                LogAdd ("[OK] Добавлена группа:  " + $group)
            }

#если стоит галочка, то создаем почтовый ящик
if ($mail_CheckBox.Checked -eq $true)
{ 
#ждем 7 секунд
    sleep 7
    #создаем почтовый ящик
    Enable-Mailbox -Identity $upn -Database $mailboxDB -Alias $alias | Out-Null
    LogAdd ("[OK] Почтовый ящик создан") 
}
else {
    LogAdd ("[WARNING] Почтовый ящик не создан")

     }


LogAdd ("[OK] Работа закончена")
}


###############################Формы##################################################

$Mform                           = New-Object system.Windows.Forms.Form
$Mform.ClientSize                = New-Object System.Drawing.Point(390,521)
$Mform.text                      = "Создание учётной записи"
$Mform.TopMost                   = $false

$create_button                   = New-Object system.Windows.Forms.Button
$create_button.text              = "Создать"
$create_button.width             = 113
$create_button.height            = 33
$create_button.location          = New-Object System.Drawing.Point(239,216)
$create_button.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$create_button.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#000000")
$create_button.add_click($Function:UserCreate)

$mail_CheckBox                   = New-Object system.Windows.Forms.CheckBox
$mail_CheckBox.text              = "Создать почтовый ящик"
$mail_CheckBox.AutoSize          = $false
$mail_CheckBox.width             = 234
$mail_CheckBox.height            = 20
$mail_CheckBox.location          = New-Object System.Drawing.Point(135,178)
$mail_CheckBox.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$fio_label                       = New-Object system.Windows.Forms.Label
$fio_label.text                  = "ФИО "
$fio_label.AutoSize              = $true
$fio_label.width                 = 25
$fio_label.height                = 10
$fio_label.location              = New-Object System.Drawing.Point(37,33)
$fio_label.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$fio_TextBox                     = New-Object system.Windows.Forms.TextBox
$fio_TextBox.multiline           = $false
$fio_TextBox.width               = 217
$fio_TextBox.height              = 20
$fio_TextBox.location            = New-Object System.Drawing.Point(136,27)
$fio_TextBox.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$jtitle_label                    = New-Object system.Windows.Forms.Label
$jtitle_label.text               = "Должность"
$jtitle_label.AutoSize           = $true
$jtitle_label.width              = 25
$jtitle_label.height             = 10
$jtitle_label.location           = New-Object System.Drawing.Point(37,60)
$jtitle_label.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$jtitle_TextBox                  = New-Object system.Windows.Forms.TextBox
$jtitle_TextBox.multiline        = $false
$jtitle_TextBox.width            = 217
$jtitle_TextBox.height           = 20
$jtitle_TextBox.location         = New-Object System.Drawing.Point(136,54)
$jtitle_TextBox.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$dep_label                       = New-Object system.Windows.Forms.Label
$dep_label.text                  = "Отдел"
$dep_label.AutoSize              = $true
$dep_label.width                 = 25
$dep_label.height                = 10
$dep_label.location              = New-Object System.Drawing.Point(37,86)
$dep_label.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$dep_textBox                     = New-Object system.Windows.Forms.TextBox
$dep_textBox.multiline           = $false
$dep_textBox.width               = 217
$dep_textBox.height              = 20
$dep_textBox.location            = New-Object System.Drawing.Point(136,82)
$dep_textBox.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$company_label                   = New-Object system.Windows.Forms.Label
$company_label.text              = "Организация"
$company_label.AutoSize          = $true
$company_label.width             = 25
$company_label.height            = 10
$company_label.location          = New-Object System.Drawing.Point(37,112)
$company_label.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$company_ComboBox                = New-Object system.Windows.Forms.ComboBox
$company_ComboBox.text           = "comboBox"
$company_ComboBox.width          = 217
$company_ComboBox.height         = 20
$company_ComboBox.location       = New-Object System.Drawing.Point(136,109)
$company_ComboBox.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$company_ComboBox.DataSource = @('АО «Завод «Киров-Энергомаш»','АО «Металлургический завод «Петросталь»','АО Завод «Универсалмаш»','АО «КировТЭК»','АО «ЭСК»','АО «Тетрамет»','ПАО «Кировский завод»','АО «Локомотив»','АО «Центр МИОТ»','АО «Петербургский тракторный завод»','АО «Промышленный комплекс «Энергия»','ООО "Охранная организация "Путиловец"')


$cancel_Button1                  = New-Object system.Windows.Forms.Button
$cancel_Button1.text             = "Отмена"
$cancel_Button1.width            = 93
$cancel_Button1.height           = 32
$cancel_Button1.location         = New-Object System.Drawing.Point(262,446)
$cancel_Button1.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$cancel_Button1.add_click({$Mform.Close()})

$LogBox                          = New-Object windows.Forms.RichTextBox
$LogBox.text                     = ""
$LogBox.width                    = 317
$LogBox.height                   = 151
$LogBox.location                 = New-Object System.Drawing.Point(37,268)
$LogBox.ReadOnly                 = "true"


$phone_textBox                   = New-Object system.Windows.Forms.TextBox
$phone_textBox.multiline         = $false
$phone_textBox.width             = 217
$phone_textBox.height            = 20
$phone_textBox.location          = New-Object System.Drawing.Point(135,138)
$phone_textBox.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$phone_label                     = New-Object system.Windows.Forms.Label
$phone_label.text                = "Телефон"
$phone_label.AutoSize            = $true
$phone_label.width               = 25
$phone_label.height              = 10
$phone_label.location            = New-Object System.Drawing.Point(37,145)
$phone_label.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Mform.controls.AddRange(@($create_button,$mail_CheckBox,$fio_label,$fio_TextBox,$jtitle_label,$jtitle_TextBox,$dep_label,$dep_textBox,$company_label,$company_ComboBox,$cancel_Button1,$LogBox,$phone_textBox,$phone_label))

$Mform.ShowDialog()