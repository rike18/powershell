#Login and Pass
# $Cred = Get-Credential Для ввода логина и пароля (Свойства объекта $Cred.Username и $Cred.Password)
$username = "pavel.tarasov"
$pass = Read-Host "Введите пароль" –AsSecureString
#Загрузка имён хостов
$computers = "10-pc122" #Get-Content S:\computers.txt
#Имя программы
$appname = "VLC media player"
$appkey = "*VLC*"
#Путь программы
$file = 'S:\Skype-8.65.0.78.msi'
foreach ($computer in $computers) {
    #Проверка хоста
    if (test-Connection -ComputerName $computer -quiet) { 

        Write-Output "=========================";
        Write-Output "         $computer";
        Write-Output "=========================";

        Write-Output "Поиск программы..."
        $app = Get-WmiObject -ComputerName $computer Win32_Product -Filter "Name like '$appname%'"

        if ($app.Name -like "$appkey") {
            Write-Output "Программа уже установлена!"
            Write-Output "Хотите удалить ? [Y] Да - Y [N] НЕТ - Y"
            $x = Read-Host
            switch ($x) {
                Y {"Началось удаление программы...";$app.Uninstall();"Программа удалена!"}
                N {Write-Output "ОК"}
                default {Write-Output "ВЫХОД"} 
            }
        }
        else {
            Write-Output "Программа не установлена.Запускается процесс установки..."
            #Копируем файл
            Copy-Item -Path $file -Destination "\\$computer\c$\windows\temp\installer.msi"
            #Начало установки
            psexec.exe \\$computer -s -u "$username" -p "$pass" msiexec /i C:\windows\temp\installer.msi /qb
            Write-Output "Установка завершена!";
        }
    }
    else {Write-Output "IN NOT ONLINE"}
}