#Импорт модуля AD
Import-Module ActiveDirectory
#Путь до csv файла
$filePath = "S:\users&group.csv"
#Импорт cvs
$mytable = Import-Csv -Path $filePath
#Инициализация хеш-таблицы
$hash_users = @{}
#Формирование хеш-таблицы
foreach($property in $mytable)
{
    $hash_users[$property.login] = $property.group
}
#Функция получиния списка групп
function Get-Group {
    param (
        [string]
        $user
    )
    process {
        $groups_list = Get-ADPrincipalGroupMembership $user | Select-Object 'Name'
        return $groups_list.Name
    }
}
#Функция добавления групп
function Add-Group {
    param (
        [string]
        $user,
        [string]
        $first,
        [string]
        $second
    )
    process {
        Add-ADGroupMember -Identity "$($first)" -Members "$($user)"
        Add-ADGroupMember -Identity "$($first)_conf" -Members "$($user)"
        Add-ADGroupMember -Identity "$($second)" -Members "$($user)"
        Add-ADGroupMember -Identity "$($second)_conf" -Members "$($user)"
        return Write-Host "Группы добавлены!" -ForegroundColor Green
    }
}
foreach($user in $hash_users.GetEnumerator()) {
    #Парс числа основной группы
    $value = $user.Value[-2] + $user.Value[-1] 
    #Переменные групп
    $first_group = "remoteapp_logistics_development$($value)"
    $second_group = "remoteapp_logistics_development_cf$($value)"
    #Имя текущей учетки
    Write-Host "----"$user.Key"----" -ForegroundColor Red
    #Добавление основных групп
    Add-Group $user.Key $first_group $second_group 
    #Получение списка групп
    $group_list = Get-Group $user.Key
    #Начало поиска лишних групп
    Write-Host "Поиск лишних групп..." -ForegroundColor Yellow
    foreach ($group in $group_list) { 
        if ($group -notlike "$($first_group)" -and 
            $group -notlike "$($first_group)_conf" -and 
            $group -notlike "$($second_group)" -and 
            $group -notlike "$($second_group)_conf" -and 
            $group -like "*remoteapp_logistics_development*") {
            Remove-ADPrincipalGroupMembership -Identity $($user.Key) -MemberOf $group -Confirm:$false
            Write-Host $group -ForegroundColor Red
            } else {
                Write-Host $group -ForegroundColor Green
            }
        }
    Write-Host "Проверка окончена." -ForegroundColor Yellow
}
Write-Host "---- Завершение программы ----" -ForegroundColor Red