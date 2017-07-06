# ActiveDirectory_Lib

## Краткое описание
Библиотека для платформы IBM Notes/Domino, облегчающая взаимодействие и интеграцию с MS Active Directory.
* Язык: LotusScript
* Текущая версия: 2.1.0
* Лицензия: [MIT](https://opensource.org/licenses/MIT)
* Технические ограничения использования: работа только на ОС под управлением Windows

Библиотека берет на себя рутинные операции работе с MS Active Directory (MS AD) по протоколу LDAP с помощью провайдера ADSI (Active Directory Service Interface),
оборачивая громоздкие вызовы в объекты и методы. Для взаимодействия с ADSI используются его COM-объекты, включенные в семейство ОС Windows.

## Установка
1. Для использования библиотеки вы можете импортировать DXL-файл в свое Notes-приложение. Стандартной функции по импорту из DXL нет, поэтому вы можете создать агент, воспользовавшись кодом, приведенным [здесь](http://www-10.lotus.com/ldd/bpmpblog.nsf/dx/the-missing-dxl-import-menu-option?opendocument&comments)
2. Создать в Notes-приложении новую библиотеку и скопировать все содержимое lss-файла.

## Быстрый старт
1. Создайте в Notes-приложении новую библиотеку и скопируйте все содержимое lss-файла
2. Подключитесь к MS AD по своими учетными данными
```
Dim ad As New ADConnector("MyOrg.org", "", "")
```
3. Получите объект вашей учетной записи
```
Dim adUser As ADUser
Set adUser = ad.FindObj(Nothing, AD_OBJECT_USER, "YOUR_LOGIN", AD_ATTRIBUTE_SAMACCOUNTNAME)
```
4. Убедитесь что это именно он
```
MsgBox adUser.ADsPath
```

## Полная документация

### Константы
#### Типы объектов
* AD_OBJECT_OU - контейнер
* AD_OBJECT_USER - пользователь
* AD_OBJECT_GROUP - группа
* AD_OBJECT_COMPUTER - устройство

#### Предопределенные атрибуты объектов
* AD_ATTRIBUTE_DISTINGUISHEDNAME
* AD_ATTRIBUTE_SAMACCOUNTNAME
* AD_ATTRIBUTE_CN
* AD_ATTRIBUTE_NAME
* AD_ATTRIBUTE_DESCRIP

#### Виды модификации атрибутов объекта
* ADS_PROPERTY_CLEAR (1) - очистить
* ADS_PROPERTY_UPDATE (2) - обновить
* ADS_PROPERTY_APPEND (3) - добавить
* ADS_PROPERTY_DELETE (4) - удалить

### Классы и методы
* [ADConnector](./adconnector.md)
* [ADObject](./adobject.md)
* [ADOU](./concrete-classes.md)
* [ADUser](./concrete-classes.md#aduser)
* [ADGroup](./concrete-classes.md#adgroup)

## Примеры использования
### Подключение к домену
```
'Используя текущую учетную запись
Dim ad As New ADConnector("MyOrg.org", "", "")

'С указанием учетной записи для подключения
Dim ad As New ADConnector("controller-1.MyOrg.org", "MyOrg\LotusAccount", "qwerty") 
```
### Создание объекта
```
Dim ad As New ADConnector("MyOrg.org", "", "")
Dim adContainer As ADOU
Dim adUser As ADUser

'Поиск контейнера, где будет создан объект
Set adContainer = ad.FindObj(Nothing, AD_OBJECT_OU, "UserContainer", AD_ATTRIBUTE_NAME)
'Создание объекта
Set adUser = ad.CreateADObject(AD_OBJECT_USER, "Petrov-IS", adContainer) 
```
### Поиск объекта
```
Dim ad As New ADConnector("MyOrg.org", "", "")
Dim adUser As ADUser

'Поиск пользователя по всему дереву
Set adUser = ad.FindObj(Nothing, AD_OBJECT_USER, "Petrov-IS", AD_ATTRIBUTE_SAMACCOUNTNAME)

'Поиск пользователя в поддереве контейнера
Dim adContainer As ADOU
Set adContainer = ad.FindObj(Nothing, AD_OBJECT_OU, "UserContainer", AD_ATTRIBUTE_NAME) 'Поиск контейнера
Set adUser = ad.FindObj(adContainer, AD_OBJECT_USER, "Petrov-IS", AD_ATTRIBUTE_SAMACCOUNTNAME) 
```

### Получение значения атрибута объекта
```
Dim ad As New ADConnector("MyOrg.org", "", "")
Dim adUser As ADUser

Set adUser = ad.FindObj(Nothing, AD_OBJECT_USER, "Ivanov-II", AD_ATTRIBUTE_SAMACCOUNTNAME)
MsgBox ad.Get("Department")
```

### Получение родительского контейнера
```
Dim ad As New ADConnector("MyOrg.org", "", "")
Dim adUser As ADUser
Dim adContainer As ADOU

Set adUser = ad.FindObj(Nothing, AD_OBJECT_USER, "Ivanov-II", AD_ATTRIBUTE_SAMACCOUNTNAME)
Set adContainer = adUser.GetParent()
MsgBox adContainer.ToString()
```

### Изменение значения атрибута объекта
```
Dim ad As New ADConnector("MyOrg.org", "", "")
Dim adUser As ADUser

Set adUser = ad.FindObj(Nothing, AD_OBJECT_USER, "Ivanov-II", AD_ATTRIBUTE_SAMACCOUNTNAME)
Call adUser.ModifyAttribute(ADS_PROPERTY_CLEAR, "Department", "")
Call adUser.ModifyAttribute(ADS_PROPERTY_UPDATE, "Company", "Live Scripts")
Call adUser.Save()
```

### Установка значения атрибута объекта
```
Dim ad As New ADConnector("MyOrg.org", "", "")
Dim adUser As ADUser

Set adUser = ad.FindObj(Nothing, AD_OBJECT_USER, "Ivanov-II", AD_ATTRIBUTE_SAMACCOUNTNAME)
Call adUser.Put("Company", "Live Scripts")
Call adUser.Save()
```

### Перенос объекта в контейнер
```
Dim ad As New ADConnector("MyOrg.org", "", "")
Dim adUser As ADUser
Dim adContainer As ADOU

Set adUser = ad.FindObj(Nothing, AD_OBJECT_USER, "Ivanov-II", AD_ATTRIBUTE_SAMACCOUNTNAME)
Set adContainer = ad.FindObj(Nothing, AD_OBJECT_OU, "HR Department", AD_ATTRIBUTE_NAME)
MsgBox adContainer.MoveHere(adUser, "")
```

### Добавление пользователя в группу
```
Dim ad As New ADConnector("MyOrg.org", "", "")
Dim adUser As ADUser
Dim adGroup As ADGroup

Set adUser = ad.FindObj(Nothing, AD_OBJECT_USER, "Ivanov-II", AD_ATTRIBUTE_SAMACCOUNTNAME)
Set adGroup = ad.FindObj(Nothing, AD_OBJECT_Group, "Group Name", AD_ATTRIBUTE_NAME)
MsgBox adGroup.AddMember(adUser)
```

