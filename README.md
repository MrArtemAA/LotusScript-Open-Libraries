# Библиотеки для IBM Notes/Domino

## Описание
Открытые библиотеки на языке LotusScript для IBM Notes/Domino
* Язык: LotusScript
* Лицензия: [MIT](https://opensource.org/licenses/MIT)

Сборник библиотек с открытым исходным кодом на языке LotusScript.

Репозиторий содержит:
* Domino-базу в виде On Disk Project с библиотеками и кодом агента для импорта из DXL-файл
* Документацию по библиотекам
* Библиотеки в виде DXL-файлов

### [ActiveDirectory_Lib](docs/ru/active-directory-lib/main.md)
Библиотека для платформы IBM Notes/Domino, облегчающая взаимодействие и интеграцию с MS Active Directory
### [Oracle_Lib](docs/ru/oracle-lib/main.md)
Библиотека для платформы IBM Notes/Domino для заимодействия с СУБД Oracle через COM-объекты OO4O

## Быстрый старт
1. Перейдите в папку `\Public Libraries\Code\ScriptLibraries\`
2. Откройте lss-файл нужной вам библиотеки
3. Скопируйте ее содержимое
4. Создайте в Notes-приложении новую библиотеку и вставьте то, что скопировали

## Использование
Для получения всех библиотек и создания Domino-базы с ними:
1. Клонируйте репозиторий (ветку master)
2. В Domino Designer откройте Navigator `(Window -> Show Eclipse View -> Navagator)`
3. В Domino Designer создайте проект `(В окне навигатора клик правой кнопкой мыши -> New -> Project -> General -> Project)`
4. При создании проекта укажите путь до репозитория, до папки `Public Libraries`.
5. После импорта проекта, клик правой кнопкой мыши по проекту -> Team development -> Accosiate with New NSF -> Указываете сервер и название файлы новой БД

### Импорт библиотеки из DXL-файла
1. Создайте агент для импорта из DXL, пример агента лежит в базе. Код взят [отсюда](http://www-10.lotus.com/ldd/bpmpblog.nsf/dx/the-missing-dxl-import-menu-option?opendocument&comments)
2. Выберите нужную библиотеку из папке DXL, расположенной в корня репозитория.