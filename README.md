# IPUSN2ExcelTemplates

Плагин для программы ИП УСН2 для печати документов в формате XLSX с QR-кодом (опционально) на основе Excel-шаблонов

Поддерживает документы:
- счет (для юр лиц) - на основе [шаблона xlsx](ExcelTemplates/bin/Debug/Templates)
- счет с QRCode (для физ лиц) - на основе [шаблона xlsx](ExcelTemplates/bin/Debug/Templates) (формируется при отстуствии ИНН и КПП покупателя)
- акт - на основе [шаблона xlsx](ExcelTemplates/bin/Debug/Templates)     
- ... [подробнее](ExcelTemplates/bin/Debug/Информация.txt)

Индивиуальные шаблоны для каждого вашего ИНН:
- счет (для юр лиц)
- счет с QRCode (для физ лиц) (используется при отстуствии ИНН и КПП покупателя)
- акт
- ... [подробнее](ExcelTemplates/bin/Debug/Информация.txt)

## Информация

**Установка**: 
- Распакуйте во временную папку и запустите: **INSTALL.cmd**    
- В меню пуск появится ярлык `ИП УСН2 ExcelTemplates - Настройки`    

**Ссылки**:
- [Сайт программы](https://ipusn.dynsoft.ru/)     
- [Сайт плагина](https://github.com/dkxce/IPUSN2ExcelTemplates)       
- [Релизы](https://github.com/dkxce/IPUSN2ExcelTemplates/releases)     
- [Архивы](Binaries)     

**Правка шаблонов**:
- До установки: Папка **Templates** архива    
- После установки: Подкаталог программы **<ПУТЬ>\Plugins\ExcelTemplate\Templates**
- После установки: Подкаталог программы **C:\IPUSN2\Plugins\ExcelTemplate\Templates**

**Доп Информация**:    
- [подробнее](ExcelTemplates/bin/Debug/Информация.txt)

**Примеры**:    
- [Примеры](Examples)         
![bill_example](Examples/bill_example.png)    
![act_example](Examples/act_example.png)    

**Настройки**:    
![config](Examples/configurate.png)    

## Что нового

**v1.0.2.7**:
- Добавлена поддержка картинок в шаблонах (со сдвигом относительно списка товаров/услуг)    
- Добавлены новый шаблоны    
- Добавлена опция выбора шаблона по-умолчанию на основе списка    
- Добавлена возможность вывода QR-кода на счетах для ИП    

**v1.0.2.9**:
- Добавлена поддержка DataMatrix кода (%MATRIX%)    
- Добавлена поддержка Code128 кода (%BARCODE%)    