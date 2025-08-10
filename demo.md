краткое демо с объяснением логики обработки файлов на примере этих файлов:
| оригинал | то что увидит пользователь | дата |
|---|---|---|
| Call recording Alex Novak_250809_072805.m4a | Alex | 09.08.25 07:28:05 |
| Анна Викторовна Тест(0970000000)_20220201150604.mp3 | Анна | 01.02.22 15:06:04 |
| Мария Петровна Иванова_221109_094953.m4a | Мария | 09.11.22 09:49:53 |
| Ольга_Михайловна_Смирнова_0910001122_2017_08_19_15_52_03.ogg | Ольга | 19.08.17 15:52:03 |

по умолчанию открыается папка    
G:\My Drive\recordings   
если такой папки нету, то пользователь сам может выбрать нужную папку нажам на кнопку folder

показываю процесс инсталяции
- как скачать
- как создать ярлык
- прописать в ярлыке
    - консоль не сворачивается: powershell.exe -ExecutionPolicy Bypass -File "C:\path\to\FileManager.ps1"
    - консоль сворачивается: powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\path\to\FileManager.ps1"
- как узнать где стоит PowerShell $PSHOME и прописать полный путь вместо powershell.exe, на случай если PowerShell не прописан в path
- C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
- "C:\Program Files\PowerShell\7\pwsh.exe"

обратная связи:
- пользователи - каменты на ютюбе
- програмисты или продвинутые пользователи: issues, pull requests, descussions на https://github.com/SVDotsenko/cleaner