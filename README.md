<div align="center">

<img src="src/W-Fix.App/Assets/icon.ico" width="96" alt="W-Fix Icon"/>

# W-Fix — Printer Troubleshooter

**Диагностика и исправление принтеров Windows — одним портативным файлом.**

[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue?logo=windows)](https://www.microsoft.com/windows)
[![.NET](https://img.shields.io/badge/.NET-8.0-purple?logo=dotnet)](https://dotnet.microsoft.com/)
[![Release](https://img.shields.io/badge/version-3.1.0--beta.1-blue)](https://github.com/OneDeadMachine-Dev/W-FIX/releases)
[![License](https://img.shields.io/badge/license-MIT-orange)](LICENSE)
[![Author](https://img.shields.io/badge/author-OneDeadMachine-red)](https://github.com/OneDeadMachine)

</div>

---

## 🖨️ Что такое W-Fix?

**W-Fix** — компактная WPF-утилита для системных администраторов и продвинутых пользователей Windows, которая автоматически диагностирует и устраняет типичные проблемы с принтерами.

Вместо ручного ковыряния в реестре, PowerShell и журналах событий — всё в один клик.

> **Portable.** Один `.exe` файл, без установки. Носи на флешке.

---

## ✨ Возможности

| # | Фиксер | Ошибка | Что делает |
|---|--------|--------|------------|
| 1 | **Перезапуск Spooler** | `0x00000008`, `spooler` | Стоп → очистка очереди → старт |
| 2 | **Error 0x0000011b** | `RPC Auth` | Патч реестра `RpcAuthnLevelPrivacyEnabled` (KB5005565) |
| 3 | **Error 0x00004005** | Ops failed | RPC + Point&Print + брандмауэр + SMB |
| 4 | **Error 0x00000709** | Default printer | Права реестра HKCU + `WScript.Network` |
| 5 | **Error 0x00000002** | File not found | Очистка `prtprocs` + `PendingFileRename` + P&P |
| 6 | **Error 0x0000007e** | DLL missing | Копирование `mscms.dll` + удаление BIDI-ключа |
| 7 | **Error 0x0000007b** | Invalid name | Удаление повреждённого драйвера + очистка spool |
| 8 | **Error 0x00000008** | Not enough memory | Диагностика памяти + очистка temp |
| 9 | **IPP Fixer** | `0xbcb`, `0xbcc` | Windows Feature + IPP Class Driver + порт 631 |
| 10 | **Сетевая диагностика** | Network / DNS | Ping / DNS / TCP-порты (read-only) |
| 11 | **Переустановка драйвера** | `0x0000007b`, driver | INF / UNC / Авто (с диалогом выбора) |
| 12 | **Принтер по умолчанию** | Default printer | Сброс реестра + `SetDefaultPrinter` |

### Плюс:
- 🌐 **Remote Repair Center v3** — несколько доменных ПК, preflight, инвентаризация, автоматический план, проверка и изолированный rollback
- 🔗 **Pair Repair Wizard v3.1** — ремонт общего принтера между двумя ПК: TLS pairing, общий диагноз, двусторонний rollback и offline fallback
- 🧠 **Диагностические правила** — вывод по наблюдаемым фактам, уверенность и официальный источник известной проблемы
- 🔐 **Credential Manager** — альтернативная доменная учётная запись без пароля в конфигурации, CLI или логах
- 📊 **Отчёты** — JSON/HTML для каждого запуска и обезличенный ZIP support bundle
- 🌐 **Удалённые машины** — системные шаги через WinRM, а пользовательские настройки принтера — в реальном интерактивном сеансе
- 🔍 **Обнаружение принтеров** — CIM / WMI / Get-Printer (тройной fallback)
- 📋 **Полный лог** — каждый шаг фиксера отображается в реальном времени
- ↩️ **Точечный откат** — снимок изменяемых ключей реестра и кнопка восстановления после фикса
- 🗂 **Active Directory** — обзор принтеров в домене

---

## 🚀 Быстрый старт

1. Скачай `W-Fix.exe` из раздела [Releases](https://github.com/OneDeadMachine-Dev/W-FIX/releases)
2. Запусти **от имени администратора** (правой кнопкой → «Запуск от имени администратора»)
3. Выбери принтер в левой панели
4. Выбери фиксер в правой панели → нажми **«Применить»**

Для нескольких компьютеров открой **Remote Center**:

1. Добавь имена вручную либо через поиск Active Directory.
2. Выполни **Preflight** — ping носит справочный характер, рабочее подключение определяется по WinRM.
3. Запусти **Диагностику**, проверь факты и сформированный план.
4. Подтверди пакетный ремонт. Сбой одной машины не останавливает остальные; обратимые шаги неуспешной цели откатываются.

Если общий принтер подключён к другому ПК, открой **Pair Repair** на обеих машинах. Хост создаёт одноразовый `.wfixpair`, оба пользователя сверяют код, а клиент формирует и выполняет общий план. Подробная инструкция: [Pair Repair на русском](docs/pair-repair.md) / [English guide](docs/pair-repair.en.md).

> ⚠️ Права администратора обязательны — фиксеры изменяют реестр и службы Windows.

---

## 🛠 Сборка из исходников

### Требования
- Windows 10/11 x64
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

### Debug-запуск
```powershell
git clone https://github.com/OneDeadMachine-Dev/W-FIX.git
cd W-Fix
dotnet run --project src/W-Fix.App
```

### Портативный exe
```powershell
dotnet publish src/W-Fix.App/W-Fix.App.csproj `
  -c Release -r win-x64 --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeAllContentForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true `
  -o ./publish
```

Результат: `publish/W-Fix.exe` (~90 МБ, полностью автономный).

---

## 🏗 Архитектура

```
W-Fix/
├── src/
│   ├── W-Fix.Core/               # Бизнес-логика
│   │   ├── Abstractions/          # Контракты сессии, диагностики, ремонта и отчётов
│   │   ├── Remote/                # WinRM preflight и инвентаризация
│   │   ├── Diagnostics/           # Evidence-based правила
│   │   ├── Repair/                # Планировщик, legacy-адаптер и batch executor
│   │   ├── Pairing/               # TLS pairing, SMB/RPC diagnostics и двухузловая saga
│   │   ├── Catalog/               # Подписанный каталог Windows known issues
│   │   ├── Fixers/               # 12 фиксеров (FixerBase → IFixer)
│   │   ├── Services/
│   │   │   ├── WmiService.cs     # Обнаружение принтеров (CIM/WMI/PS)
│   │   │   ├── PowerShellEngine.cs  # Встроенный PS SDK + внешний fallback
│   │   │   ├── InteractiveUserPowerShellService.cs # Remote-команды в сеансе пользователя
│   │   │   ├── SystemStateBackupService.cs # Снимок реестра + restore.ps1
│   │   │   └── FixerRegistry.cs  # Регистрация и поиск фиксеров
│   │   └── Models/               # PrinterInfo, FixResult, LogEntry ...
│   └── W-Fix.App/                # WPF UI (MVVM + CommunityToolkit)
│       ├── ViewModels/
│       ├── RemoteCenterWindow.xaml
│       ├── PairRepairWindow.xaml
│       ├── Views/
│       └── Assets/icon.ico
└── publish/W-Fix.exe             # Готовый портативный файл
```

**Стек:** WPF · .NET 8 · ModernWpfUI · CommunityToolkit.Mvvm · PowerShell SDK · Serilog

Подробности внутренних границ, выполнения PowerShell и правил для фиксеров:
[docs/architecture.md](docs/architecture.md).

Работа с новым центром: [docs/remote-center.md](docs/remote-center.md). Формат и подпись базы проблем:
[docs/known-issues-catalog.md](docs/known-issues-catalog.md).

---

## 📋 Системные требования

| Компонент | Минимум |
|-----------|---------|
| ОС | Windows 10 1903+ / Windows 11 |
| Архитектура | x64 |
| .NET Runtime | Встроен (self-contained) |
| Права | Администратор |
| PowerShell | 5.1+ (встроен в Windows) |

---

## 📄 Лицензия

MIT © 2026 [OneDeadMachine](https://github.com/OneDeadMachine)

---

<div align="center">
<sub>Сделано с ❤️ для тех, кто устал объяснять пользователям как перезапустить Spooler</sub>
</div>
