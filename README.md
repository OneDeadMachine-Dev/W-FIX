<div align="center">

<img src="src/W-Fix.App/Assets/icon.ico" width="96" alt="W-Fix Icon"/>

# W-Fix — Printer Troubleshooter

**Диагностика и исправление принтеров Windows — одним портативным файлом.**

[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue?logo=windows)](https://www.microsoft.com/windows)
[![.NET](https://img.shields.io/badge/.NET-8.0-purple?logo=dotnet)](https://dotnet.microsoft.com/)
[![Release](https://img.shields.io/badge/version-2.2.0-green)](https://github.com/OneDeadMachine/W-Fix/releases)
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
- 🌐 **Удалённые машины** — применяй фиксеры через WinRM/PowerShell Remoting
- 🔍 **Обнаружение принтеров** — CIM / WMI / Get-Printer (тройной fallback)
- 📋 **Полный лог** — каждый шаг фиксера отображается в реальном времени
- 🗂 **Active Directory** — обзор принтеров в домене

---

## 🚀 Быстрый старт

1. Скачай `W-Fix.exe` из раздела [Releases](https://github.com/OneDeadMachine/W-Fix/releases)
2. Запусти **от имени администратора** (правой кнопкой → «Запуск от имени администратора»)
3. Выбери принтер в левой панели
4. Выбери фиксер в правой панели → нажми **«Применить»**

> ⚠️ Права администратора обязательны — фиксеры изменяют реестр и службы Windows.

---

## 🛠 Сборка из исходников

### Требования
- Windows 10/11 x64
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

### Debug-запуск
```powershell
git clone https://github.com/OneDeadMachine/W-Fix.git
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
│   │   ├── Fixers/               # 12 фиксеров (FixerBase → IFixer)
│   │   ├── Services/
│   │   │   ├── WmiService.cs     # Обнаружение принтеров (CIM/WMI/PS)
│   │   │   ├── PowerShellEngine.cs  # Встроенный PS SDK + внешний fallback
│   │   │   └── FixerRegistry.cs  # Регистрация и поиск фиксеров
│   │   └── Models/               # PrinterInfo, FixResult, LogEntry ...
│   └── W-Fix.App/                # WPF UI (MVVM + CommunityToolkit)
│       ├── ViewModels/
│       ├── Views/
│       └── Assets/icon.ico
└── publish/W-Fix.exe             # Готовый портативный файл
```

**Стек:** WPF · .NET 8 · ModernWpfUI · CommunityToolkit.Mvvm · PowerShell SDK · Serilog

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
