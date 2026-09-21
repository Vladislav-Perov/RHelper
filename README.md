
<div align="center">

![VBA](https://img.shields.io/badge/VBA-Office_Macros-1F6FEB?style=for-the-badge&logo=microsoftoffice&logoColor=white)
![WPF](https://img.shields.io/badge/WPF-.NET_10-512BD4?style=for-the-badge&logo=dotnet&logoColor=white)

# 📝 Report Helper

**Надёжный помощник для заполнения отчётов в MS Office Word**

Меньше рутины, больше времени на содержание.

[Что это](#what) · [Как пользоваться](#usage) · [Установка](#install) · [Структура](#structure)

</div>

---

<a id="what"></a>
## 💡 Что это
> **Макросы** — встроенные в в среду MS Office мини-программы на интегрированном VBA, автоматизирующие рутину.

**Report Helper** берёт на себя однотипную работу при оформлении отчётов в Microsoft Word: подставляет данные, заполняет шаблоны и избавляет от ручного копирования.

| | |
|---|---|
| 🎯 **Задача** | Быстрое и аккуратное заполнение отчётов |
| 🧩 **Платформа** | MS Office Word |
| ⚙️ **Два режима** | [WPF-клиент](#wpf) или [отдельные макросы](#macros) |

[↑ К началу](#top)

---

<a id="usage"></a>
## 🚀 Как пользоваться

Выберите удобный способ.

<a id="wpf"></a>
### 🖥 Способ 1. Через WPF-клиент

Подходит, если нужен удобный интерфейс без работы с кодом.

1. Скачайте последнюю версию из раздела [Releases](https://github.com/YOUR_USER/Report-Helper/releases)
2. Запустите `ReportHelper.exe`
3. ...

Требования к окружению описаны в разделе [Установка](#install).

<a id="macros"></a>
### 📜 Способ 2. Через отдельные макросы

Подходит, если нужно быстро настроить среду под себя.

1. Откройте Word и нажмите `Alt + F11`, чтобы открыть редактор VBA
2. Импортируйте модуль: **File → Import File…** и выберите необходимый модуль, например [`Modules/RTemplates.bas`](Modules/RTemplates.bas)
3. Вернитесь в документ и запустите макрос через `Alt + F8`

<details>
<summary>📋 Список доступных макросов</summary>

| Макрос | Что делает |
|---|---|
| `MacroName1` | Описание |
| `MacroName2` | Описание |

</details>

[↑ К началу](#top)

---

<a id="install"></a>
## 📦 Установка

```bash
git clone https://github.com/Vladislav-Perov/RHelper.git
```

**Требования:**
- Microsoft Word (версия 2016 и новее)
- .NET 10 Runtime (только для [WPF-клиента](#wpf))
- Разрешённые [макросы](#macros) в Word

[↑ К началу](#top)

---
<a id="structure"></a>
## 🗂 Структура проекта

```
Report-Helper/
├── Desktop/        # WPF-клиент
├── Modules/        # VBA-макросы
└── README.md
```

[↑ К началу](#top)

---

<div align="center">

Если проект оказался полезным, поставьте ⭐

</div>