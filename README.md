<div align="center">
  
![.NET](https://img.shields.io/badge/.NET-8.0-512BD4?style=for-the-badge&logo=dotnet&logoColor=white)
![WPF](https://img.shields.io/badge/WPF-Windows%2520Desktop-0078D4?style=for-the-badge&logo=windows&logoColor=white)
![C#](https://img.shields.io/badge/C%2523-239120?style=for-the-badge&logo=c-sharp&logoColor=white)

</div>

<div align="center">
  
![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=flat-square&logo=microsoft-excel&logoColor=white)
![Microsoft Word](https://img.shields.io/badge/Microsoft_Word-2B579A?style=flat-square&logo=microsoft-word&logoColor=white)
![OpenXML](https://img.shields.io/badge/OpenXML-00A4EF?style=flat-square&logo=microsoft&logoColor=white)

</div>

# Excel to Word Converter

Приложение для конвертации файлов Excel в документы Word с графиком сдачи зачетов и экзаменов.

## Описание

Приложение позволяет:
- Выбирать несколько файлов Excel для обработки
- Автоматически определять год начала подготовки из файлов
- Генерировать документы Word с таблицами зачетов и экзаменов по семестрам
- Отображать прогресс конвертации

## Технологии

- **WPF** - для пользовательского интерфейса
- **MVVM** - архитектура приложения (CommunityToolkit.Mvvm)
- **EPPlus** - работа с Excel файлами
- **OpenXML** - создание Word документов
- **FontAwesome.Sharp** - иконки в интерфейсе

## Установка и запуск

### Требования
- .NET 8.0 или выше
- Windows OS

### Сборка
1. Клонируйте репозиторий
2. Восстановите NuGet пакеты:
   ```bash
   dotnet restore
   ```
3. Соберите проект:
   ```bash
   dotnet build
   ```
4. Запустите приложение:
   ```bash
   dotnet run
   ```

## Особенности

### Автоматическое определение семестров
Приложение автоматически определяет семестры для отображения на основе года начала подготовки:

| Разница лет | Семестры |
|-------------|----------|
| 0           | 1, 2     |
| 1           | 3, 4     |
| 2           | 5, 6     |
| 3           | 7, 8     |

### Поддерживаемые колонки в Excel
- **Наименование** - название дисциплины
- **Зачет** - семестры для зачетов
- **Зачет с оценкой** - семестры для зачетов с оценкой  
- **КП** - семестры для курсовых проектов
- **Экзамен** - семестры для экзаменов

### Формат выходного документа
Генерируется Word документ с:
- Заголовком "ГРАФИК СДАЧИ ЗАЧЕТОВ И ЭКЗАМЕНОВ"
- Годом начала подготовки
- Таблицами для каждого семестра с двумя колонками:
  - Зачеты
  - Экзамены

## Структура проекта

```
ExcelToWordConverter/
├── Views/
│   └── MainWindow.xaml          # Главное окно
├── Models/
│   └── ExamConverter.cs         # Логика конвертации
├── ViewModels/
│   └── MainViewModel.cs         # MVVM ViewModel
├── Converters/
│   └── FileNameConverter.cs     # Конвертер для отображения имен файлов
└── Resources/                   # Ресурсы приложения
```

## Лицензия

Проект использует EPPlus под Non-Commercial лицензией. Для коммерческого использования требуется соответствующая лицензия.
