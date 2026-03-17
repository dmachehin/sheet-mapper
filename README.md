# Sheet Mapper

PHP библиотека, которая превращает строки CSV/XLS/XLSX в DTO-классы на основе атрибутов-схем.
Внутри используется PhpSpreadsheet, так что поддерживаются как таблицы с заголовками, так и файлы с фиксированными колонками.

## Возможности

- `#[SheetMapping]` на классе описывает лист и наличие строки заголовков.
- `#[SheetField]` на свойствах задаёт номер колонки (с нуля) либо имя заголовка.
- Поиск колонок по регулярным выражениям: `header_regexp` помогает, если заголовки динамические.
- Опциональная проверка значений по регулярному выражению через `value_regexp`.
- Кастомные преобразования значений перед приведением типов через `value_callback`.
- Проверка структуры листа: `enforce_field_mapping=true` гарантирует наличие всех обязательных колонок/заголовков и отсутствие лишних. По умолчанию `enforce_field_mapping=false`. `ignored_columns` позволяет пометить исключения (по индексу колонки или имени заголовка).
- Автоматическое приведение типов для `string`, `int`, `float`, `bool`, а также `DateTimeInterface` и enum.
- Пропуск полностью пустых строк и понятные исключения, если значения отсутствуют или заголовок не найден.

## Быстрый старт

DTO с атрибутами:

```php
<?php

use SheetMapper\Attributes\SheetMapping;
use SheetMapper\Attributes\SheetField;

#[SheetMapping(target_sheet: 'Sheet1', has_header_row: true)]
class Item
{
    #[SheetField(header: 'Name')]
    public string $name;

    #[SheetField(header: 'Amount')]
    public float $amount;

    #[SheetField(header: 'Active')]
    public bool $active;
}
```

Чтение файла:

```php
<?php

use SheetMapper\SheetMapper;

require __DIR__ . '/vendor/autoload.php';

$mapper = new SheetMapper();
$items = $mapper->map(__DIR__ . '/storage/items.xlsx', Item::class);

foreach ($items as $item) {
    echo $item->name . ' => ' . $item->amount . PHP_EOL;
}

// Если данные уже загружены, можно передать массив строк:
$rows = [
    ['Name', 'Amount', 'Active'],
    ['Apple', 10.5, true],
];

$items = $mapper->mapFromArray($rows, Item::class);
```

### Чтение без заголовков

Если в листе нет строки заголовков, установите `has_header_row: false` и укажите индекс колонки:

```php
#[SheetMapping(has_header_row: false)]
class ColumnItem
{
    #[SheetField(column: 0)]
    public string $name;

    #[SheetField(column: 1)]
    public DateTimeImmutable $purchasedAt;
}
```

### Проверка структуры листа

```php
#[SheetMapping(
    has_header_row: true,
    enforce_field_mapping: true,
    ignored_columns: ['optional', 5] // строка заголовка или индекс колонки
)]
class ValidatedItem
{
    #[SheetField(header: 'Name')]
    public string $name;

    #[SheetField(header: 'Optional')]
    public ?string $optional = null;

    #[SheetField(column: 5)]
    public ?float $legacyAmount = null;
}
```

Если обязательная колонка/заголовок не найден, будет выброшено `SheetMapperException`. Лишние заголовки тоже приводят к ошибке. Все элементы из `ignored_columns` допускается пропустить и/или оставить в файле как исключения — полю останется значение по умолчанию.

### Регулярные выражения для заголовков

```php
#[SheetMapping(has_header_row: true)]
class RegexItem
{
    #[SheetField(header_regexp: '/^product/i')]
    public string $name;

    #[SheetField(header: 'Amount')]
    public int $amount;
}
```

Поле `name` найдёт любую колонку, чей заголовок начинается со слова «Product». Можно комбинировать `header` и `header_regexp`, но как минимум одно из `column`, `header`, `header_regexp` должно быть задано.

### Регулярные выражения для значений

```php
#[SheetMapping(has_header_row: true)]
class ValueItem
{
    #[SheetField(header: 'Code', value_regexp: '/^[A-Z]{3}-\d{3}$/')]
    public string $code;
}
```

`value_regexp` проверяет каждое непустое значение и выбрасывает `SheetMapperException`, если оно не совпало с шаблоном.

### Кастомные преобразования значений

Если стандартного приведения типов недостаточно, используйте `value_callback` — он получает исходное значение ячейки и должен вернуть преобразованный результат (подойдёт любая callable: `'trim'`, `['ClassName', 'method']` и т.п.).

```php
final class FieldCallbacks
{
    public static function ruYesNoToBool(mixed $value): bool
    {
        return strtolower(trim((string) $value)) === 'да';
    }
}

#[SheetMapping(has_header_row: true)]
class CallbackItem
{
    #[SheetField(header: 'Active', value_callback: [FieldCallbacks::class, 'ruYesNoToBool'])]
    public bool $active;
}
```

При ошибке внутри callback будет выброшено `SheetMapperException`, так что можно смело кидать собственные исключения с пояснениями.
