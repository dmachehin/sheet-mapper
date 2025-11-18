<?php

namespace SheetMapper\Tests;

use DateTimeImmutable;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date as SpreadsheetDate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PHPUnit\Framework\TestCase;
use SheetMapper\Attributes\SheetField;
use SheetMapper\Attributes\SheetMapping;
use SheetMapper\Exception\SheetMapperException;
use SheetMapper\SheetMapper;

class MapperTest extends TestCase
{
    private array $tempFiles = [];

    protected function tearDown(): void
    {
        foreach ($this->tempFiles as $file) {
            if (is_file($file)) {
                unlink($file);
            }
        }

        $this->tempFiles = [];
    }

    public function testMapsCsvWithHeaders(): void
    {
        $file = $this->createCsvFile([
            ['Name', 'Amount', 'Active'],
            ['Apple', '10.5', 'true'],
            ['Pear', '7', 'false'],
            ['', '', ''],
        ]);

        $mapper = new SheetMapper();
        $items = $mapper->map($file, CsvItem::class);

        self::assertCount(2, $items);
        self::assertSame('Apple', $items[0]->name);
        self::assertSame(10.5, $items[0]->amount);
        self::assertTrue($items[0]->active);
        self::assertFalse($items[1]->active);
    }

    public function testMapFromArrayWithHeaders(): void
    {
        $rows = [
            ['Name', 'Amount', 'Active'],
            ['Apple', '10.5', 'true'],
            ['Pear', '7', 'false'],
        ];

        $mapper = new SheetMapper();
        $items = $mapper->mapFromArray($rows, CsvItem::class);

        self::assertCount(2, $items);
        self::assertSame('Apple', $items[0]->name);
        self::assertSame(10.5, $items[0]->amount);
        self::assertFalse($items[1]->active);
    }

    public function testMapsXlsxWithColumnIndexesAndDates(): void
    {
        $first = new DateTimeImmutable('2024-01-01 12:34:56');
        $second = new DateTimeImmutable('2024-02-03 08:00:00');

        $file = $this->createXlsxFile([
            ['Orange', SpreadsheetDate::PHPToExcel($first)],
            ['Banana', SpreadsheetDate::PHPToExcel($second)],
        ]);

        $mapper = new SheetMapper();
        $items = $mapper->map($file, ColumnItem::class);

        self::assertCount(2, $items);
        self::assertInstanceOf(DateTimeImmutable::class, $items[0]->purchasedAt);
        self::assertSame($first->format(DATE_ATOM), $items[0]->purchasedAt->format(DATE_ATOM));
        self::assertSame($second->format(DATE_ATOM), $items[1]->purchasedAt->format(DATE_ATOM));
    }

    public function testMapFromArrayWithoutHeaders(): void
    {
        $first = new DateTimeImmutable('2024-01-01 12:34:56');
        $second = new DateTimeImmutable('2024-02-03 08:00:00');

        $rows = [
            ['Orange', $first->format('Y-m-d H:i:s')],
            ['Banana', $second->format('Y-m-d H:i:s')],
        ];

        $mapper = new SheetMapper();
        $items = $mapper->mapFromArray($rows, ColumnItem::class);

        self::assertCount(2, $items);
        self::assertSame('Orange', $items[0]->name);
        self::assertInstanceOf(DateTimeImmutable::class, $items[0]->purchasedAt);
        self::assertSame($first->format(DATE_ATOM), $items[0]->purchasedAt->format(DATE_ATOM));
        self::assertSame($second->format(DATE_ATOM), $items[1]->purchasedAt->format(DATE_ATOM));
    }

    public function testValidateFieldsMissingColumnThrows(): void
    {
        $file = $this->createCsvFile([
            ['OnlyOneColumn'],
        ]);

        $mapper = new SheetMapper();
        $this->expectException(SheetMapperException::class);
        $this->expectExceptionMessage('Column index 1');
        $mapper->map($file, MissingColumnItem::class);
    }

    public function testValidateFieldsSkipAllowsMissingHeader(): void
    {
        $file = $this->createCsvFile([
            ['Name'],
            ['Apple'],
        ]);

        $mapper = new SheetMapper();
        $items = $mapper->map($file, SkipHeaderItem::class);

        self::assertCount(1, $items);
        self::assertSame('Apple', $items[0]->name);
        self::assertNull($items[0]->optional);
    }

    public function testHeaderRegexpMatches(): void
    {
        $file = $this->createCsvFile([
            ['Product Name', 'Amount'],
            ['Widget', '42'],
        ]);

        $mapper = new SheetMapper();
        $items = $mapper->map($file, RegexHeaderItem::class);

        self::assertCount(1, $items);
        self::assertSame('Widget', $items[0]->name);
        self::assertSame(42, $items[0]->amount);
    }

    public function testMapsEnumFields(): void
    {
        $file = $this->createCsvFile([
            ['Type', 'State'],
            ['fruit', 'Published'],
            ['vegetable', 'Draft'],
        ]);

        $mapper = new SheetMapper();
        $items = $mapper->map($file, EnumItem::class);

        self::assertSame(ItemType::Fruit, $items[0]->type);
        self::assertSame(ItemState::Published, $items[0]->state);
        self::assertSame(ItemType::Vegetable, $items[1]->type);
        self::assertSame(ItemState::Draft, $items[1]->state);
    }

    public function testValueRegexpValidatesValues(): void
    {
        $file = $this->createCsvFile([
            ['Code'],
            ['ABC-123'],
        ]);

        $mapper = new SheetMapper();
        $items = $mapper->map($file, ValueRegexItem::class);

        self::assertCount(1, $items);
        self::assertSame('ABC-123', $items[0]->code);
    }

    public function testValueRegexpThrowsOnMismatch(): void
    {
        $file = $this->createCsvFile([
            ['Code'],
            ['invalid'],
        ]);

        $mapper = new SheetMapper();
        $this->expectException(SheetMapperException::class);
        $this->expectExceptionMessage('does not match pattern');
        $mapper->map($file, ValueRegexItem::class);
    }

    /**
     * @param list<list<string>> $rows
     */
    private function createCsvFile(array $rows): string
    {
        $file = $this->tempFilePath('csv');
        $handle = fopen($file, 'w');

        foreach ($rows as $row) {
            fputcsv($handle, $row, ',', '"', '\\');
        }

        fclose($handle);

        return $file;
    }

    /**
     * @param list<list<int|float|string>> $rows
     */
    private function createXlsxFile(array $rows, string $sheetName = 'Sheet1'): string
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle($sheetName);

        foreach ($rows as $rowIndex => $row) {
            foreach ($row as $columnIndex => $value) {
                $address = Coordinate::stringFromColumnIndex($columnIndex + 1) . ($rowIndex + 1);
                $sheet->setCellValue($address, $value);
            }
        }

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $file = $this->tempFilePath('xlsx');
        $writer->save($file);
        $spreadsheet->disconnectWorksheets();

        return $file;
    }

    private function tempFilePath(string $extension): string
    {
        $tmp = tempnam(sys_get_temp_dir(), 'sheet_mapper_');
        $file = $tmp . '.' . $extension;
        rename($tmp, $file);
        $this->tempFiles[] = $file;

        return $file;
    }
}

#[SheetMapping(has_header_row: true)]
class CsvItem
{
    #[SheetField(header: 'Name')]
    public string $name;

    #[SheetField(header: 'Amount')]
    public float $amount;

    #[SheetField(header: 'Active')]
    public bool $active;
}

#[SheetMapping(target_sheet: 'Sheet1', has_header_row: false)]
class ColumnItem
{
    #[SheetField(column: 0)]
    public string $name;

    #[SheetField(column: 1)]
    public DateTimeImmutable $purchasedAt;
}

#[SheetMapping(has_header_row: false, enforce_field_mapping: true)]
class MissingColumnItem
{
    #[SheetField(column: 1)]
    public string $name;
}

#[SheetMapping(has_header_row: true, enforce_field_mapping: true, ignored_columns: ['optional'])]
class SkipHeaderItem
{
    #[SheetField(header: 'Name')]
    public string $name;

    #[SheetField(header: 'Optional')]
    public ?string $optional = null;
}

#[SheetMapping(has_header_row: true)]
class RegexHeaderItem
{
    #[SheetField(header_regexp: '/^product/i')]
    public string $name;

    #[SheetField(header: 'Amount')]
    public int $amount;
}

#[SheetMapping(has_header_row: true)]
class EnumItem
{
    #[SheetField(header: 'Type')]
    public ItemType $type;

    #[SheetField(header: 'State')]
    public ItemState $state;
}

#[SheetMapping(has_header_row: true)]
class ValueRegexItem
{
    #[SheetField(header: 'Code', value_regexp: '/^[A-Z]{3}-\d{3}$/')]
    public string $code;
}

enum ItemType: string
{
    case Fruit = 'fruit';
    case Vegetable = 'vegetable';
}

enum ItemState
{
    case Draft;
    case Published;
}
