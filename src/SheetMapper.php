<?php

namespace SheetMapper;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Shared\Date as SpreadsheetDate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use SheetMapper\Exception\SheetMapperException;
use SheetMapper\Schema\ClassSchema;
use SheetMapper\Schema\FieldDefinition;
use SheetMapper\Schema\SchemaResolver;

class SheetMapper
{
    private const DATETIME_SERIALIZATION_FORMAT = 'Y-m-d\TH:i:s.uP';

    public function __construct(
        private readonly SchemaResolver $schema_resolver = new SchemaResolver(),
    ) {
    }

    /**
     * @template T of object
     *
     * @param class-string<T> $class_name
     * @return T[]
     */
    public function map(string $file_path, string $class_name): array
    {
        if (!is_file($file_path)) {
            throw new SheetMapperException(sprintf('File "%s" was not found.', $file_path));
        }

        $schema = $this->schema_resolver->fromClass($class_name);
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file_path);

        try {
            $worksheet = $this->resolveWorksheet($spreadsheet, $schema->target_sheet);
            return $this->mapWorksheet($worksheet, $schema);
        } finally {
            $spreadsheet->disconnectWorksheets();
        }
    }

    /**
     * @template T of object
     *
     * @param class-string<T> $class_name
     * @param list<array<int|string, mixed>> $rows
     * @return T[]
     */
    public function mapFromArray(array $rows, string $class_name): array
    {
        $schema = $this->schema_resolver->fromClass($class_name);
        $spreadsheet = new Spreadsheet();

        try {
            $worksheet = $spreadsheet->getActiveSheet();
            if ($schema->target_sheet !== null) {
                $worksheet->setTitle($schema->target_sheet);
            }

            $this->populateWorksheet($worksheet, $rows);

            return $this->mapWorksheet($worksheet, $schema);
        } finally {
            $spreadsheet->disconnectWorksheets();
        }
    }

    /**
     * @return list<object>
     */
    private function mapWorksheet(Worksheet $worksheet, ClassSchema $schema): array
    {
        $header_data = $schema->has_header_row
            ? $this->buildHeaderMap($worksheet)
            : ['map' => [], 'raw' => []];

        if ($schema->enforce_field_mapping) {
            $this->validateFields($worksheet, $schema, $header_data);
        }

        $start_row = $schema->has_header_row ? 2 : 1;
        $highest_row = $worksheet->getHighestRow();

        $result = [];
        for ($row = $start_row; $row <= $highest_row; $row++) {
            if ($this->isRowEmpty($worksheet, $row, $schema, $header_data)) {
                continue;
            }

            try {
                $result[] = $this->hydrateObject($worksheet, $row, $schema, $header_data);
            } catch (SheetMapperException $exception) {
                throw $this->withRowContext($exception, $row);
            }
        }

        return $result;
    }

    private function withRowContext(SheetMapperException $exception, int $row): SheetMapperException
    {
        $message = sprintf('Row %d: %s', $row, $exception->getMessage());

        return new SheetMapperException($message, previous: $exception);
    }

    /**
     * @param list<array<int|string, mixed>> $rows
     */
    private function populateWorksheet(Worksheet $worksheet, array $rows): void
    {
        foreach (array_values($rows) as $row_index => $row) {
            if (!is_array($row)) {
                throw new SheetMapperException(sprintf('Row at index %d must be an array.', $row_index));
            }

            $next_column_index = 0;
            foreach ($row as $column_key => $value) {
                if (is_int($column_key)) {
                    if ($column_key < 0) {
                        throw new SheetMapperException(sprintf('Column indexes must be 0 or greater (row %d).', $row_index));
                    }
                    $column_index = $column_key;
                    $next_column_index = max($next_column_index, $column_index + 1);
                } else {
                    $column_index = $next_column_index;
                    $next_column_index++;
                }

                $coordinate = Coordinate::stringFromColumnIndex($column_index + 1) . ($row_index + 1);
                $worksheet->setCellValue($coordinate, $value);
            }
        }
    }

    private function resolveWorksheet(Spreadsheet $spreadsheet, ?string $target_sheet): Worksheet
    {
        if ($target_sheet === null) {
            return $spreadsheet->getActiveSheet();
        }

        $sheet = $spreadsheet->getSheetByName($target_sheet);
        if ($sheet === null) {
            throw new SheetMapperException(sprintf('Sheet "%s" was not found in workbook.', $target_sheet));
        }

        return $sheet;
    }

    /**
     * @return array{map: array<string, int>, raw: array<int, string>}
     */
    private function buildHeaderMap(Worksheet $worksheet): array
    {
        $map = [];
        $raw_headers = [];
        $max_column = Coordinate::columnIndexFromString($worksheet->getHighestColumn());

        for ($column = 1; $column <= $max_column; $column++) {
            $coordinate = Coordinate::stringFromColumnIndex($column) . '1';
            $value = $worksheet->getCell($coordinate)?->getValue();
            if ($value === null || $value === '') {
                continue;
            }
            $raw_value = (string) $value;
            $map[$this->normalizeHeader($raw_value)] = $column - 1;
            $raw_headers[$column - 1] = $raw_value;
        }

        return [
            'map' => $map,
            'raw' => $raw_headers,
        ];
    }

    /**
     * @param array{map: array<string, int>, raw: array<int, string>} $header_data
     */
    private function isRowEmpty(Worksheet $worksheet, int $row, ClassSchema $schema, array $header_data): bool
    {
        foreach ($schema->fields as $field) {
            $column_index = $this->resolveColumnIndex($field, $header_data, $schema);
            if ($column_index === null) {
                continue;
            }
            $value = $this->readCellValue($worksheet, $row, $column_index, $field->allow_merge);
            if ($value !== null && $value !== '') {
                return false;
            }
        }

        return true;
    }

    /**
     * @param array{map: array<string, int>, raw: array<int, string>} $header_data
     */
    private function hydrateObject(Worksheet $worksheet, int $row, ClassSchema $schema, array $header_data): object
    {
        /** @var class-string $class_name */
        $class_name = $schema->class_name;

        $field_values = [];

        foreach ($schema->fields as $field) {
            $column_index = $this->resolveColumnIndex($field, $header_data, $schema);
            if ($column_index === null) {
                continue;
            }

            $raw_value = $this->readCellValue($worksheet, $row, $column_index, $field->allow_merge);
            $this->assertValueMatchesPattern($raw_value, $field);
            $processed_value = $this->applyValueCallback($raw_value, $field);
            $value = $this->castValue($processed_value, $field);
            $field_values[$field->property] = $value;
        }

        $object = $this->hasPromotedFields($schema)
            ? $this->instantiateWithConstructorPromotion($class_name, $field_values)
            : new $class_name();

        foreach ($schema->fields as $field) {
            if (!array_key_exists($field->property, $field_values)) {
                continue;
            }

            if ($field->reflection_property->isPromoted()) {
                continue;
            }

            $this->assignValue($object, $field, $field_values[$field->property]);
        }

        return $object;
    }

    private function hasPromotedFields(ClassSchema $schema): bool
    {
        foreach ($schema->fields as $field) {
            if ($field->reflection_property->isPromoted()) {
                return true;
            }
        }

        return false;
    }

    /**
     * @param class-string $class_name
     * @param array<string, mixed> $field_values
     */
    private function instantiateWithConstructorPromotion(string $class_name, array $field_values): object
    {
        $class = new \ReflectionClass($class_name);
        $constructor = $class->getConstructor();

        if ($constructor === null) {
            return $class->newInstance();
        }

        $args = [];
        foreach ($constructor->getParameters() as $parameter) {
            if ($parameter->isPromoted()) {
                $property_name = $parameter->getName();
                if (array_key_exists($property_name, $field_values)) {
                    $args[] = $field_values[$property_name];
                    continue;
                }

                if ($parameter->isDefaultValueAvailable()) {
                    $args[] = $parameter->getDefaultValue();
                    continue;
                }

                if ($parameter->allowsNull()) {
                    $args[] = null;
                    continue;
                }

                throw new SheetMapperException(sprintf(
                    'Unable to resolve value for promoted property "%s" on "%s".',
                    $property_name,
                    $class_name,
                ));
            }

            if ($parameter->isDefaultValueAvailable()) {
                $args[] = $parameter->getDefaultValue();
                continue;
            }

            if ($parameter->allowsNull()) {
                $args[] = null;
                continue;
            }

            throw new SheetMapperException(sprintf(
                'Cannot instantiate "%s": constructor parameter "$%s" has no mapped value.',
                $class_name,
                $parameter->getName(),
            ));
        }

        return $class->newInstanceArgs($args);
    }

    /**
     * @param array{map: array<string, int>, raw: array<int, string>} $header_data
     */
    private function resolveColumnIndex(FieldDefinition $field, array $header_data, ClassSchema $schema): ?int
    {
        $header_map = $header_data['map'] ?? [];
        $raw_headers = $header_data['raw'] ?? [];

        if ($field->column !== null) {
            return $field->column;
        }

        if (!$schema->has_header_row) {
            throw new SheetMapperException(sprintf(
                'Field "%s" relies on header name "%s", but the sheet does not declare a header row.',
                $field->property,
                $field->header ?? 'n/a',
            ));
        }

        if ($field->header_regexp !== null) {
            $column = $this->findHeaderByPattern($field->header_regexp, $raw_headers);
            if ($column !== null) {
                return $column;
            }

            if ($schema->enforce_field_mapping && $this->isFieldSkipped($field, $schema)) {
                return null;
            }

            throw new SheetMapperException(sprintf(
                'Header matching pattern "%s" for field "%s" was not found.',
                $field->header_regexp,
                $field->property,
            ));
        }

        $header = $this->resolveHeaderName($field);
        $normalized = $this->normalizeHeader($header);
        if (!array_key_exists($normalized, $header_map)) {
            if ($schema->enforce_field_mapping && $this->isFieldSkipped($field, $schema)) {
                return null;
            }

            throw new SheetMapperException(sprintf('Header "%s" for field "%s" was not found.', $header, $field->property));
        }

        return $header_map[$normalized];
    }

    /**
     * @param array{map: array<string, int>, raw: array<int, string>} $header_data
     */
    private function validateFields(Worksheet $worksheet, ClassSchema $schema, array $header_data): void
    {
        $max_column_index = Coordinate::columnIndexFromString($worksheet->getHighestColumn()) - 1;
        $header_map = $header_data['map'] ?? [];
        $raw_headers = $header_data['raw'] ?? [];

        foreach ($schema->fields as $field) {
            if ($this->isFieldSkipped($field, $schema)) {
                continue;
            }

            if ($field->column !== null) {
                if ($field->column > $max_column_index) {
                    throw new SheetMapperException(sprintf(
                        'Column index %d for field "%s" was not found on sheet.',
                        $field->column,
                        $field->property,
                    ));
                }

                continue;
            }

            if ($field->header_regexp !== null) {
                if ($this->findHeaderByPattern($field->header_regexp, $raw_headers) === null) {
                    throw new SheetMapperException(sprintf(
                        'Header matching pattern "%s" for field "%s" was not found.',
                        $field->header_regexp,
                        $field->property,
                    ));
                }

                continue;
            }

            if (!$schema->has_header_row) {
                throw new SheetMapperException(sprintf(
                    'Field "%s" relies on header name "%s", but the sheet does not declare a header row.',
                    $field->property,
                    $field->header ?? $field->property,
                ));
            }

            $header = $this->resolveHeaderName($field);
            $normalized = $this->normalizeHeader($header);
            if (!array_key_exists($normalized, $header_map)) {
                throw new SheetMapperException(sprintf('Header "%s" for field "%s" was not found.', $header, $field->property));
            }
        }

        $this->validateUnexpectedHeaders($schema, $raw_headers);
    }

    /**
     * @param array<int, string> $raw_headers
     */
    private function findHeaderByPattern(string $pattern, array $raw_headers): ?int
    {
        foreach ($raw_headers as $column_index => $raw_header) {
            $match = @preg_match($pattern, $raw_header);
            if ($match === false) {
                throw new SheetMapperException(sprintf('Invalid header_regexp "%s".', $pattern));
            }
            if ($match === 1) {
                return $column_index;
            }
        }

        return null;
    }

    private function isFieldSkipped(FieldDefinition $field, ClassSchema $schema): bool
    {
        if (!$schema->enforce_field_mapping || $schema->ignored_columns === []) {
            return false;
        }

        foreach ($schema->ignored_columns as $skip) {
            if (is_int($skip) && $field->column !== null && $field->column === $skip) {
                return true;
            }

            if (is_string($skip)) {
                $field_header = $this->resolveHeaderName($field);
                if ($this->normalizeHeader($skip) === $this->normalizeHeader($field_header)) {
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * @param array<int, string> $raw_headers
     */
    private function validateUnexpectedHeaders(ClassSchema $schema, array $raw_headers): void
    {
        if (!$schema->has_header_row || $raw_headers === []) {
            return;
        }

        $mapped_columns = [];
        foreach ($schema->fields as $field) {
            if ($field->column !== null) {
                $mapped_columns[$field->column] = true;
                continue;
            }

            foreach ($raw_headers as $column_index => $raw_header) {
                if ($field->header_regexp !== null) {
                    $match = @preg_match($field->header_regexp, $raw_header);
                    if ($match === false) {
                        throw new SheetMapperException(sprintf('Invalid header_regexp "%s".', $field->header_regexp));
                    }

                    if ($match === 1) {
                        $mapped_columns[$column_index] = true;
                    }

                    continue;
                }

                $header = $this->resolveHeaderName($field);
                if ($this->normalizeHeader($raw_header) === $this->normalizeHeader($header)) {
                    $mapped_columns[$column_index] = true;
                }
            }
        }

        foreach ($raw_headers as $column_index => $raw_header) {
            if (isset($mapped_columns[$column_index])) {
                continue;
            }

            if ($this->isColumnIgnored($column_index, $raw_header, $schema)) {
                continue;
            }

            throw new SheetMapperException(sprintf(
                'Unexpected header "%s" at column index %d.',
                $raw_header,
                $column_index,
            ));
        }
    }

    private function isColumnIgnored(int $column_index, string $raw_header, ClassSchema $schema): bool
    {
        foreach ($schema->ignored_columns as $skip) {
            if (is_int($skip) && $skip === $column_index) {
                return true;
            }

            if (is_string($skip) && $this->normalizeHeader($skip) === $this->normalizeHeader($raw_header)) {
                return true;
            }
        }

        return false;
    }

    private function assertValueMatchesPattern(mixed $value, FieldDefinition $field): void
    {
        $pattern = $field->value_regexp;
        if ($pattern === null) {
            return;
        }

        if ($value === null || $value === '') {
            return;
        }

        if (is_array($value)) {
            throw new SheetMapperException(sprintf(
                'Value regexp can only be used with scalar or stringable values on field "%s".',
                $field->property,
            ));
        }

        if (is_object($value) && !method_exists($value, '__toString')) {
            throw new SheetMapperException(sprintf(
                'Value regexp can only be used with scalar or stringable values on field "%s".',
                $field->property,
            ));
        }

        $string_value = (string) $value;
        $match = @preg_match($pattern, $string_value);
        if ($match === false) {
            throw new SheetMapperException(sprintf('Invalid value_regexp "%s".', $pattern));
        }

        if ($match !== 1) {
            throw new SheetMapperException(sprintf(
                'Value "%s" for field "%s" does not match pattern "%s".',
                $string_value,
                $field->property,
                $pattern,
            ));
        }
    }

    private function resolveHeaderName(FieldDefinition $field): string
    {
        return $field->header !== null && $field->header !== ''
            ? $field->header
            : $field->property;
    }

    private function readCellValue(Worksheet $worksheet, int $row, int $column_index, bool $allow_merge = false): mixed
    {
        $coordinate = Coordinate::stringFromColumnIndex($column_index + 1) . $row;
        $cell = $worksheet->getCell($coordinate);

        if ($allow_merge && $cell->isInMergeRange() && !$cell->isMergeRangeValueCell()) {
            $merge_range = $cell->getMergeRange();
            if (is_string($merge_range)) {
                $ranges = Coordinate::splitRange($merge_range);
                [$master_coordinate] = $ranges[0];
                $cell = $worksheet->getCell($master_coordinate);
            }
        }

        return $cell?->getCalculatedValue();
    }

    private function normalizeHeader(string $value): string
    {
        return strtolower(trim($value));
    }

    private function applyValueCallback(mixed $value, FieldDefinition $field): mixed
    {
        $callback = $field->value_callback;
        if ($callback === null) {
            return $value;
        }

        if (!is_callable($callback)) {
            throw new SheetMapperException(sprintf('Value callback configured on field "%s" is not callable.', $field->property));
        }

        try {
            if (!$callback instanceof \Closure) {
                $callback = \Closure::fromCallable($callback);
            }

            return $callback($value);
        } catch (\Throwable $exception) {
            throw new SheetMapperException(sprintf(
                'Value callback failed for field "%s": %s',
                $field->property,
                $exception->getMessage(),
            ), previous: $exception);
        }
    }

    private function castValue(mixed $value, FieldDefinition $field): mixed
    {
        $type = $field->reflection_property->getType();

        if ($value === null/* || $value === ''*/) {
            if ($type !== null && !$type->allowsNull()) {
                throw new SheetMapperException(sprintf('Field "%s" is not nullable but cell is empty.', $field->property));
            }

            return null;
        }

        if ($type === null) {
            return $value;
        }

        if ($type instanceof \ReflectionUnionType) {
            throw new SheetMapperException(sprintf('Union types are not supported for field "%s".', $field->property));
        }

        if (!$type instanceof \ReflectionNamedType) {
            return $value;
        }

        $type_name = $type->getName();

        if ($type->isBuiltin()) {
            return $this->castToBuiltin($value, $type_name, $field);
        }

        $resolved_type_name = $this->resolveNamedType($type_name, $field);
        $is_enum = enum_exists($resolved_type_name);

        if (!$is_enum
            && !class_exists($resolved_type_name)
            && !interface_exists($resolved_type_name)
        ) {
            throw new SheetMapperException(sprintf(
                'Type "%s" was not found for field "%s".',
                $resolved_type_name,
                $field->property,
            ));
        }

        if ($is_enum) {
            return $this->castToEnum($value, $resolved_type_name, $field);
        }

        if (is_a($resolved_type_name, \DateTimeInterface::class, true)) {
            return $this->castToDateTime($value, $resolved_type_name, $field);
        }

        return $value;
    }

    private function resolveNamedType(string $type_name, FieldDefinition $field): string
    {
        return match ($type_name) {
            'self', 'static' => $field->reflection_property->getDeclaringClass()->getName(),
            'parent' => $this->resolveParentType($field),
            default => $type_name,
        };
    }

    private function resolveParentType(FieldDefinition $field): string
    {
        $parent = $field->reflection_property->getDeclaringClass()->getParentClass();
        if ($parent === false) {
            throw new SheetMapperException(sprintf(
                'Field "%s" references parent type but "%s" does not have a parent class.',
                $field->property,
                $field->reflection_property->getDeclaringClass()->getName(),
            ));
        }

        return $parent->getName();
    }

    /**
     * @param class-string<\UnitEnum> $enum_class
     */
    private function castToEnum(mixed $value, string $enum_class, FieldDefinition $field): \UnitEnum
    {
        if (!enum_exists($enum_class)) {
            throw new SheetMapperException(sprintf('Type "%s" is not an enum for field "%s".', $enum_class, $field->property));
        }

        if (is_subclass_of($enum_class, \BackedEnum::class)) {
            return $this->castToBackedEnum($value, $enum_class, $field);
        }

        $string_value = trim((string) $value);
        foreach ($enum_class::cases() as $case) {
            if ($case->name === $string_value) {
                return $case;
            }
        }

        throw new SheetMapperException(sprintf(
            'Value "%s" is not a valid case for enum "%s" on field "%s".',
            $string_value,
            $enum_class,
            $field->property,
        ));
    }

    /**
     * @param class-string<\BackedEnum> $enum_class
     */
    private function castToBackedEnum(mixed $value, string $enum_class, FieldDefinition $field): \BackedEnum
    {
        $reflection = new \ReflectionEnum($enum_class);
        $backing_type = $reflection->getBackingType()?->getName();

        $cast_value = $value;
        if ($backing_type === 'int') {
            $cast_value = (int) $value;
        } elseif ($backing_type === 'string') {
            $cast_value = (string) $value;
        }

        /** @var \BackedEnum|null $case */
        $case = $enum_class::tryFrom($cast_value);
        if ($case === null) {
            throw new SheetMapperException(sprintf(
                'Value "%s" is not a valid backed enum value for "%s" on field "%s".',
                (string) $value,
                $enum_class,
                $field->property,
            ));
        }

        return $case;
    }

    private function castToBuiltin(mixed $value, string $type_name, FieldDefinition $field): mixed
    {
        return match ($type_name) {
            'string' => (string) $value,
            'int' => (int) $value,
            'float' => (float) $value,
            'bool' => filter_var($value, FILTER_VALIDATE_BOOL, FILTER_NULL_ON_FAILURE) ?? (bool) $value,
            default => throw new SheetMapperException(sprintf('Unsupported builtin type "%s" on field "%s".', $type_name, $field->property)),
        };
    }

    private function castToDateTime(mixed $value, string $type_name, FieldDefinition $field): \DateTimeInterface
    {
        try {
            if (is_numeric($value)) {
                $date_time = SpreadsheetDate::excelToDateTimeObject((float) $value);

                return $this->instantiateDateTimeFromInterface($type_name, $date_time, $field);
            }

            if ($type_name === \DateTimeInterface::class) {
                return new \DateTimeImmutable((string) $value);
            }

            $instance = new $type_name((string) $value);
            if (!$instance instanceof \DateTimeInterface) {
                throw new SheetMapperException(sprintf('Type "%s" is not a DateTimeInterface implementation for field "%s".', $type_name, $field->property));
            }

            return $instance;
        } catch (\Exception $exception) {
            throw new SheetMapperException(sprintf(
                'Failed to cast value for field "%s" to "%s": %s',
                $field->property,
                $type_name,
                $exception->getMessage(),
            ), previous: $exception);
        }
    }

    private function instantiateDateTimeFromInterface(string $type_name, \DateTimeInterface $source, FieldDefinition $field): \DateTimeInterface
    {
        if ($type_name === \DateTimeInterface::class) {
            return \DateTimeImmutable::createFromInterface($source);
        }

        if ($type_name === \DateTimeImmutable::class || is_subclass_of($type_name, \DateTimeImmutable::class)) {
            $immutable = \DateTimeImmutable::createFromInterface($source);
            if ($type_name === \DateTimeImmutable::class) {
                return $immutable;
            }

            return new $type_name($immutable->format(self::DATETIME_SERIALIZATION_FORMAT));
        }

        if ($type_name === \DateTime::class || is_subclass_of($type_name, \DateTime::class)) {
            if ($source instanceof \DateTime && $type_name === \DateTime::class) {
                return clone $source;
            }

            return new $type_name($source->format(self::DATETIME_SERIALIZATION_FORMAT));
        }

        throw new SheetMapperException(sprintf('Unsupported DateTime type "%s" on field "%s".', $type_name, $field->property));
    }

    private function assignValue(object $object, FieldDefinition $field, mixed $value): void
    {
        $property = $field->reflection_property;
        if (!$property->isPublic()) {
            $property->setAccessible(true);
        }

        $property->setValue($object, $value);
    }
}
