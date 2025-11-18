<?php

namespace SheetMapper\Schema;

class ClassSchema
{
    /**
     * @param FieldDefinition[] $fields
     */
    public function __construct(
        public readonly string $class_name,
        public readonly ?string $target_sheet,
        public readonly bool $has_header_row,
        public readonly bool $enforce_field_mapping,
        /**
         * @var list<int|string>
         */
        public readonly array $ignored_columns,
        public readonly array $fields,
    ) {
    }
}
