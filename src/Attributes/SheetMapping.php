<?php

namespace SheetMapper\Attributes;

use Attribute;
use SheetMapper\Exception\SheetMapperException;

#[Attribute(Attribute::TARGET_CLASS)]
class SheetMapping
{
    public function __construct(
        public readonly ?string $target_sheet = null,
        public readonly bool $has_header_row = false,
        public readonly bool $enforce_field_mapping = false,
        /**
         * @var list<int|string>
         */
        public readonly array $ignored_columns = [],
    ) {
        foreach ($this->ignored_columns as $skip) {
            if (!is_int($skip) && !is_string($skip)) {
                throw new SheetMapperException('Ignored columns must be integers (column indexes) or strings (header names).');
            }
        }
    }
}
