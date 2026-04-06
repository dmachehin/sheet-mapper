<?php

namespace SheetMapper\Schema;

use ReflectionProperty;

class FieldDefinition
{
    public function __construct(
        public readonly string $property,
        public readonly ?int $column,
        public readonly ?string $header,
        public readonly ReflectionProperty $reflection_property,
        public readonly ?string $header_regexp,
        public readonly bool $allow_merge,
        public readonly ?string $value_regexp,
        /**
         * @var callable|null
         */
        public readonly mixed $value_callback,
    ) {
    }
}
