<?php

namespace SheetMapper\Attributes;

use Attribute;
use SheetMapper\Exception\SheetMapperException;

#[Attribute(Attribute::TARGET_PROPERTY)]
class SheetField
{
    public function __construct(
        public readonly ?int $column = null,
        public readonly ?string $header = null,
        public readonly ?string $header_regexp = null,
        public readonly bool $allow_merge = false,
        public readonly ?string $value_regexp = null,
        /**
         * @var callable|null
         */
        public readonly mixed $value_callback = null,
    ) {
        if ($column === null && ($header === null || $header === '') && ($header_regexp === null || $header_regexp === '')) {
            throw new SheetMapperException('Provide either column index, header value, or header_regexp for SheetField.');
        }
        if ($column !== null && $column < 0) {
            throw new SheetMapperException('Column index must be 0 or greater.');
        }
        if ($header_regexp !== null && $header_regexp !== '') {
            $this->assertValidPattern($header_regexp, 'header_regexp');
        }
        if ($value_regexp !== null && $value_regexp !== '') {
            $this->assertValidPattern($value_regexp, 'value_regexp');
        }
        if ($value_callback !== null && !is_callable($value_callback)) {
            throw new SheetMapperException('Value callback must be a valid callable reference.');
        }
    }

    private function assertValidPattern(string $pattern, string $label): void
    {
        set_error_handler(static function (int $errno, string $errstr) use ($pattern, $label): bool {
            throw new SheetMapperException(sprintf('Invalid %s "%s": %s', $label, $pattern, $errstr));
        });

        try {
            if (@preg_match($pattern, '') === false) {
                throw new SheetMapperException(sprintf('Invalid %s "%s".', $label, $pattern));
            }
        } finally {
            restore_error_handler();
        }
    }
}
