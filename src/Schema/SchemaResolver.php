<?php

namespace SheetMapper\Schema;

use ReflectionClass;
use SheetMapper\Attributes\SheetField;
use SheetMapper\Attributes\SheetMapping as SheetMapperAttribute;
use SheetMapper\Exception\SheetMapperException;

class SchemaResolver
{
    public function fromClass(string $class_name): ClassSchema
    {
        if (!class_exists($class_name)) {
            throw new SheetMapperException(sprintf('Class "%s" was not found.', $class_name));
        }

        $class = new ReflectionClass($class_name);

        if ($class->isAbstract()) {
            throw new SheetMapperException(sprintf('Class "%s" must be instantiable.', $class_name));
        }

        $mapper_attribute = $this->resolveMapperAttribute($class);
        $target_sheet = $mapper_attribute?->target_sheet;
        $has_header_row = $mapper_attribute?->has_header_row ?? false;
        $enforce_field_mapping = $mapper_attribute?->enforce_field_mapping ?? false;
        $ignored_columns = $mapper_attribute?->ignored_columns ?? [];

        $fields = [];

        foreach ($class->getProperties() as $property) {
            $attributes = $property->getAttributes(SheetField::class);
            if ($attributes === []) {
                continue;
            }

            $attribute = $attributes[0]->newInstance();
            $header = $attribute->header !== null ? trim($attribute->header) : null;
            $header = $header === '' ? null : $header;
            $header_regexp = $attribute->header_regexp !== null ? trim($attribute->header_regexp) : null;
            $header_regexp = $header_regexp === '' ? null : $header_regexp;
            $value_regexp = $attribute->value_regexp !== null ? trim($attribute->value_regexp) : null;
            $value_regexp = $value_regexp === '' ? null : $value_regexp;
            $fields[] = new FieldDefinition(
                property: $property->getName(),
                column: $attribute->column,
                header: $header,
                reflection_property: $property,
                header_regexp: $header_regexp,
                value_regexp: $value_regexp,
                value_callback: $attribute->value_callback,
            );
        }

        if ($fields === []) {
            throw new SheetMapperException(sprintf('Class "%s" must declare at least one #[SheetField] attribute.', $class_name));
        }

        return new ClassSchema(
            class_name: $class_name,
            target_sheet: $target_sheet,
            has_header_row: $has_header_row,
            enforce_field_mapping: $enforce_field_mapping,
            ignored_columns: $ignored_columns,
            fields: $fields,
        );
    }

    private function resolveMapperAttribute(ReflectionClass $class): ?SheetMapperAttribute
    {
        $attributes = $class->getAttributes(SheetMapperAttribute::class);
        if ($attributes === []) {
            return null;
        }

        /** @var SheetMapperAttribute $instance */
        $instance = $attributes[0]->newInstance();

        return $instance;
    }
}
