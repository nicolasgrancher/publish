<?php

declare(strict_types=1);

namespace Publish;

/**
 * Interface TemplateProcessorInterface
 *
 * @package Publish
 */
interface TemplateProcessorInterface
{
    public function open(string $filename): void;

    /**
     * @return array<string>
     */
    public function getVariablesNames(): array;

    public function setVariable(string $name, string $value): void;
}
