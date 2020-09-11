<?php

namespace Publish;

use Laminas\Stdlib\StringUtils;
use Publish\Exception\FileException;

/**
 * Class WordTemplateProcessor
 *
 * @package Publish
 */
class WordTemplateProcessor implements TemplateProcessorInterface
{
    const TAG_START = '${';
    const TAG_END = '}';

    /**
     * @var array<string>
     */
    protected $documentHeaders;

    /**
     * @var array<string>
     */
    protected $documentFooters;

    /**
     * @var string
     */
    protected $documentMainPart;

    /**
     * @param string $filename
     *
     * @throws FileException
     */
    public function open(string $filename): void
    {
        // Temporary document content extraction
        $zipClass = new \ZipArchive();
        if (true !== $res = $zipClass->open($filename)) {
            throw new FileException(sprintf('Could not open file "%s" (code %s)', $filename, $res));
        }

        $index = 1;
        while (false !== $zipClass->locateName($this->getHeaderName($index))) {
            $header = $zipClass->getFromName($this->getHeaderName($index));
            if (false !== $header) {
                $this->documentHeaders[$index] = $this->fixBrokenTags($header);
            }
            $index++;
        }

        $index = 1;
        while (false !== $zipClass->locateName($this->getFooterName($index))) {
            $footer = $zipClass->getFromName($this->getFooterName($index));
            if (false !== $footer) {
                $this->documentFooters[$index] = $this->fixBrokenTags($footer);
            }
            $index++;
        }

        $mainPart = $zipClass->getFromName($this->getMainPartName());
        if (false !== $mainPart) {
            $this->documentMainPart = $this->fixBrokenTags($mainPart);
        }
    }

    protected function getHeaderName(int $index): string
    {
        return sprintf('word/header%d.xml', $index);
    }

    protected function getFooterName(int $index): string
    {
        return sprintf('word/footer%d.xml', $index);
    }

    protected function getMainPartName(): string
    {
        return 'word/document.xml';
    }

    public function getVariablesNames(): array
    {
        preg_match_all('/\${.*}/', $this->documentMainPart, $variablesMainPart);

        return array_merge($variablesMainPart);
    }

    public function setVariable(string $name, string $value): void
    {
        $value = $this->cleanString($value);

        $search = $this->getFullTag($name);
        $replace = $this->getUtf8String($value);

        $this->documentHeaders = $this->setVariableInXml($search, $replace, $this->documentHeaders);
        $this->documentMainPart = $this->setVariableInXml($search, $replace, $this->documentMainPart);
        $this->documentFooters = $this->setVariableInXml($search, $replace, $this->documentFooters);
    }

    protected function fixBrokenTags(string $documentPart): string
    {
        /** @var string $fixedDocumentPart */
        $fixedDocumentPart = preg_replace_callback('|\$[^{]*\{[^}]*\}|U', function ($match) {
            return strip_tags($match[0]);
        }, $documentPart);

        return $fixedDocumentPart;
    }

    protected function cleanString(string $input): string
    {
        return \str_replace(['&', '<', '>', "\n"], ['&amp;', '&lt;', '&gt;', "\n<w:br/>"], $input);
    }

    protected function getFullTag(string $tag): string
    {
        if (substr($tag, 0, 2) !== '${' && substr($tag, -1) !== '}') {
            $tag = '${' . $tag . '}';
        }

        return $tag;
    }

    protected function getUtf8String(string $value): string
    {
        if (!StringUtils::isValidUtf8($value)) {
            $value = utf8_encode($value);
        }

        return $value;
    }

    /**
     * @template T
     *
     * @param string $name
     * @param string $value
     * @param T      $xmlContent
     *
     * @return T
     */
    protected function setVariableInXml(string $name, string $value, $xmlContent)
    {
        $name = $this->getFullTag($name);

        /** @var T $result */
        $result = preg_replace('/' . preg_quote($name, '/') . '/u', $value, $xmlContent);

        return $result;
    }
}
