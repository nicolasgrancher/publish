<?php declare(strict_types=1);

namespace Tests\Publish;

use PHPUnit\Framework\TestCase;
use Publish\Exception\FileException;
use Publish\WordTemplateProcessor;

/**
 * Class WordTemplateProcessorTest
 *
 * @package Tests\Publish
 */
class WordTemplateProcessorTest extends TestCase
{
    const MODEL_SAMPLE_PATH = __DIR__ . '/Resources/model_sample.docx';

    /**
     * @throws FileException
     * @throws \ReflectionException
     */
    public function testOpen(): void
    {
        $object = new WordTemplateProcessor();
        $object->open(self::MODEL_SAMPLE_PATH);

        $reflectedClass = new \ReflectionClass(WordTemplateProcessor::class);
        $propertyDocumentHeaders = $reflectedClass->getProperty('documentHeaders');
        $propertyDocumentHeaders->setAccessible(true);
        $propertyDocumentFooters = $reflectedClass->getProperty('documentFooters');
        $propertyDocumentFooters->setAccessible(true);
        $propertyDocumentMainPart = $reflectedClass->getProperty('documentMainPart');
        $propertyDocumentMainPart->setAccessible(true);

        $this->assertNotEmpty($propertyDocumentHeaders->getValue($object));
        $this->assertNotEmpty($propertyDocumentFooters->getValue($object));
        $this->assertNotEmpty($propertyDocumentMainPart->getValue($object));
    }

    /**
     * @throws \ReflectionException
     */
    public function testGetHeaderName(): void
    {
        $object = new WordTemplateProcessor();
        $reflectedClass = new \ReflectionClass(WordTemplateProcessor::class);
        $method = $reflectedClass->getMethod('getHeaderName');
        $method->setAccessible(true);

        $this->assertEquals('word/header1.xml', $method->invokeArgs($object, [1]));
        $this->assertEquals('word/header5.xml', $method->invokeArgs($object, [5]));
    }

    /**
     * @throws \ReflectionException
     */
    public function testGetFooterName(): void
    {
        $object = new WordTemplateProcessor();
        $reflectedClass = new \ReflectionClass(WordTemplateProcessor::class);
        $method = $reflectedClass->getMethod('getFooterName');
        $method->setAccessible(true);

        $this->assertEquals('word/footer1.xml', $method->invokeArgs($object, [1]));
        $this->assertEquals('word/footer5.xml', $method->invokeArgs($object, [5]));
    }

    /**
     * @throws \ReflectionException
     */
    public function testGetMainPartName(): void
    {
        $object = new WordTemplateProcessor();
        $reflectedClass = new \ReflectionClass(WordTemplateProcessor::class);
        $method = $reflectedClass->getMethod('getMainPartName');
        $method->setAccessible(true);

        $this->assertEquals('word/document.xml', $method->invoke($object));
    }

    /**
     * @throws \ReflectionException
     */
    public function testGetVariablesNames(): void
    {
        $object = new WordTemplateProcessor();
        $reflectedClass = new \ReflectionClass(WordTemplateProcessor::class);

        $propertyDocumentHeaders = $reflectedClass->getProperty('documentHeaders');
        $propertyDocumentHeaders->setAccessible(true);
        $propertyDocumentHeaders->setValue($object, '<w:p w:rsidR="00606228" w:rsidRPr="008F1420" w:rsidRDefault="00606228" w:rsidP="00606228"><w:pPr><w:pStyle w:val="Variables"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>${variableHeader}</w:t></w:r></w:p>');
        $propertyDocumentFooters = $reflectedClass->getProperty('documentFooters');
        $propertyDocumentFooters->setAccessible(true);
        $propertyDocumentFooters->setValue($object, '<w:p w:rsidR="00606228" w:rsidRPr="008F1420" w:rsidRDefault="00606228" w:rsidP="00606228"><w:pPr><w:pStyle w:val="Variables"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>${variableFooter}</w:t></w:r></w:p>');
        $propertyDocumentMainPart = $reflectedClass->getProperty('documentMainPart');
        $propertyDocumentMainPart->setAccessible(true);
        $propertyDocumentMainPart->setValue($object, '<w:p w:rsidR="00606228" w:rsidRPr="008F1420" w:rsidRDefault="00606228" w:rsidP="00606228"><w:pPr><w:pStyle w:val="Variables"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>${variable1}</w:t></w:r></w:p><w:p w:rsidR="008E7B54" w:rsidRPr="008F1420" w:rsidRDefault="00681D69" w:rsidP="008E7B54"><w:r><w:t>Display</w:t></w:r><w:r w:rsidR="008E7B54" w:rsidRPr="008F1420"><w:t xml:space="preserve"> block B : </w:t></w:r><w:r w:rsidR="008E7B54" w:rsidRPr="008F1420"><w:rPr><w:rStyle w:val="VariablesCar"/></w:rPr><w:t>${variable2}</w:t></w:r></w:p>');

        $method = $reflectedClass->getMethod('getVariablesNames');
        $method->setAccessible(true);

        $this->assertEqualsCanonicalizing(['variableHeader', 'variable1', 'variable2', 'variableFooter'], $method->invoke($object));
    }

    /**
     * @throws \ReflectionException
     */
    public function testSetVariable(): void
    {
        $object = new WordTemplateProcessor();
        $reflectedClass = new \ReflectionClass(WordTemplateProcessor::class);

        $propertyDocumentHeaders = $reflectedClass->getProperty('documentHeaders');
        $propertyDocumentHeaders->setAccessible(true);
        $propertyDocumentHeaders->setValue($object, '<w:p w:rsidR="00606228" w:rsidRPr="008F1420" w:rsidRDefault="00606228" w:rsidP="00606228"><w:pPr><w:pStyle w:val="Variables"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>${variableHeader}</w:t></w:r></w:p>');
        $propertyDocumentFooters = $reflectedClass->getProperty('documentFooters');
        $propertyDocumentFooters->setAccessible(true);
        $propertyDocumentFooters->setValue($object, '<w:p w:rsidR="00606228" w:rsidRPr="008F1420" w:rsidRDefault="00606228" w:rsidP="00606228"><w:pPr><w:pStyle w:val="Variables"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>${variableFooter}</w:t></w:r></w:p>');
        $propertyDocumentMainPart = $reflectedClass->getProperty('documentMainPart');
        $propertyDocumentMainPart->setAccessible(true);
        $propertyDocumentMainPart->setValue($object, '<w:p w:rsidR="00606228" w:rsidRPr="008F1420" w:rsidRDefault="00606228" w:rsidP="00606228"><w:pPr><w:pStyle w:val="Variables"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>${variable1}</w:t></w:r></w:p><w:p w:rsidR="008E7B54" w:rsidRPr="008F1420" w:rsidRDefault="00681D69" w:rsidP="008E7B54"><w:r><w:t>Display</w:t></w:r><w:r w:rsidR="008E7B54" w:rsidRPr="008F1420"><w:t xml:space="preserve"> block B : </w:t></w:r><w:r w:rsidR="008E7B54" w:rsidRPr="008F1420"><w:rPr><w:rStyle w:val="VariablesCar"/></w:rPr><w:t>${variable2}</w:t></w:r></w:p>');

        $method = $reflectedClass->getMethod('setVariable');
        $method->setAccessible(true);
        $method->invokeArgs($object, ['variable1', 'my value']);

        $this->assertEquals(
            '<w:p w:rsidR="00606228" w:rsidRPr="008F1420" w:rsidRDefault="00606228" w:rsidP="00606228"><w:pPr><w:pStyle w:val="Variables"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>${variableHeader}</w:t></w:r></w:p>',
            $propertyDocumentHeaders->getValue($object)
        );
        $this->assertEquals(
            '<w:p w:rsidR="00606228" w:rsidRPr="008F1420" w:rsidRDefault="00606228" w:rsidP="00606228"><w:pPr><w:pStyle w:val="Variables"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>${variableFooter}</w:t></w:r></w:p>',
            $propertyDocumentFooters->getValue($object)
        );
        $this->assertEquals(
            '<w:p w:rsidR="00606228" w:rsidRPr="008F1420" w:rsidRDefault="00606228" w:rsidP="00606228"><w:pPr><w:pStyle w:val="Variables"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>my value</w:t></w:r></w:p><w:p w:rsidR="008E7B54" w:rsidRPr="008F1420" w:rsidRDefault="00681D69" w:rsidP="008E7B54"><w:r><w:t>Display</w:t></w:r><w:r w:rsidR="008E7B54" w:rsidRPr="008F1420"><w:t xml:space="preserve"> block B : </w:t></w:r><w:r w:rsidR="008E7B54" w:rsidRPr="008F1420"><w:rPr><w:rStyle w:val="VariablesCar"/></w:rPr><w:t>${variable2}</w:t></w:r></w:p>',
            $propertyDocumentMainPart->getValue($object)
        );
    }

    /**
     * @throws \ReflectionException
     */
    public function testFixBrokenTags(): void
    {
        $object = new WordTemplateProcessor();
        $reflectedClass = new \ReflectionClass(WordTemplateProcessor::class);
        $method = $reflectedClass->getMethod('fixBrokenTags');
        $method->setAccessible(true);

        $this->assertEquals(
            'Lorem ipsum <w:r><w:t>${var1}</w:t></w:r> dolor sit amet',
            $method->invokeArgs($object, [
                'Lorem ipsum <w:r><w:t>${</w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t>var1</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:t>}</w:t></w:r> dolor sit amet'
            ])
        );
    }

    /**
     * @throws \ReflectionException
     */
    public function testCleanString(): void
    {
        $object = new WordTemplateProcessor();
        $reflectedClass = new \ReflectionClass(WordTemplateProcessor::class);

        $method = $reflectedClass->getMethod('cleanString');
        $method->setAccessible(true);

        $this->assertEquals("my &lt; clean &gt; string\n<w:br/>&amp; on two lines", $method->invokeArgs($object, ["my < clean > string\n& on two lines"]));
    }

    /**
     * @throws \ReflectionException
     */
    public function testGetFullTag(): void
    {
        $object = new WordTemplateProcessor();
        $reflectedClass = new \ReflectionClass(WordTemplateProcessor::class);

        $method = $reflectedClass->getMethod('cleanString');
        $method->setAccessible(true);

        $this->assertEquals('${tag1}', $method->invokeArgs($object, ['tag1']));
    }

    /**
     * @throws \ReflectionException
     */
    public function testGetUtf8String(): void
    {
        $object = new WordTemplateProcessor();
        $reflectedClass = new \ReflectionClass(WordTemplateProcessor::class);

        $method = $reflectedClass->getMethod('getUtf8String');
        $method->setAccessible(true);

        $this->assertEquals('ma chaîne encodée', $method->invokeArgs($object, [iconv("UTF-8", "Windows-1252", 'ma chaîne encodée')]));
    }

    /**
     * @throws \ReflectionException
     */
    public function testSetVariableInXml(): void
    {
        $object = new WordTemplateProcessor();
        $reflectedClass = new \ReflectionClass(WordTemplateProcessor::class);

        $method = $reflectedClass->getMethod('setVariableInXml');
        $method->setAccessible(true);

        $this->assertEquals(
            '<w:p w:rsidR="00606228" w:rsidRPr="008F1420" w:rsidRDefault="00606228" w:rsidP="00606228"><w:pPr><w:pStyle w:val="Variables"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>my value</w:t></w:r></w:p><w:p w:rsidR="008E7B54" w:rsidRPr="008F1420" w:rsidRDefault="00681D69" w:rsidP="008E7B54"><w:r w:rsidR="008E7B54" w:rsidRPr="008F1420"><w:rPr><w:rStyle w:val="VariablesCar"/></w:rPr><w:t>${variable2}</w:t></w:r></w:p>',
            $method->invokeArgs($object, [
                'variable1',
                'my value',
                '<w:p w:rsidR="00606228" w:rsidRPr="008F1420" w:rsidRDefault="00606228" w:rsidP="00606228"><w:pPr><w:pStyle w:val="Variables"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>${variable1}</w:t></w:r></w:p><w:p w:rsidR="008E7B54" w:rsidRPr="008F1420" w:rsidRDefault="00681D69" w:rsidP="008E7B54"><w:r w:rsidR="008E7B54" w:rsidRPr="008F1420"><w:rPr><w:rStyle w:val="VariablesCar"/></w:rPr><w:t>${variable2}</w:t></w:r></w:p>',
            ])
        );
    }
}
