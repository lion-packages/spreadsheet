<?php

declare(strict_types=1);

namespace Tests;

use Exception;
use LionSpreadsheet\Spreadsheet;
use LionTest\Test;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PhpSpreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Tests\Provider\SpreadsheetProviderTrait;

class SpreadsheetTest extends Test
{
    use SpreadsheetProviderTrait;

    const SAVE_PATH = './storage/';
    const SUPPORT_PATH = './tests/support-files/';
    const FILE_NAME = 'template.xlsx';
    const FILE_NAME_MULTIPLE_SHEETS = 'template-multiple-sheets.xlsx';
    const FILE_NAME_MULTIPLE_SHEETS_DATA_VALIDATION = 'template-multiple-sheets-data-validation.xlsx';
    const FILE_PATH = self::SUPPORT_PATH . self::FILE_NAME;
    const FILE_PATH_MULTIPLE_SHEETS = self::SUPPORT_PATH . self::FILE_NAME_MULTIPLE_SHEETS;
    const FILE_PATH_MULTIPLE_SHEETS_DATA_VALIDATION = self::SUPPORT_PATH . self::FILE_NAME_MULTIPLE_SHEETS_DATA_VALIDATION;
    const FILE_TYPE = 'fileType';
    const SPREADSHEET = 'spreadsheet';
    const WORKSHEET = 'worksheet';
    const CONTENT_TYPE = 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const CONTENT_DISPOSITION = 'Content-Disposition: attachment; filename=' . self::FILE_NAME;

    protected function setUp(): void
    {
        $this->createDirectory(self::SAVE_PATH);
    }

    protected function tearDown(): void
    {
        $this->rmdirRecursively(self::SAVE_PATH);
    }

    private function saveFile(Spreadsheet $spreadsheet, string $fileName): void
    {
        $fileName = $fileName . '-' . self::FILE_NAME;
        $spreadsheet->save(self::SAVE_PATH . $fileName);

        $this->assertFileExists(self::SAVE_PATH . $fileName);
    }

    public function testConstruct(): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH);
        $this->initReflection($spreadsheet);

        $this->assertSame(Spreadsheet::XLSX, $this->getPrivateProperty(self::FILE_TYPE));
        $this->assertInstanceOf(PhpSpreadsheet::class, $this->getPrivateProperty(self::SPREADSHEET));
        $this->assertInstanceOf(Worksheet::class, $this->getPrivateProperty(self::WORKSHEET));

        $this->saveFile($spreadsheet, uniqid('testConstruct-', true));
    }

    public function testSave(): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH);
        $spreadsheet->save(self::SAVE_PATH . self::FILE_NAME);

        $this->assertFileExists(self::SAVE_PATH . self::FILE_NAME);
        $this->saveFile($spreadsheet, uniqid('testSave-', true));
    }

    public function testDownload(): void
    {
        $fileName = 'testDownload-' . self::FILE_NAME;
        $spreadsheet = new Spreadsheet(self::FILE_PATH);
        $spreadsheet->save(self::SAVE_PATH . $fileName);

        $this->assertFileExists(self::SAVE_PATH . $fileName);

        ob_start();
        $spreadsheet->download(self::SAVE_PATH, $fileName);
        ob_end_clean();

        $this->assertFileDoesNotExist(self::SAVE_PATH . $fileName);
    }

    /**
     * @dataProvider changeWorksheetProvider
     * */
    public function testGetSheetName(string $fromSheet): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $fromSheet);

        $this->assertSame($fromSheet, $spreadsheet->getSheetName());

        $this->saveFile($spreadsheet, uniqid('testGetSheetName-', true));
    }

    /**
     * @dataProvider changeWorksheetProvider
     * */
    public function testChangeWorksheet(string $fromSheet, string $toSheet, string $value, string $column): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $fromSheet);
        $this->initReflection($spreadsheet);

        $spreadsheet->setCell($column, $value);

        $this->assertSame($value, $spreadsheet->getCell($column));
        $this->assertSame($fromSheet, $spreadsheet->getSheetName());

        $spreadsheet->changeWorksheet($toSheet);

        $this->assertSame($toSheet, $spreadsheet->getSheetName());
        $this->assertInstanceOf(Worksheet::class, $this->getPrivateProperty(self::WORKSHEET));

        $spreadsheet->setCell($column, $value);

        $this->assertSame($value, $spreadsheet->getCell($column));
        $this->saveFile($spreadsheet, uniqid('testChangeWorksheet-', true));
    }

    /**
     * @dataProvider getCellProvider
     * */
    public function testGetCell(string $sheetName, array $cells): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheetName);

        foreach ($cells as $cell) {
            $this->assertNull($spreadsheet->getCell($cell));
        }

        $this->saveFile($spreadsheet, uniqid('testGetCell-', true));
    }

    /**
     * @dataProvider setCellProvider
     * */
    public function testSetCell(string $sheetName, array $cells): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheetName);

        foreach ($cells as $column => $value) {
            $spreadsheet->setCell($column, $value);

            $this->assertSame($value, $spreadsheet->getCell($column));
        }

        $this->saveFile($spreadsheet, uniqid('testSetCell-', true));
    }

    /**
     * @dataProvider addAlignmentHorizontalProvider
     * */
    public function testAddAlignmentHorizontal(string $sheetName, array $cells): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheetName);

        foreach ($cells as $column => $alignment) {
            $spreadsheet->setCell($column, $alignment);

            $this->assertSame($alignment, $spreadsheet->getCell($column));

            $spreadsheet->addAlignmentHorizontal($column, $alignment);

            $this->assertSame($alignment, $spreadsheet->getAlignmentHorizontal($column));
        }

        $this->saveFile($spreadsheet, uniqid('testAddAlignmentHorizontal-', true));
    }

    /**
     * @dataProvider addAlignmentHorizontalProvider
     * */
    public function testGetAlignmentHorizontal(string $sheetName, array $cells): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheetName);

        foreach ($cells as $column => $alignment) {
            $spreadsheet->setCell($column, $alignment);

            $this->assertSame($alignment, $spreadsheet->getCell($column));

            $spreadsheet->addAlignmentHorizontal($column, $alignment);

            $this->assertSame($alignment, $spreadsheet->getAlignmentHorizontal($column));
        }

        $this->saveFile($spreadsheet, uniqid('testGetAlignmentHorizontal-', true));
    }

    /**
     * @dataProvider addBorderProvider
     * */
    public function testAddBorder(array $sheets, array $rows): void
    {
        foreach ($sheets as $sheet => $color) {
            $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheet);

            foreach ($rows as $row) {
                $spreadsheet->setCell($row['column'], $row['value']);

                $this->assertSame($row['value'], $spreadsheet->getCell($row['column']));

                $spreadsheet->addBorder($row['cells'], $row['border'], $color);
            }

            $this->saveFile($spreadsheet, uniqid('testAddBorder-', true));
        }
    }

    /**
     * @dataProvider addBoldProvider
     * */
    public function testAddBold(string $sheetName, string $group, array $cells, string $value): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheetName);
        $this->initReflection($spreadsheet);

        foreach ($cells as $cell) {
            $spreadsheet->setCell($cell, $value);

            $this->assertSame($value, $spreadsheet->getCell($cell));
        }

        $spreadsheet->addBold($group);

        $this->assertTrue($this->getPrivateProperty(self::WORKSHEET)->getStyle($group)->getFont()->getBold());
        $this->saveFile($spreadsheet, uniqid('testAddBold-', true));
    }

    /**
     * @dataProvider addColorProvider
     * */
    public function testAddColor(string $sheetName, string $group, array $cells, string $value, string $color): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheetName);
        $this->initReflection($spreadsheet);

        foreach ($cells as $cell) {
            $spreadsheet->setCell($cell, $value);

            $this->assertSame($value, $spreadsheet->getCell($cell));
        }

        $spreadsheet->addColor($group, $color);

        /** @var Worksheet $worksheet */
        $worksheet = $this->getPrivateProperty(self::WORKSHEET);

        $this->assertSame("FF{$color}", $worksheet->getStyle($group)->getFont()->getColor()->getARGB());
        $this->saveFile($spreadsheet, uniqid('testAddColor-', true));
    }

    /**
     * @dataProvider addBackgroundProvider
     * */
    public function testAddBackground(array $sheets, array $rows): void
    {
        foreach ($sheets as $sheet) {
            $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheet);
            $this->initReflection($spreadsheet);

            foreach ($rows as $row) {
                foreach ($row['cells'] as $cell) {
                    $spreadsheet->setCell($cell, $row['value']);

                    $this->assertSame($row['value'], $spreadsheet->getCell($cell));
                }

                $spreadsheet->addBackground($row['group'], $row['color'], $row['fillType']);

                $colorARGB = $this
                    ->getPrivateProperty(self::WORKSHEET)
                    ->getStyle($row['group'])
                    ->getFill()
                    ->setFillType($row['fillType'])
                    ->getStartColor()
                    ->getARGB($row['color']);

                $this->assertSame("FF{$row['color']}", $colorARGB);
            }

            $this->saveFile($spreadsheet, uniqid('testAddBackground-', true));
        }
    }

    /**
     * @dataProvider addDataValidationProvider
     * */
    public function testAddDataValidation(string $sheetName, string $color, array $data): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS_DATA_VALIDATION, $sheetName);
        $spreadsheet->addDataValidation($data);
        $this->initReflection($spreadsheet);

        /** @var Worksheet $worksheet */
        $worksheet = $this->getPrivateProperty(self::WORKSHEET);

        foreach ($data['columns'] as $column) {
            $validation = $worksheet->getCell($column)->getDataValidation();

            $this->assertSame(Spreadsheet::TYPE_LIST, $validation->getType());
            $this->assertSame(Spreadsheet::STYLE_INFORMATION, $validation->getErrorStyle());
        }

        $spreadsheet->changeWorksheet($data['config']['worksheet']);
        $spreadsheet->addColor("{$data['config']['column']}{$data['config']['start']}", $color);

        $this->saveFile($spreadsheet, uniqid('testAddDataValidation-', true));
    }

    /**
     * @dataProvider addDataValidationWithErrorsProvider
     * */
    public function testAddDataValidationWithErrors(array $data): void
    {
        $this->expectException(Exception::class);

        (new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS_DATA_VALIDATION))->addDataValidation($data);
    }
}
