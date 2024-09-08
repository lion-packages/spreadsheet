<?php

declare(strict_types=1);

namespace Tests;

use Exception;
use Lion\Spreadsheet\Spreadsheet;
use Lion\Test\Test;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PhpSpreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\Attributes\Test as Testing;
use Tests\Provider\SpreadsheetProviderTrait;

class SpreadsheetTest extends Test
{
    use SpreadsheetProviderTrait;

    private const string SAVE_PATH = './storage/';
    private const string SUPPORT_PATH = './tests/support-files/';
    private const string FILE_NAME = 'template.xlsx';
    private const string FILE_NAME_MULTIPLE_SHEETS = 'template-multiple-sheets.xlsx';
    private const string FILE_NAME_MULTIPLE_SHEETS_DATA_VALIDATION = 'template-multiple-sheets-data-validation.xlsx';
    private const string FILE_PATH = self::SUPPORT_PATH . self::FILE_NAME;
    private const string FILE_PATH_MULTIPLE_SHEETS = self::SUPPORT_PATH . self::FILE_NAME_MULTIPLE_SHEETS;
    private const string FILE_PATH_MULTIPLE_SHEETS_DATA_VALIDATION = self::SUPPORT_PATH . self::FILE_NAME_MULTIPLE_SHEETS_DATA_VALIDATION;
    private const string FILE_TYPE = 'fileType';
    private const string SPREADSHEET = 'spreadsheet';
    private const string WORKSHEET = 'worksheet';
    private const string CONTENT_TYPE = 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    private const string CONTENT_DISPOSITION = 'Content-Disposition: attachment; filename=' . self::FILE_NAME;

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

    #[Testing]
    public function construct(): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH);

        $this->initReflection($spreadsheet);

        $this->assertSame(Spreadsheet::XLSX, $this->getPrivateProperty(self::FILE_TYPE));
        $this->assertInstanceOf(PhpSpreadsheet::class, $this->getPrivateProperty(self::SPREADSHEET));
        $this->assertInstanceOf(Worksheet::class, $this->getPrivateProperty(self::WORKSHEET));

        $this->saveFile($spreadsheet, uniqid('testConstruct-', true));
    }

    #[Testing]
    public function save(): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH);

        $spreadsheet->save(self::SAVE_PATH . self::FILE_NAME);

        $this->assertFileExists(self::SAVE_PATH . self::FILE_NAME);
        $this->saveFile($spreadsheet, uniqid('testSave-', true));
    }

    #[Testing]
    public function download(): void
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

    #[Testing]
    #[DataProvider('changeWorksheetProvider')]
    public function getSheetName(string $fromSheet, string $toSheet, string $value, string $column): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $fromSheet);

        $this->assertSame($fromSheet, $spreadsheet->getSheetName());

        $this->saveFile($spreadsheet, uniqid('testGetSheetName-', true));
    }

    #[Testing]
    #[DataProvider('changeWorksheetProvider')]
    public function changeWorksheet(string $fromSheet, string $toSheet, string $value, string $column): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $fromSheet);

        $this->initReflection($spreadsheet);

        $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->setCell($column, $value));
        $this->assertSame($value, $spreadsheet->getCell($column));
        $this->assertSame($fromSheet, $spreadsheet->getSheetName());
        $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->changeWorksheet($toSheet));
        $this->assertSame($toSheet, $spreadsheet->getSheetName());
        $this->assertInstanceOf(Worksheet::class, $this->getPrivateProperty(self::WORKSHEET));
        $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->setCell($column, $value));
        $this->assertSame($value, $spreadsheet->getCell($column));
        $this->saveFile($spreadsheet, uniqid('testChangeWorksheet-', true));
    }

    #[Testing]
    #[DataProvider('getCellProvider')]
    public function getCell(string $sheetName, array $cells): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheetName);

        foreach ($cells as $cell) {
            $this->assertNull($spreadsheet->getCell($cell));
        }

        $this->saveFile($spreadsheet, uniqid('testGetCell-', true));
    }

    #[Testing]
    #[DataProvider('setCellProvider')]
    public function setCell(string $sheetName, array $cells): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheetName);

        foreach ($cells as $column => $value) {
            $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->setCell($column, $value));

            $this->assertSame($value, $spreadsheet->getCell($column));
        }

        $this->saveFile($spreadsheet, uniqid('testSetCell-', true));
    }

    #[Testing]
    #[DataProvider('addAlignmentHorizontalProvider')]
    public function addAlignmentHorizontal(string $sheetName, array $cells): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheetName);

        foreach ($cells as $column => $alignment) {
            $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->setCell($column, $alignment));
            $this->assertSame($alignment, $spreadsheet->getCell($column));
            $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->addAlignmentHorizontal($column, $alignment));
            $this->assertSame($alignment, $spreadsheet->getAlignmentHorizontal($column));
        }

        $this->saveFile($spreadsheet, uniqid('testAddAlignmentHorizontal-', true));
    }

    #[Testing]
    #[DataProvider('addAlignmentHorizontalProvider')]
    public function getAlignmentHorizontal(string $sheetName, array $cells): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheetName);

        foreach ($cells as $column => $alignment) {
            $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->setCell($column, $alignment));
            $this->assertSame($alignment, $spreadsheet->getCell($column));
            $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->addAlignmentHorizontal($column, $alignment));
            $this->assertSame($alignment, $spreadsheet->getAlignmentHorizontal($column));
        }

        $this->saveFile($spreadsheet, uniqid('testGetAlignmentHorizontal-', true));
    }

    #[Testing]
    #[DataProvider('addBorderProvider')]
    public function addBorder(array $sheets, array $rows): void
    {
        foreach ($sheets as $sheet => $color) {
            $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheet);

            foreach ($rows as $row) {
                $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->setCell($row['column'], $row['value']));
                $this->assertSame($row['value'], $spreadsheet->getCell($row['column']));

                $this->assertInstanceOf(
                    Spreadsheet::class,
                    $spreadsheet->addBorder($row['cells'], $row['border'], $color)
                );
            }

            $this->saveFile($spreadsheet, uniqid('testAddBorder-', true));
        }
    }

    #[Testing]
    #[DataProvider('addBoldProvider')]
    public function addBold(string $sheetName, string $group, array $cells, string $value): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheetName);

        $this->initReflection($spreadsheet);

        foreach ($cells as $cell) {
            $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->setCell($cell, $value));
            $this->assertSame($value, $spreadsheet->getCell($cell));
        }

        $spreadsheet->addBold($group);

        $this->assertTrue($this->getPrivateProperty(self::WORKSHEET)->getStyle($group)->getFont()->getBold());
        $this->saveFile($spreadsheet, uniqid('testAddBold-', true));
    }

    #[Testing]
    #[DataProvider('addColorProvider')]
    public function addColor(string $sheetName, string $group, array $cells, string $value, string $color): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheetName);

        $this->initReflection($spreadsheet);

        foreach ($cells as $cell) {
            $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->setCell($cell, $value));
            $this->assertSame($value, $spreadsheet->getCell($cell));
        }

        $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->addColor($group, $color));

        /** @var Worksheet $worksheet */
        $worksheet = $this->getPrivateProperty(self::WORKSHEET);

        $this->assertSame("FF{$color}", $worksheet->getStyle($group)->getFont()->getColor()->getARGB());
        $this->saveFile($spreadsheet, uniqid('testAddColor-', true));
    }

    #[Testing]
    #[DataProvider('addBackgroundProvider')]
    public function addBackground(array $sheets, array $rows): void
    {
        foreach ($sheets as $sheet) {
            $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS, $sheet);

            $this->initReflection($spreadsheet);

            foreach ($rows as $row) {
                foreach ($row['cells'] as $cell) {
                    $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->setCell($cell, $row['value']));
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
     * @throws Exception
     */
    #[Testing]
    #[DataProvider('addDataValidationProvider')]
    public function addDataValidation(string $sheetName, string $color, array $data): void
    {
        $spreadsheet = new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS_DATA_VALIDATION, $sheetName);

        $spreadsheet->addDataValidation($data);

        $this->initReflection($spreadsheet);

        /** @var Worksheet $worksheet */
        $worksheet = $this->getPrivateProperty(self::WORKSHEET);

        foreach ($data['columns'] as $column) {
            $validation = $worksheet->getCell($column)->getDataValidation();

            $this->assertSame(DataValidation::TYPE_LIST, $validation->getType());
            $this->assertSame(DataValidation::STYLE_INFORMATION, $validation->getErrorStyle());
        }

        $this->assertInstanceOf(Spreadsheet::class, $spreadsheet->changeWorksheet($data['config']['worksheet']));

        $this->assertInstanceOf(
            Spreadsheet::class,
            $spreadsheet->addColor("{$data['config']['column']}{$data['config']['start']}", $color)
        );

        $this->saveFile($spreadsheet, uniqid('testAddDataValidation-', true));
    }

    #[Testing]
    #[DataProvider('addDataValidationWithErrorsProvider')]
    public function addDataValidationWithErrors(array $data, string $exceptionMessage): void
    {
        $this->expectException(Exception::class);
        $this->expectExceptionMessage($exceptionMessage);
        $this->expectExceptionCode(500);

        (new Spreadsheet(self::FILE_PATH_MULTIPLE_SHEETS_DATA_VALIDATION))->addDataValidation($data);
    }
}
