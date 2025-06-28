<?php

declare(strict_types=1);

namespace Lion\Spreadsheet;

use Exception;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PHPSpreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use RuntimeException;

/**
 * Helps streamline Spreadsheet processes more easily
 *
 * @package Lion\Spreadsheet
 */
class Spreadsheet
{
    /**
     * [Constant to define a spreadsheet with .xlsx extension]
     *
     * @const XLSX
     */
    public const string XLSX = 'Xlsx';

    /**
     * [Spreadsheet class object]
     *
     * @var PHPSpreadsheet $spreadsheet
     */
    private PHPSpreadsheet $spreadsheet;

    /**
     * [Worksheet class object]
     *
     * @var Worksheet $worksheet
     */
    private Worksheet $worksheet;

    /**
     * [File type]
     *
     * @var string $fileType
     */
    private string $fileType;

    /**
     * Class constructor
     *
     * @param string $path [File path]
     * @param string $sheetName [Sheet name]
     *
     * @throws Exception [If the worksheet does not exist]
     */
    public function __construct(string $path, string $sheetName = '')
    {
        $this->fileType = self::XLSX;

        $this->spreadsheet = IOFactory::createReader($this->fileType)
            ->load($path);

        $this->worksheet = $this->spreadsheet->getActiveSheet();

        if (!empty($sheetName)) {
            $this->changeWorksheet($sheetName);
        }
    }

    /**
     * Store spreadsheets in a path
     *
     * @param string $path [File path]
     *
     * @return void
     */
    public function save(string $path): void
    {
        IOFactory::createWriter($this->spreadsheet, $this->fileType)
            ->save($path);
    }

    /**
     * Download the spreadsheet from the defined path
     *
     * @param string $path [File path]
     * @param string $fileName [File name]
     *
     * @return void
     *
     * @infection-ignore-all
     */
    public function download(string $path, string $fileName): void
    {
        $filePath = rtrim($path, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . $fileName;

        if (!is_readable($filePath)) {
            throw new RuntimeException("The file does not exist or cannot be read: " . $filePath);
        }

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        header('Content-Disposition: attachment; filename="' . basename($fileName) . '"');

        header('Content-Length: ' . filesize($filePath));

        readfile($filePath);

        if (!unlink($filePath)) {
            throw new RuntimeException("The file could not be deleted: " . $filePath);
        }
    }

    /**
     * Gets the name of the sheet
     *
     * @return string
     */
    public function getSheetName(): string
    {
        return $this->spreadsheet
            ->getActiveSheet()
            ->getTitle();
    }

    /**
     * Switch spreadsheets
     *
     * @param string $sheetName [Sheet name]
     *
     * @return Spreadsheet
     *
     * @throws Exception [If the worksheet does not exist]
     */
    public function changeWorksheet(string $sheetName): Spreadsheet
    {
        $this->spreadsheet->setActiveSheetIndexByName($sheetName);

        $worksheet = $this->spreadsheet->getSheetByName($sheetName);

        if (!$worksheet) {
            throw new Exception("Worksheet not found");
        }

        $this->worksheet = $worksheet;

        return $this;
    }

    /**
     * Gets the value of one or more cells
     *
     * @param string $columns [Spreadsheet columns]
     *
     * @return mixed
     */
    public function getCell(string $columns): mixed
    {
        return $this->worksheet
            ->getCell($columns)
            ->getValue();
    }

    /**
     * Modify the value of one or more cells
     *
     * @param string $columns [Spreadsheet columns]
     * @param mixed $value [Cell value]
     *
     * @return Spreadsheet
     */
    public function setCell(string $columns, mixed $value): Spreadsheet
    {
        $this->worksheet->setCellValue($columns, $value);

        return $this;
    }

    /**
     * Horizontally aligns the value of one or more cells
     *
     * @param string $columns [Spreadsheet columns]
     * @param string $alignment [Horizontal alignment]
     *
     * @return Spreadsheet
     */
    public function addAlignmentHorizontal(string $columns, string $alignment): Spreadsheet
    {
        $this->worksheet
            ->getStyle($columns)
            ->getAlignment()
            ->setHorizontal($alignment);

        return $this;
    }

    /**
     * Gets the horizontal alignment
     *
     * @param string $column [Spreadsheet column]
     *
     * @return string|null
     */
    public function getAlignmentHorizontal(string $column): ?string
    {
        return $this->worksheet
            ->getStyle($column)
            ->getAlignment()
            ->getHorizontal();
    }

    /**
     * Add border to one or more cells
     *
     * @param string $cells [Spreadsheet cells]
     * @param string $style [Border style]
     * @param string $color [Border Color]
     *
     * @return Spreadsheet
     *
     * @infection-ignore-all
     */
    public function addBorder(
        string $cells,
        string $style = Border::BORDER_THIN,
        string $color = 'FF0000'
    ): Spreadsheet {
        $newColor = new Color($color);

        $this->worksheet
            ->getStyle($cells)
            ->getBorders()
            ->getOutline()
            ->setBorderStyle($style)
            ->setColor($newColor);

        return $this;
    }

    /**
     * Add bold to one or more cells
     *
     * @param string $columns [Spreadsheet columns]
     *
     * @return Spreadsheet
     */
    public function addBold(string $columns): Spreadsheet
    {
        $this->worksheet
            ->getStyle($columns)
            ->getFont()
            ->setBold(true);

        return $this;
    }

    /**
     * Add letter color
     *
     * @param string $columns [Spreadsheet columns]
     * @param string $color [Letter color]
     *
     * @return Spreadsheet
     */
    public function addColor(string $columns, string $color): Spreadsheet
    {
        $this->worksheet
            ->getStyle($columns)
            ->getFont()
            ->getColor()
            ->setARGB($color);

        return $this;
    }

    /**
     * Add a background color to one or more cells
     *
     * @param string $cells Spreadsheet columns
     * @param string $color Background color
     * @param string $colorStyle Color style
     *
     * @return Spreadsheet
     *
     * @link https://phpspreadsheet.readthedocs.io/en/latest/topics/recipes/#formatting-cells
     */
    public function addBackground(string $cells, string $color, string $colorStyle = Fill::FILL_SOLID): Spreadsheet
    {
        $this->worksheet
            ->getStyle($cells)
            ->getFill()
            ->setFillType($colorStyle)
            ->getStartColor()
            ->setARGB($color);

        return $this;
    }

    /**
     * Allows you to control what type of information can be entered into a cell
     * or range of cells. With this feature, you can set rules that limit the
     * allowed values, ensuring that the data entered is correct and consistent
     *
     * @param array{
     *     columns: array<int, string>,
     *     config: array{
     *         error-title: string,
     *         error-message: string,
     *         worksheet: string,
     *         column: string,
     *         start: string,
     *         end: string
     *     }
     * } $data [Configuration data list]
     *
     * @return void
     *
     * @throws Exception [If any of the specified parameters are accessible or
     * incorrect]
     */
    public function addDataValidation(array $data): void
    {
        foreach ($data['columns'] as $column) {
            $validation = $this->worksheet
                ->getCell($column)
                ->getDataValidation();

            $validation->setType(DataValidation::TYPE_LIST);

            $validation->setErrorStyle(DataValidation::STYLE_INFORMATION);

            $validation->setAllowBlank(false);

            $validation->setShowInputMessage(true);

            $validation->setShowErrorMessage(true);

            $validation->setShowDropDown(true);

            $validation->setErrorTitle($data['config']['error-title']);

            $validation->setError($data['config']['error-message']);

            $formula = '=' . $data['config']['worksheet'] . '!$' . $data['config']['column'];

            $formula .= '$' . $data['config']['start'] . ':';

            $formula .= '$' . $data['config']['column'] . '$' . $data['config']['end'];

            $validation->setFormula1($formula);
        }
    }
}
