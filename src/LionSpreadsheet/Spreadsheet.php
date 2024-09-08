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

/**
 * Helps streamline Spreadsheet processes more easily
 *
 * @property PHPSpreadsheet $spreadsheet [Spreadsheet class object]
 * @property Worksheet $worksheet [Worksheet class object]
 * @property string $fileType [File type]
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
     */
    public function __construct(string $path, string $sheetName = '')
    {
        $this->fileType = self::XLSX;

        $this->spreadsheet = IOFactory::createReader($this->fileType)->load($path);

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
        IOFactory::createWriter($this->spreadsheet, $this->fileType)->save($path);
    }

    /**
     * Download the spreadsheet from the defined path
     *
     * @param string $path [File path]
     * @param string $fileName [File name]
     *
     * @return void
     */
    public function download(string $path, string $fileName): void
    {
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        header('Content-Disposition: attachment; filename=' . $fileName);

        header('Content-Length: ' . filesize($path . $fileName));

        readfile($path . $fileName);

        unlink($path . $fileName);
    }

    /**
     * Gets the name of the sheet
     *
     * @return string
     */
    public function getSheetName(): string
    {
        return $this->spreadsheet->getActiveSheet()->getTitle();
    }

    /**
     * Switch spreadsheets
     *
     * @param string $sheetName [Sheet name]
     *
     * @return Spreadsheet
     */
    public function changeWorksheet(string $sheetName): Spreadsheet
    {
        $this->spreadsheet->setActiveSheetIndexByName($sheetName);

        $this->worksheet = $this->spreadsheet->getSheetByName($sheetName);

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
        return $this->worksheet->getCell($columns)->getValue();
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
        $this->worksheet->getStyle($columns)->getAlignment()->setHorizontal($alignment);

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
        return $this->worksheet->getStyle($column)->getAlignment()->getHorizontal();
    }

    /**
     * Add border to one or more cells
     *
     * @param string $columns [Spreadsheet columns]
     * @param string $style [Border style]
     * @param string $color [Border Color]
     *
     * @return Spreadsheet
     */
    public function addBorder(string $columns, string $style = Border::BORDER_THIN, string $color = 'FF0000'): Spreadsheet
    {
        $newColor = new Color($color);

        $this->worksheet->getStyle($columns)->getBorders()->getOutline()->setBorderStyle($style)->setColor($newColor);

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
        $this->worksheet->getStyle($columns)->getFont()->setBold(true);

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
        $this->worksheet->getStyle($columns)->getFont()->getColor()->setARGB($color);

        return $this;
    }

    /**
     * Add a background color to one or more cells
     *
     * @param string $columns [Spreadsheet columns]
     * @param string $color [Background color]
     * @param string $colorStyle [Color style]
     *
     * @return Spreadsheet
     */
    public function addBackground(string $columns, string $color, string $colorStyle = Fill::FILL_SOLID): Spreadsheet
    {
		$this->worksheet->getStyle($columns)->getFill()->setFillType($colorStyle)->getStartColor()->setARGB($color);

        return $this;
	}

    /**
     * Allows you to control what type of information can be entered into a cell
     * or range of cells. With this feature, you can set rules that limit the
     * allowed values, ensuring that the data entered is correct and consistent
     *
     * @param array $data<string, string|array<string, int|string>> [
     * Configuration data list]
     *
     * @throws Exception [If any of the specified parameters are accessible or
     * incorrect]
     */
    public function addDataValidation(array $data): void
    {
        if (empty($data)) {
            throw new Exception('the data configuration is empty', 500);
        }

        if (empty($data['columns'])) {
            throw new Exception('the required columns have not been defined', 500);
        }

        if (empty($data['config'])) {
            throw new Exception('the required configuration has not been defined', 500);
        }

        if (empty($data['config']['error-title'])) {
            throw new Exception('error title not defined', 500);
        }

        if (empty($data['config']['error-message'])) {
            throw new Exception('error message not defined', 500);
        }

        if (empty($data['config']['worksheet'])) {
            throw new Exception('spreadsheet not defined', 500);
        }

        if (empty($data['config']['column'])) {
            throw new Exception('column not defined', 500);
        }

        if (empty($data['config']['start'])) {
            throw new Exception('undefined start', 500);
        }

        if (empty($data['config']['end'])) {
            throw new Exception('undefined end', 500);
        }

        foreach ($data['columns'] as $column) {
            $validation = $this->worksheet->getCell($column)->getDataValidation();

            $validation->setType(DataValidation::TYPE_LIST);

            $validation->setErrorStyle(DataValidation::STYLE_INFORMATION);

            $validation->setAllowBlank(false);

            $validation->setShowInputMessage(true);

            $validation->setShowErrorMessage(true);

            $validation->setShowDropDown(true);

            $validation->setErrorTitle($data['config']['error-title']);

            $validation->setError($data['config']['error-message']);

            $formula = '=' . $data['config']['worksheet'] . '!$' . $data['config']['column'];

            $formula .= '$' . $data['config']['start']. ':$' . $data['config']['column'] . '$' . $data['config']['end'];

            $validation->setFormula1($formula);
        }
    }
}
