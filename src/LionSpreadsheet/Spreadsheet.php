<?php

declare(strict_types=1);

namespace Lion\Spreadsheet;

use Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PHPSpreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Spreadsheet
{
    const XLSX = 'Xlsx';

    const BORDER_NONE = 'none';
    const BORDER_DASHDOT = 'dashDot';
    const BORDER_DASHDOTDOT = 'dashDotDot';
    const BORDER_DASHED = 'dashed';
    const BORDER_DOTTED = 'dotted';
    const BORDER_DOUBLE = 'double';
    const BORDER_HAIR = 'hair';
    const BORDER_MEDIUM = 'medium';
    const BORDER_MEDIUMDASHDOT = 'mediumDashDot';
    const BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot';
    const BORDER_MEDIUMDASHED = 'mediumDashed';
    const BORDER_SLANTDASHDOT = 'slantDashDot';
    const BORDER_THICK = 'thick';
    const BORDER_THIN = 'thin';
    const BORDER_OMIT = 'omit';

    const FILL_NONE = 'none';
    const FILL_SOLID = 'solid';
    const FILL_GRADIENT_LINEAR = 'linear';
    const FILL_GRADIENT_PATH = 'path';
    const FILL_PATTERN_DARKDOWN = 'darkDown';
    const FILL_PATTERN_DARKGRAY = 'darkGray';
    const FILL_PATTERN_DARKGRID = 'darkGrid';
    const FILL_PATTERN_DARKHORIZONTAL = 'darkHorizontal';
    const FILL_PATTERN_DARKTRELLIS = 'darkTrellis';
    const FILL_PATTERN_DARKUP = 'darkUp';
    const FILL_PATTERN_DARKVERTICAL = 'darkVertical';
    const FILL_PATTERN_GRAY0625 = 'gray0625';
    const FILL_PATTERN_GRAY125 = 'gray125';
    const FILL_PATTERN_LIGHTDOWN = 'lightDown';
    const FILL_PATTERN_LIGHTGRAY = 'lightGray';
    const FILL_PATTERN_LIGHTGRID = 'lightGrid';
    const FILL_PATTERN_LIGHTHORIZONTAL = 'lightHorizontal';
    const FILL_PATTERN_LIGHTTRELLIS = 'lightTrellis';
    const FILL_PATTERN_LIGHTUP = 'lightUp';
    const FILL_PATTERN_LIGHTVERTICAL = 'lightVertical';
    const FILL_PATTERN_MEDIUMGRAY = 'mediumGray';

    const TYPE_NONE = 'none';
    const TYPE_CUSTOM = 'custom';
    const TYPE_DATE = 'date';
    const TYPE_DECIMAL = 'decimal';
    const TYPE_LIST = 'list';
    const TYPE_TEXTLENGTH = 'textLength';
    const TYPE_TIME = 'time';
    const TYPE_WHOLE = 'whole';

    const STYLE_STOP = 'stop';
    const STYLE_WARNING = 'warning';
    const STYLE_INFORMATION = 'information';

	private PHPSpreadsheet $spreadsheet;
    private Worksheet $worksheet;

    private string $fileType;

    public function __construct(string $path, string $sheetName = '')
    {
        $this->fileType = self::XLSX;
        $this->spreadsheet = IOFactory::createReader($this->fileType)->load($path);
        $this->worksheet = $this->spreadsheet->getActiveSheet();

        if (!empty($sheetName)) {
            $this->changeWorksheet($sheetName);
        }
    }

    public function save(string $path): void
    {
        IOFactory::createWriter($this->spreadsheet, $this->fileType)->save($path);
    }

    public function download(string $path, string $file_name): void
    {
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename=' . $file_name);
        header('Content-Length: ' . filesize($path . $file_name));
        readfile($path . $file_name);
        unlink($path . $file_name);
    }

    public function getSheetName(): mixed
    {
        return $this->spreadsheet->getActiveSheet()->getTitle();
    }

    public function changeWorksheet(string $sheetName): void
    {
        $this->spreadsheet->setActiveSheetIndexByName($sheetName);
        $this->worksheet = $this->spreadsheet->getSheetByName($sheetName);
    }

    public function getCell(string $columns): ?string
    {
        return $this->worksheet->getCell($columns)->getValue();
    }

    public function setCell(string $columns, mixed $value): void
    {
        $this->worksheet->setCellValue($columns, $value);
    }

    public function addAlignmentHorizontal(string $columns, string $alignment): void
    {
        $this->worksheet->getStyle($columns)->getAlignment()->setHorizontal($alignment);
    }

    public function getAlignmentHorizontal(string $column): string
    {
        return $this->worksheet->getStyle($column)->getAlignment()->getHorizontal();
    }

    public function addBorder(string $columns, string $style = self::BORDER_THIN, string $color = 'FF0000'): void
    {
        $newColor = new Color($color);

        $this->worksheet->getStyle($columns)->getBorders()->getOutline()->setBorderStyle($style)->setColor($newColor);
    }

    public function addBold(string $columns): void
    {
        $this->worksheet->getStyle($columns)->getFont()->setBold(true);
    }

    public function addColor(string $columns, string $color): void
    {
        $this->worksheet->getStyle($columns)->getFont()->getColor()->setARGB($color);
    }

    public function addBackground(string $columns, string $color, ?string $type_color = self::FILL_SOLID): void
    {
		$this->worksheet->getStyle($columns)->getFill()->setFillType($type_color)->getStartColor()->setARGB($color);
	}

    public function addDataValidation(array $data): void
    {
        if (empty($data['columns'])) {
            throw new Exception('the required columns have not been defined');
        }

        if (empty($data['config'])) {
            throw new Exception('the required configuration has not been defined');
        }

        if (empty($data['config']['error-title'])) {
            throw new Exception('error title not defined');
        }

        if (empty($data['config']['error-message'])) {
            throw new Exception('error message not defined');
        }

        if (empty($data['config']['worksheet'])) {
            throw new Exception('spreadsheet not defined');
        }

        if (empty($data['config']['column'])) {
            throw new Exception('column not defined');
        }

        if (empty($data['config']['start'])) {
            throw new Exception('undefined start');
        }

        if (empty($data['config']['end'])) {
            throw new Exception('undefined end');
        }

        foreach ($data['columns'] as $column) {
            $validation = $this->worksheet->getCell($column)->getDataValidation();
            $validation->setType(self::TYPE_LIST);
            $validation->setErrorStyle(self::STYLE_INFORMATION);
            $validation->setAllowBlank(false);
            $validation->setShowInputMessage(true);
            $validation->setShowErrorMessage(true);
            $validation->setShowDropDown(true);
            $validation->setErrorTitle($data['config']['error-title']);
            $validation->setError($data['config']['error-message']);

            if (isset($data['config']['worksheet'])) {
                $validation->setFormula1(
                    '=' . $data['config']['worksheet'] . '!$' . $data['config']['column'] . '$' . $data['config']['start'] . ':$' . $data['config']['column'] . '$' . $data['config']['end']
                );
            } else {
                $validation->setFormula1(
                    '=$' . $data['config']['column'] . '$' . $data['config']['start'] . ':$' . $data['config']['column'] . '$' . $data['config']['end']
                );
            }
        }
    }
}
