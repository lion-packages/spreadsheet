<?php

declare(strict_types=1);

namespace LionSpreadsheet;

use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PHPSpreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Spreadsheet
{
    const XLSX = 'Xlsx';
    const XLS = 'Xls';

	private PHPSpreadsheet $spreadsheet;
    private Worksheet $worksheet;

    private string $fileType;

    public function load(string $path, string $type = self::XLSX, string $name = ''): void
    {
        $this->fileType = $type;
        $this->spreadsheet = IOFactory::createReader($this->fileType)->load($path);

        $this->worksheet = empty($name)
        	? $this->spreadsheet->getActiveSheet()
        	: $this->spreadsheet->getSheetByName($name);
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

    public function changeWorksheet(string $name): void
    {
        $this->worksheet = $this->spreadsheet->getSheetByName($name);
    }

    public function getCell(string $column): ?string
    {
        return $this->worksheet->getCell($column)->getValue();
    }

    public function setCell(string $column, mixed $value): void
    {
        $this->worksheet->setCellValue($column, $value);
    }

    public function addAlignmentHorizontal(string $columns, string $alignment)
    {
        $this->worksheet->getStyle($columns)->getAlignment()->setHorizontal($alignment);
    }

    public function addBorder(string $columns, string $style, string $color): void
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

    public function addBackground(string $columns, string $color, ?string $type_color = Fill::FILL_SOLID): void
    {
		$this->worksheet->getStyle($columns)->getFill()->setFillType($type_color)->getStartColor()->setARGB($color);
	}

    public function addDataValidation(array $data): void
    {
        foreach ($data['columns'] as $column) {
            $validation = $this->worksheet->getCell($column)->getDataValidation();
            $validation->setType(DataValidation::TYPE_LIST);
            $validation->setErrorStyle(DataValidation::STYLE_INFORMATION);
            $validation->setAllowBlank(false);
            $validation->setShowInputMessage(true);
            $validation->setShowErrorMessage(true);
            $validation->setShowDropDown(true);
            $validation->setErrorTitle($data['config']['error_title']);
            $validation->setError($data['config']['error_message']);

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
