<?php

namespace LionSpreadsheet;

use LionSpreadsheet\Traits\Singleton;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PHPSpreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Spreadsheet {

	use Singleton;

	private PHPSpreadsheet $spreadsheet;
    private Worksheet $worksheet;

    private array $excel = [];

    private function loadExcel(string $path, string $name = ""): void {
        $this->spreadsheet = IOFactory::createReader('Xlsx')->load($path);
        $this->worksheet = $name ===  ""
        	? $this->spreadsheet->getActiveSheet()
        	: $this->spreadsheet->getSheetByName($name);
    }

    private function saveExcel(string $path): void {
        IOFactory::createWriter($this->spreadsheet, "Xlsx")->save($path);
    }

    private function changeWorksheet(string $name): void {
        $this->worksheet = $this->spreadsheet->getSheetByName($name);
    }

    private function getCell(string $column): ?string {
        return $this->worksheet->getCell($column)->getValue();
    }

    private function setCell(string $column, mixed $value): void {
        $this->worksheet->setCellValue($column, $value);
    }

    private function addAlignmentHorizontal(string $columns, string $alignment) {
        $this->worksheet->getStyle($columns)->getAlignment()->setHorizontal($alignment);
    }

    private function addBorder(string $columns, string $style, string $color): void {
        $this->worksheet
            ->getStyle($columns)
            ->getBorders()
            ->getOutline()
            ->setBorderStyle($style)
            ->setColor(new Color($color));
    }

    private function addBold(string $columns): void {
        $this->worksheet->getStyle($columns)->getFont()->setBold(true);
    }

    private function addColor(string $columns, string $color): void {
        $this->worksheet
            ->getStyle($columns)
            ->getFont()
            ->getColor()
            ->setARGB($color);
    }

    private function addDataValidation(array $columns, array $config): void {
        foreach ($columns as $key => $column) {
            $validation = $this->worksheet->getCell($column)->getDataValidation();
            $validation->setType(DataValidation::TYPE_LIST);
            $validation->setErrorStyle(DataValidation::STYLE_INFORMATION);
            $validation->setAllowBlank(false);
            $validation->setShowInputMessage(true);
            $validation->setShowErrorMessage(true);
            $validation->setShowDropDown(true);
            $validation->setErrorTitle($config['error_title']);
            $validation->setError($config['error_message']);

            $validation->setFormula1(
            	isset($config['worksheet'])
            		? '=' . $config['worksheet'] . '!$' . $config['column'] . '$' . $config['start'] . ':$' . $config['column'] . '$' . $config['end']
            		: '=$' . $config['column'] . '$' . $config['start'] . ':$' . $config['column'] . '$' . $config['end']
            );
        }
    }

}