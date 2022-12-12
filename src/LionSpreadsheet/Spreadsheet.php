<?php

namespace LionSpreadsheet;

use LionSpreadsheet\Traits\Singleton;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PHPSpreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Spreadsheet {

	use Singleton;

	private static PHPSpreadsheet $spreadsheet;
    private static Worksheet $worksheet;

    private static array $excel = [];

    public static function loadExcel(string $path, string $name = ""): void {
        self::$spreadsheet = IOFactory::createReader('Xlsx')->load($path);
        self::$worksheet = $name ===  ""
        	? self::$spreadsheet->getActiveSheet()
        	: self::$spreadsheet->getSheetByName($name);
    }

    public static function saveExcel(string $path): void {
        IOFactory::createWriter(self::$spreadsheet, "Xlsx")->save($path);
    }

    public static function changeWorksheet(string $name): void {
        self::$worksheet = self::$spreadsheet->getSheetByName($name);
    }

    public static function getCell(string $column): ?string {
        return self::$worksheet->getCell($column)->getValue();
    }

    public static function setCell(string $column, mixed $value): void {
        self::$worksheet->setCellValue($column, $value);
    }

    public static function addAlignmentHorizontal(string $columns, string $alignment) {
        self::$worksheet->getStyle($columns)->getAlignment()->setHorizontal($alignment);
    }

    public static function addBorder(string $columns, string $style, string $color): void {
        self::$worksheet
            ->getStyle($columns)
            ->getBorders()
            ->getOutline()
            ->setBorderStyle($style)
            ->setColor(new Color($color));
    }

    public static function addBold(string $columns): void {
        self::$worksheet->getStyle($columns)->getFont()->setBold(true);
    }

    public static function addColor(string $columns, string $color): void {
        self::$worksheet
            ->getStyle($columns)
            ->getFont()
            ->getColor()
            ->setARGB($color);
    }

    public static function addBackground(string $columns, string $color, ?string $type_color = Fill::FILL_SOLID): void {
		self::$worksheet
            ->getStyle($columns)
            ->getFill()
            ->setFillType($type_color)
            ->getStartColor()
            ->setARGB($color);
	}

    public static function addDataValidation(array $columns, array $config): void {
        foreach ($columns as $key => $column) {
            $validation = self::$worksheet->getCell($column)->getDataValidation();
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