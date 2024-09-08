<?php

declare(strict_types=1);

namespace Tests\Provider;

use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

trait SpreadsheetProviderTrait
{
    public static function changeWorksheetProvider(): array
    {
        return [
            [
                'fromSheet' => 'Hoja1',
                'toSheet' => 'Hoja2',
                'value' => uniqid('VALUE-', true),
                'column' => 'A2'
            ],
            [
                'fromSheet' => 'Hoja1',
                'toSheet' => 'Hoja3',
                'value' => uniqid('VALUE-', true),
                'column' => 'B3'
            ],
            [
                'fromSheet' => 'Hoja1',
                'toSheet' => 'Hoja4',
                'value' => uniqid('VALUE-', true),
                'column' => 'C4'
            ],
            [
                'fromSheet' => 'Hoja2',
                'toSheet' => 'Hoja1',
                'value' => uniqid('VALUE-', true),
                'column' => 'D5'
            ]
        ];
    }

    public static function getCellProvider(): array
    {
        return [
            [
                'sheetName' => 'Hoja1',
                'cells' => ['A1', 'B1', 'C1', 'D1', 'E1']
            ],
            [
                'sheetName' => 'Hoja2',
                'cells' => ['A2', 'B2', 'C2', 'D2', 'E2']
            ],
            [
                'sheetName' => 'Hoja3',
                'cells' => ['A3', 'B3', 'C3', 'D3', 'E3']
            ],
            [
                'sheetName' => 'Hoja4',
                'cells' => ['A4', 'B4', 'C4', 'D4', 'E4']
            ]
        ];
    }

    public static function setCellProvider(): array
    {
        return [
            [
                'sheetName' => 'Hoja1',
                'cells' => [
                    'A1' => 'VALUE-A1',
                    'B1' => 'VALUE-B1',
                    'C1' => 'VALUE-C1',
                    'D1' => 'VALUE-D1',
                    'E1' => 'VALUE-E1'
                ]
            ],
            [
                'sheetName' => 'Hoja2',
                'cells' => [
                    'A2' => 'VALUE-A2',
                    'B2' => 'VALUE-B2',
                    'C2' => 'VALUE-C2',
                    'D2' => 'VALUE-D2',
                    'E2' => 'VALUE-E2'
                ]
            ],
            [
                'sheetName' => 'Hoja3',
                'cells' => [
                    'A3' => 'VALUE-A3',
                    'B3' => 'VALUE-B3',
                    'C3' => 'VALUE-C3',
                    'D3' => 'VALUE-D3',
                    'E3' => 'VALUE-E3'
                ]
            ],
            [
                'sheetName' => 'Hoja4',
                'cells' => [
                    'A4' => 'VALUE-A4',
                    'B4' => 'VALUE-B4',
                    'C4' => 'VALUE-C4',
                    'D4' => 'VALUE-D4',
                    'E4' => 'VALUE-E4'
                ]
            ]
        ];
    }

    public static function addAlignmentHorizontalProvider(): array
    {
        return [
            [
                'sheetName' => 'Hoja1',
                'cells' => [
                    'A1' => 'center',
                    'B1' => 'right',
                    'C1' => 'left'
                ]
            ],
            [
                'sheetName' => 'Hoja2',
                'cells' => [
                    'A2' => 'center',
                    'B2' => 'right',
                    'C2' => 'left'
                ]
            ],
            [
                'sheetName' => 'Hoja3',
                'cells' => [
                    'A3' => 'center',
                    'B3' => 'right',
                    'C3' => 'left'
                ]
            ],
            [
                'sheetName' => 'Hoja4',
                'cells' => [
                    'A4' => 'center',
                    'B4' => 'right',
                    'C4' => 'left'
                ]
            ]
        ];
    }

    public static function addBorderProvider(): array
    {
        return [
            [
                'sheets' => [
                    'Hoja1' => '17FF00',
                    'Hoja2' => 'FF4200',
                    'Hoja3' => '007CFF',
                    'Hoja4' => '9B00FF'
                ],
                'rows' => [
                    [
                        'cells' => 'A1:E1',
                        'border' => Border::BORDER_NONE,
                        'value' => 'VALUE',
                        'column' => 'A1'
                    ],
                    [
                        'cells' => 'A3:E3',
                        'border' => Border::BORDER_DASHDOT,
                        'value' => 'VALUE',
                        'column' => 'A3'
                    ],
                    [
                        'cells' => 'A5:E5',
                        'border' => Border::BORDER_DASHDOTDOT,
                        'value' => 'VALUE',
                        'column' => 'A5'
                    ],
                    [
                        'cells' => 'A7:E7',
                        'border' => Border::BORDER_DASHED,
                        'value' => 'VALUE',
                        'column' => 'A7'
                    ],
                    [
                        'cells' => 'A9:E9',
                        'border' => Border::BORDER_DOTTED,
                        'value' => 'VALUE',
                        'column' => 'A9'
                    ],
                    [
                        'cells' => 'A11:E11',
                        'border' => Border::BORDER_DOUBLE,
                        'value' => 'VALUE',
                        'column' => 'A11'
                    ],
                    [
                        'cells' => 'A13:E13',
                        'border' => Border::BORDER_HAIR,
                        'value' => 'VALUE',
                        'column' => 'A13'
                    ],
                    [
                        'cells' => 'A15:E15',
                        'border' => Border::BORDER_MEDIUM,
                        'value' => 'VALUE',
                        'column' => 'A15'
                    ],
                    [
                        'cells' => 'A17:E17',
                        'border' => Border::BORDER_MEDIUMDASHDOT,
                        'value' => 'VALUE',
                        'column' => 'A17'
                    ],
                    [
                        'cells' => 'A19:E19',
                        'border' => Border::BORDER_MEDIUMDASHDOTDOT,
                        'value' => 'VALUE',
                        'column' => 'A19'
                    ],
                    [
                        'cells' => 'A21:E21',
                        'border' => Border::BORDER_MEDIUMDASHED,
                        'value' => 'VALUE',
                        'column' => 'A21'
                    ],
                    [
                        'cells' => 'A23:E23',
                        'border' => Border::BORDER_SLANTDASHDOT,
                        'value' => 'VALUE',
                        'column' => 'A23'
                    ],
                    [
                        'cells' => 'A25:E25',
                        'border' => Border::BORDER_THICK,
                        'value' => 'VALUE',
                        'column' => 'A25'
                    ],
                    [
                        'cells' => 'A27:E27',
                        'border' => Border::BORDER_THIN,
                        'value' => 'VALUE',
                        'column' => 'A27'
                    ],
                    [
                        'cells' => 'A29:E29',
                        'border' => Border::BORDER_DASHDOT,
                        'value' => 'VALUE',
                        'column' => 'A29'
                    ]
                ]
            ]
        ];
    }

    public static function addBoldProvider(): array
    {
        return [
            [
                'sheetName' => 'Hoja1',
                'group' => 'A1:E1',
                'cells' => ['A1', 'B1', 'C1', 'D1', 'E1'],
                'value' => 'VALUE'
            ],
            [
                'sheetName' => 'Hoja2',
                'group' => 'A2:E2',
                'cells' => ['A2', 'B2', 'C2', 'D2', 'E2'],
                'value' => 'VALUE'
            ],
            [
                'sheetName' => 'Hoja3',
                'group' => 'A3:E3',
                'cells' => ['A3', 'B3', 'C3', 'D3', 'E3'],
                'value' => 'VALUE'
            ],
            [
                'sheetName' => 'Hoja4',
                'group' => 'A4:E4',
                'cells' => ['A4', 'B4', 'C4', 'D4', 'E4'],
                'value' => 'VALUE'
            ]
        ];
    }

    public static function addColorProvider(): array
    {
        return [
            [
                'sheetName' => 'Hoja1',
                'group' => 'A1:E1',
                'cells' => ['A1', 'B1', 'C1', 'D1', 'E1'],
                'value' => 'VALUE',
                'color' => '17FF00'
            ],
            [
                'sheetName' => 'Hoja2',
                'group' => 'A2:E2',
                'cells' => ['A2', 'B2', 'C2', 'D2', 'E2'],
                'value' => 'VALUE',
                'color' => 'FF4200'
            ],
            [
                'sheetName' => 'Hoja3',
                'group' => 'A3:E3',
                'cells' => ['A3', 'B3', 'C3', 'D3', 'E3'],
                'value' => 'VALUE',
                'color' => '007CFF'
            ],
            [
                'sheetName' => 'Hoja4',
                'group' => 'A4:E4',
                'cells' => ['A4', 'B4', 'C4', 'D4', 'E4'],
                'value' => 'VALUE',
                'color' => '9B00FF'
            ]
        ];
    }

    public static function addBackgroundProvider(): array
    {
        return [
            [
                'sheets' => ['Hoja1', 'Hoja2', 'Hoja3', 'Hoja4'],
                'rows' => [
                    [
                        'group' => 'A1:E1',
                        'cells' => ['A1', 'B1', 'C1', 'D1', 'E1'],
                        'value' => 'VALUE',
                        'color' => '17FF00',
                        'fillType' => Fill::FILL_NONE
                    ],
                    [
                        'group' => 'A3:E3',
                        'cells' => ['A3', 'B3', 'C3', 'D3', 'E3'],
                        'value' => 'VALUE',
                        'color' => 'FF4200',
                        'fillType' => Fill::FILL_SOLID
                    ],
                    [
                        'group' => 'A5:E5',
                        'cells' => ['A5', 'B5', 'C5', 'D5', 'E5'],
                        'value' => 'VALUE',
                        'color' => '007CFF',
                        'fillType' => Fill::FILL_GRADIENT_LINEAR
                    ],
                    [
                        'group' => 'A7:E7',
                        'cells' => ['A7', 'B7', 'C7', 'D7', 'E7'],
                        'value' => 'VALUE',
                        'color' => '9B00FF',
                        'fillType' => Fill::FILL_GRADIENT_PATH
                    ],
                    [
                        'group' => 'A9:E9',
                        'cells' => ['A9', 'B9', 'C9', 'D9', 'E9'],
                        'value' => 'VALUE',
                        'color' => '17FF00',
                        'fillType' => Fill::FILL_PATTERN_DARKDOWN
                    ],
                    [
                        'group' => 'A11:E11',
                        'cells' => ['A11', 'B11', 'C11', 'D11', 'E11'],
                        'value' => 'VALUE',
                        'color' => 'FF4200',
                        'fillType' => Fill::FILL_PATTERN_DARKGRAY
                    ],
                    [
                        'group' => 'A13:E13',
                        'cells' => ['A13', 'B13', 'C13', 'D13', 'E13'],
                        'value' => 'VALUE',
                        'color' => '007CFF',
                        'fillType' => Fill::FILL_PATTERN_DARKGRID
                    ],
                    [
                        'group' => 'A15:E15',
                        'cells' => ['A15', 'B15', 'C15', 'D15', 'E15'],
                        'value' => 'VALUE',
                        'color' => '9B00FF',
                        'fillType' => Fill::FILL_PATTERN_DARKHORIZONTAL
                    ],
                    [
                        'group' => 'A17:E17',
                        'cells' => ['A17', 'B17', 'C17', 'D17', 'E17'],
                        'value' => 'VALUE',
                        'color' => '17FF00',
                        'fillType' => Fill::FILL_PATTERN_DARKTRELLIS
                    ],
                    [
                        'group' => 'A19:E19',
                        'cells' => ['A19', 'B19', 'C19', 'D19', 'E19'],
                        'value' => 'VALUE',
                        'color' => 'FF4200',
                        'fillType' => Fill::FILL_PATTERN_DARKUP
                    ],
                    [
                        'group' => 'A21:E21',
                        'cells' => ['A21', 'B21', 'C21', 'D21', 'E21'],
                        'value' => 'VALUE',
                        'color' => '007CFF',
                        'fillType' => Fill::FILL_PATTERN_DARKVERTICAL
                    ],
                    [
                        'group' => 'A23:E23',
                        'cells' => ['A23', 'B23', 'C23', 'D23', 'E23'],
                        'value' => 'VALUE',
                        'color' => '9B00FF',
                        'fillType' => Fill::FILL_PATTERN_GRAY0625
                    ],
                    [
                        'group' => 'A25:E25',
                        'cells' => ['A25', 'B25', 'C25', 'D25', 'E25'],
                        'value' => 'VALUE',
                        'color' => '17FF00',
                        'fillType' => Fill::FILL_PATTERN_GRAY125
                    ],
                    [
                        'group' => 'A27:E27',
                        'cells' => ['A27', 'B27', 'C27', 'D27', 'E27'],
                        'value' => 'VALUE',
                        'color' => 'FF4200',
                        'fillType' => Fill::FILL_PATTERN_LIGHTDOWN
                    ],
                    [
                        'group' => 'A29:E29',
                        'cells' => ['A29', 'B29', 'C29', 'D29', 'E29'],
                        'value' => 'VALUE',
                        'color' => '007CFF',
                        'fillType' => Fill::FILL_PATTERN_LIGHTGRAY
                    ],
                    [
                        'group' => 'A31:E31',
                        'cells' => ['A31', 'B31', 'C31', 'D31', 'E31'],
                        'value' => 'VALUE',
                        'color' => '9B00FF',
                        'fillType' => Fill::FILL_PATTERN_LIGHTGRID
                    ],
                    [
                        'group' => 'A33:E33',
                        'cells' => ['A33', 'B33', 'C33', 'D33', 'E33'],
                        'value' => 'VALUE',
                        'color' => '17FF00',
                        'fillType' => Fill::FILL_PATTERN_LIGHTHORIZONTAL
                    ],
                    [
                        'group' => 'A35:E35',
                        'cells' => ['A35', 'B35', 'C35', 'D35', 'E35'],
                        'value' => 'VALUE',
                        'color' => 'FF4200',
                        'fillType' => Fill::FILL_PATTERN_LIGHTTRELLIS
                    ],
                    [
                        'group' => 'A37:E37',
                        'cells' => ['A37', 'B37', 'C37', 'D37', 'E37'],
                        'value' => 'VALUE',
                        'color' => '007CFF',
                        'fillType' => Fill::FILL_PATTERN_LIGHTUP
                    ],
                    [
                        'group' => 'A39:E39',
                        'cells' => ['A39', 'B39', 'C39', 'D39', 'E39'],
                        'value' => 'VALUE',
                        'color' => '9B00FF',
                        'fillType' => Fill::FILL_PATTERN_LIGHTVERTICAL
                    ],
                    [
                        'group' => 'A41:E41',
                        'cells' => ['A41', 'B41', 'C41', 'D41', 'E41'],
                        'value' => 'VALUE',
                        'color' => '9B00FF',
                        'fillType' => Fill::FILL_PATTERN_MEDIUMGRAY
                    ]
                ]
            ]
        ];
    }

    public static function addDataValidationWithErrorsProvider(): array
    {
        return [
            [
                'data' => [],
                'exceptionMessage' => 'the data configuration is empty',
            ],
            [
                'data' => [
                    'columns' => [],
                ],
                'exceptionMessage' => 'the required columns have not been defined',
            ],
            [
                'data' => [
                    'columns' => ['A1'],
                    'config' => [],
                ],
                'exceptionMessage' => 'the required configuration has not been defined',
            ],
            [
                'data' => [
                    'columns' => ['A1'],
                    'config' => [
                        'error-title' => null,
                    ],
                ],
                'exceptionMessage' => 'error title not defined',
            ],
            [
                'data' => [
                    'columns' => ['A1'],
                    'config' => [
                        'error-title' => 'error-title-xlsx',
                        'error-message' => null,
                    ],
                ],
                'exceptionMessage' => 'error message not defined',
            ],
            [
                'data' => [
                    'columns' => ['A1'],
                    'config' => [
                        'error-title' => 'error-title-xlsx',
                        'error-message' => 'error-message-xlsx',
                        'worksheet' => null,
                    ],
                ],
                'exceptionMessage' => 'spreadsheet not defined',
            ],
            [
                'data' => [
                    'columns' => ['A1'],
                    'config' => [
                        'error-title' => 'error-title-xlsx',
                        'error-message' => 'error-message-xlsx',
                        'worksheet' => 'Data',
                        'column' => null,
                    ],
                ],
                'exceptionMessage' => 'column not defined',
            ],
            [
                'data' => [
                    'columns' => ['A1'],
                    'config' => [
                        'error-title' => 'error-title-xlsx',
                        'error-message' => 'error-message-xlsx',
                        'worksheet' => 'Data',
                        'column' => 'E',
                        'start' => null,
                    ],
                ],
                'exceptionMessage' => 'undefined start',
            ],
            [
                'data' => [
                    'columns' => ['A1'],
                    'config' => [
                        'error-title' => 'error-title-xlsx',
                        'error-message' => 'error-message-xlsx',
                        'worksheet' => 'Data',
                        'column' => 'E',
                        'start' => 2,
                        'end' => null,
                    ],
                ],
                'exceptionMessage' => 'undefined end',
            ],
        ];
    }

    public static function addDataValidationProvider(): array
    {
        return [
            [
                'sheetName' => 'Hoja1',
                'color' => '9B00FF',
                'data' => [
                    'columns' => ['A1', 'B1'],
                    'config' => [
                        'error-title' => 'error-title-xlsx',
                        'error-message' => 'error-message-xlsx',
                        'worksheet' => 'Data',
                        'column' => 'A',
                        'start' => 1,
                        'end' => 10
                    ]
                ]
            ],
            [
                'sheetName' => 'Hoja2',
                'color' => '007CFF',
                'data' => [
                    'columns' => ['B2', 'C2'],
                    'config' => [
                        'error-title' => 'error-title-xlsx',
                        'error-message' => 'error-message-xlsx',
                        'worksheet' => 'Data',
                        'column' => 'B',
                        'start' => 1,
                        'end' => 10
                    ]
                ]
            ],
            [
                'sheetName' => 'Hoja3',
                'color' => 'FF4200',
                'data' => [
                    'columns' => ['C3', 'D3'],
                    'config' => [
                        'error-title' => 'error-title-xlsx',
                        'error-message' => 'error-message-xlsx',
                        'worksheet' => 'Data',
                        'column' => 'C',
                        'start' => 1,
                        'end' => 10
                    ]
                ]
            ],
            [
                'sheetName' => 'Hoja4',
                'color' => '17FF00',
                'data' => [
                    'columns' => ['D4', 'E4'],
                    'config' => [
                        'error-title' => 'error-title-xlsx',
                        'error-message' => 'error-message-xlsx',
                        'worksheet' => 'Data',
                        'column' => 'D',
                        'start' => 1,
                        'end' => 10
                    ]
                ]
            ]
        ];
    }
}
