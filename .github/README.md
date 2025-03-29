# 🦁 Lion-Spreadsheet

<p align="center">
  <a href="https://dev.lion-packages.com/docs/library/content" target="_blank">
    <img 
        src="https://github.com/lion-packages/framework/assets/56183278/60871c9f-1c93-4481-8c1e-d70282b33254"
        width="450" 
        alt="Lion-packages Logo"
    >
  </a>
</p>

<p align="center">
  <a href="https://packagist.org/packages/lion/spreadsheet">
    <img src="https://poser.pugx.org/lion/spreadsheet/v" alt="Latest Stable Version">
  </a>
  <a href="https://packagist.org/packages/lion/spreadsheet">
    <img src="https://poser.pugx.org/lion/spreadsheet/downloads" alt="Total Downloads">
  </a>
  <a href="https://github.com/lion-packages/spreadsheet/blob/main/LICENSE">
    <img src="https://poser.pugx.org/lion/spreadsheet/license" alt="License">
  </a>
  <a href="https://www.php.net/">
    <img src="https://poser.pugx.org/lion/spreadsheet/require/php" alt="PHP Version Require">
  </a>
</p>

🚀 **Lion-Spreadsheet** Library to facilitate the use of the spreadsheet.

---

## 📖 Features

✔️ Create XLSX files.
✔️ Read and edit existing files.
✔️ Apply styles and formats.  

---

## 📦 Installation

Install the spreadsheet using **Composer**:

```bash
composer require phpoffice/phpspreadsheet lion/spreadsheet
```

## Usage Example

```php
<?php

use Lion\Spreadsheet\Spreadsheet;

$spreadsheet = new Spreadsheet();

$spreadsheet>load('file.xlsx');

$spreadsheet->setCell('A2', 'value');

$spreadsheet->save();
```

## 📝 License

The <strong>spreadsheet</strong> is open-sourced software licensed under the [MIT License](https://github.com/lion-packages/spreadsheet/blob/main/LICENSE).
