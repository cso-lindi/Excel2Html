<h1 align="center">Welcome to Excel2Html 👋</h1>
<p>
  <a href="#" target="_blank">
    <img alt="License: MIT" src="https://img.shields.io/badge/License-MIT-yellow.svg" />
  </a>
</p>

> Genrate HTML markup from an Excel worksheet that looks perfectly alike

## Install

```sh
composer require cso/excel2html
```

## Usage

```php
$conv = HtmlConverter::fromFilepath(
    'tests/assets/test.xlsx', 
    styleOption: StyleOptions::TABLE_SIZE_FIXED | StyleOptions::WITH_COLUMN_WIDTH | StyleOptions::COLUMN_WIDTH_PROPORTIONAL, 
    worksheetName:'TestTable', 
    columns:['A', 'B', 'C', 'D', 'E', 'F'],
    scale: 1.1);

echo $conv->getHtml();
```

## Run tests

```sh
./vendor/bin/phpunit --testdox tests
```

## Author

👤 **Sebastian Lindemeier**

* Github: [@cso-lindi](https://github.com/cso-lindi)

## Show your support

Give a ⭐️ if this project helped you!

***
_This README was generated with ❤️ by [readme-md-generator](https://github.com/kefranabg/readme-md-generator)_