Word-Render
===========================
CLI tool for text replacements in MS Word documents and PDF exporter*

**PDF export require Windows platform with MS Word installed*

## Requirements
- ZipArchive class (PHP 5 >= 5.2.0, PHP 7, PECL zip >= 1.1.0)
- COM Class (The COM class allows you to instantiate an OLE compatible COM object (such as MS Word) and call its methods and access its properties) - required only to export pdf files using MS Word

This COM extension requires ``php_com_dotnet.dll`` to be enabled inside of ``php.ini``

``extension=php_com_dotnet.dll``

## Usage
``php word-render.php [options]``

```
Options:
-v, --verbose           Verbose mode
-h, --help              Help
-i, --input             Input file [required]
-o, --output            Output file (docx, pdf) [required]
-r, --replace           Placeholders with values (placeholder1=>value|placeholder2=>value...) or json file
```

**Export docx to pdf using MS Word enigne**

``php word-render.php -i input-file.docx -o output-file.pdf``

**Fill placeholders in file and export to pdf (inline)**

``php word-render.php -i input-file.docx -o output-file.pdf -r "{place1}=>value1|{place2}=>value2"``

**Fill placeholders in file and export to pdf (replacements from json file)**

``php word-render.php -i input-file.docx -o output-file.pdf -r replacements.json``

Example of replacements.json file

``
{"{place1}":"value1","{place2}":"value2"}
``
