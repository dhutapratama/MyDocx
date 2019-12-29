# MyDocx

PHP Library for merge, find some text and replace to text or image in docx document. Note that this library is never tested for OpenOffice format.

This library is downloaded from **[https://github.com/jupitern/docx](https://github.com/jupitern/docx)** since no update from Dec 2016. Issue are very welcome to report here.

## Features

- Find text and replace with text and image
- Merge docx files on one file
###### Note : Merge document here is include new file into existing file and show as one document or existing page is not modified

## Requirements

 - PHP 5.4 +

## Installation

MyDocx is installed via [Composer](https://getcomposer.org/).
To [add a dependency](https://getcomposer.org/doc/04-schema.md#package-links) to MyDocx in your project, either

Run the following to use the latest stable version
```sh
    composer require dhutapratama/mydocx
```

You can of course also manually edit your composer.json file
```json
{
    "require": {
       "dhutapratama/mydocx": "v1.0.*"
    }
}
```

## Getting started
#### Declaration
```php
use Dhutapratama\MyDocx\Docx;

// Initialization
$myDocx = new Docx('/mydir/template.docx');
```
#### Replacing Header and/or Footer
```php
$myDocx->setHeaderFooter(['text_to_find' => 'value to replace'])
  ->save();
```
#### Replacing Text
```php
$myDocx->setText(['text_to_find' => 'value to replace'])
  ->save();
```
#### Replacing Image
```php 
$myDocx->setImage(['text_to_find' => '/your/image.png'])
  ->save();
```

#### Merge Files
```php 
$myDocx->setMerge(['/your/file1.docx', '/your/file2.docx'])
  ->save();
```

#### Replace and merge
```php
$myDocx->setText(['text_to_find' => 'value to replace'])
  ->setImage(['text_to_find' => '/your/image.png'])
  ->setMerge(['/your/file1.docx', '/your/file2.docx'])
  ->save();
```

## Contributing

Please report any issue or you can also help others to resolving issues by fork and requesting merge to master branch.