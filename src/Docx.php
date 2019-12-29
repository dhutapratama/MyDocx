<?php
namespace Dhutapratama\MyDocx;

use Dhutapratama\MyDocx\Lib\Engine;

class Docx {
  private $docxPath;
  private $replaceText = [];
  private $replaceImage = [];
  private $mergeFile = [];

  public function __construct($docxPath) {
    if (!file_exists($docxPath)) {
      throw new \Exception("template file {$this->docxPath} not found");
    }

    $this->docxPath = $docxPath;
  }

  public function setHeaderFooter($replaceHeaderFooter) {
    $this->replaceHeaderFooter = $replaceHeaderFooter;
    return $this;
  }

  public function setText($replaceText) {
    $this->replaceText = $replaceText;
    return $this;
  }

  public function setImage($replaceImage) {
    $this->replaceImage = $replaceImage;
    return $this;
  }

  public function setMerge($mergeFile) {
    $this->mergeFile = $mergeFile;
    return $this;
  }

  public function save() {
    $engine = new Engine($this->docxPath);

    // find and replace header and footer
    $engine->loadHeadersAndFooters();
    foreach ($this->replaceHeaderFooter as $key => $value) {
      $engine->findAndReplaceHeadersAndFooters($key, $value);
    }

    // find and replace text
    foreach ($this->replaceText as $key => $value) {
      $engine->findAndReplaceText($key, $value);
    }

    // Find and replace image
    foreach ($this->replaceImage as $key => $value) {
      if (!file_exists($value)) {
        throw new \Exception("image file {$value} not found");
      }

      $engine->findAndReplaceImage($key, $value);
    }

    // Merge Documents
    foreach ($this->mergeFile as $value) {
      $engine->addFile($value);
    }

    $engine->flush();
    return true;
  }

  public function saveAs($docxPath) {
    if (!copy($this->docxPath, $docxPath)) {
      throw new \Exception("error creating output file {$docxPath}");
    }

    $this->docxPath = $docxPath;
    $this->save();
  }
}
