<?php
namespace Dhutapratama\MyDocx\Lib;

class Engine {
    // Path to current docx file
  private $docxPath;

    // Current _RELS data
  private $docxRels;
    // Current DOCUMENT data
  private $docxDocument;
    // Current CONTENT_TYPES data
  private $docxContentTypes;

  private $docxZip;

  private $RELS_ZIP_PATH = "word/_rels/document.xml.rels";
  private $DOC_ZIP_PATH = "word/document.xml";
  private $MEDIA_ZIP_PATH = "word/media/";

  private $CONTENT_TYPES_PATH = "[Content_Types].xml";

  private $ALT_CHUNK_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";
  private $ALT_CHUNK_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

    // Array "zip path" => "content"
  private $headerAndFootersArray = [];

  private $lastRelId = 0;
  private $lastImageId = 0;

  public function __construct($docxPath){
    $this->docxPath = $docxPath;

    $this->docxZip = new TbsZip();
    $this->docxZip->Open($this->docxPath);

    $this->docxRels = $this->readContent($this->RELS_ZIP_PATH);
    $this->docxDocument = $this->readContent($this->DOC_ZIP_PATH);
    $this->docxContentTypes = $this->readContent($this->CONTENT_TYPES_PATH);

    // Read Rels
    $relsXML = new \SimpleXMLElement($this->docxRels);

    // Read Header and Footer
    // Maybe issue because never tested
    foreach ($relsXML as $rel) {
      $relType = [
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
      ];

      if ( in_array($rel["Type"], $relType)) {
        $path = "word/" . $rel["Target"];
        $this->headerAndFootersArray[$path] = $this->readContent($path);
      }
    }
    // Read Max File ID
    foreach ($relsXML->Relationship as $relationship) {
      $number = $this->getNumberFromRelId($relationship['Id']);
      if ($number > $this->lastRelId) {
        $this->lastRelId = $number;
      }
    }
  }

  private function getNumberFromRelId($relId) {
    preg_match('!\d+!', $relId, $matches);
    return (int) $matches[0];
  }

  private function reqRelId() {
    $this->lastRelId++;
    return $this->lastRelId;
  }

  private function reqImageId() {
    $this->lastImageId++;
    return $this->lastImageId;
  }

  private function readContent($zipPath) {
    $content = $this->docxZip->FileRead($zipPath);

    return $content;
  }

  private function writeContent($content, $zipPath) {
    $this->docxZip->FileReplace($zipPath, $content, TBSZIP_STRING);
  }

    // Merge
  public function addFile($filePath){
    $refID = "rId" . $this->reqRelId();
    $file['name'] = "merge_" . $refID . ".docx";
    $file['path'] = "word/merge/" . $file['name'];

    $file['content']  = file_get_contents($filePath);
    $this->docxZip->FileAdd($file['path'], $file['content']);

    $this->addReference($file['name'], $refID);
    $this->addAltChunk($refID);
    $this->addContentType($file['path']);
  }

  private function addReference($fileName, $refID){
    $relXmlString = '<Relationship Target="merge/' . $fileName . '" Type="' . $this->ALT_CHUNK_TYPE . '" Id="' . $refID . '"/>';

    $p = strpos($this->docxRels, '</Relationships>');
    $this->docxRels = substr_replace($this->docxRels, $relXmlString, $p, 0);
  }

  private function addContentType($zipPath) {
    $xmlItem = '<Override ContentType="' . $this->ALT_CHUNK_CONTENT_TYPE . '" PartName="/' . $zipPath . '"/>';

    $p = strpos($this->docxContentTypes, '</Types>');
    $this->docxContentTypes = substr_replace($this->docxContentTypes, $xmlItem, $p, 0);
  }

  private function addAltChunk($refID) {
    $xmlItem = '<w:p><w:r><w:br w:type="page" /></w:r></w:p><w:altChunk r:id="' . $refID . '"/>';

    $p = strpos($this->docxDocument, '</w:body>');
    $this->docxDocument = substr_replace($this->docxDocument, $xmlItem, $p, 0);
  }

    // Replace Text
  public function findAndReplaceText($key, $value) {
    // Search/Replace in document
    $this->docxDocument = str_replace($key, htmlspecialchars($value), $this->docxDocument);

    // Search/Replace in footers and headers
    foreach ($this->headerAndFootersArray as $path => $content) {
      $this->headerAndFootersArray[$path] = str_replace($key, $value, $content);
    }
  }

  // Replace Image
  public function addImage($filePath) {
    $refID = "rId" . $this->reqRelId();
    $file = [
      'mime' => mime_content_type($filePath),
      'ext' => pathinfo($filePath, PATHINFO_EXTENSION),
    ];
    $file['name'] = "image" . $this->reqImageId() . "." . $file['ext'];
    $file['path'] = 'word/media/' . $file['name'];
    $file['content'] = file_get_contents($filePath);

    $this->docxZip->FileAdd($file['path'], $file['content']);

    $this->addRefImg($file['name'], $refID);
    $this->addContentTypeImg($file['ext'], $file['mime']);

    return $refID;
  }

  private function addRefImg($fileName, $refID) {
    $relXmlString = '<Relationship Target="media/' . $fileName . '" ';
    $relXmlString .= 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ';
    $relXmlString .= 'Id="' . $refID . '"/>';

    $p = strpos($this->docxRels, '</Relationships>');
    $this->docxRels = substr_replace($this->docxRels, $relXmlString, $p, 0);
  }

  private function addContentTypeImg($ext, $mime) {
    $xmlItem = '<Default Extension="' . $ext . '" ContentType="' . $mime . '" />';

    $p = strpos($this->docxContentTypes, '</Types>');
    $this->docxContentTypes = substr_replace($this->docxContentTypes, $xmlItem, $p, 0);
  }

  public function findAndReplaceImage($key, $value) {
    $refID = $this->addImage($value);
    $template = file_get_contents(__DIR__ . "/../template/image.xml");
    $replace = str_replace("{rId}", $refID, $template);

    // Search/Replace in document
    $pattern = "/<w:r[ >](?:(?!<w:r>|<\/w:r>).)+" . preg_quote($key, '/') . "(?:(?!<w:r>|<\/w:r>).)+<\/w:r>/s";
    if(preg_match_all($pattern, $this->docxDocument, $matches)) {
      $this->docxDocument = preg_replace($pattern, $replace, $this->docxDocument);
    }
  }

  public function flush() {
    // Save RELS data
    $this->writeContent($this->docxRels, $this->RELS_ZIP_PATH);
    // Save DOCUMENT data
    $this->writeContent($this->docxDocument, $this->DOC_ZIP_PATH);
    // Save CONTENT TYPES data
    $this->writeContent($this->docxContentTypes, $this->CONTENT_TYPES_PATH);
    // Save footers and headers
    foreach ($this->headerAndFootersArray as $path => $content) {
      $this->writeContent($content, $path);
    }

    // Save the merge into a third file
    // We cannot save to current file because it damages ZIP file
    $tempFile = tempnam(dirname($this->docxPath), "dm");

    $this->docxZip->Flush(TBSZIP_FILE, $tempFile);

    // Replace current file with tempFile content
    if (strtoupper(substr(PHP_OS, 0, 3)) === 'WIN') {
      copy($tempFile, $this->docxPath);
      unlink($tempFile);
    } else {
      rename($tempFile, $this->docxPath);
    }
  }
}
