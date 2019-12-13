<?php
namespace Dhutapratama\MyDocx;

use Dhutapratama\MyDocx\Lib\Engine;

class Docx {
	private $docxPath;
	private $replaceText 	= [];
	private $replaceImage = [];
	private $mergeFile = [];

	public function __construct($docxPath) {
		$this->docxPath = $docxPath;

		if (!file_exists($this->templateFilePath)) {
			throw new \Exception("template file {$this->docxPath} not found");
		}
	}

	public function setText($data) {
		$this->data = $data;
		return $this;
	}

	public function save() {
		$engine = new Engine($this->docxPath);

		// find and replace text
		foreach ($this->data as $key => $value) {
			$engine->findAndReplaceText($key, $value);
		}

		// Find and replace image
		foreach ($this->data as $key => $value) {
			if (!file_exists($value)) {
				throw new \Exception("image file {$value} not found");
			}

			$engine->findAndReplaceImage($key, $value);
		}

		// Merge Documents
		foreach ($this->mergeFile as $value) {
			$engine->addFile($this->value);
		}

		$engine->flush();
		return true;
	}

}