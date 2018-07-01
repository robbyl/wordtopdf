<?php
//Word Doc to PDF using Com
ini_set("com.allow_dcom", "true");

$doc_path = "G:/word/";

try {
    $word = new com('word.application') or die('MS Word could not be loaded');
} catch (com_exception $e) {
    $nl = "<br />";
    echo $e->getMessage() . $nl;
    echo $e->getCode() . $nl;
    echo $e->getTraceAsString();
    echo $e->getFile() . " LINE: " . $e->getLine();
    $word->Quit();
    $word = null;
    die;
}

$word->Visible = 0;
$word->DisplayAlerts = 0;

try {
    $doc = $word->Documents->Open($doc_path . 'large.docx');
} catch (com_exception $e) {
    $nl = "<br />";
    echo $e->getMessage() . $nl;
    echo $e->getCode() . $nl;
    echo $e->getFile() . " LINE: " . $e->getLine();
    $word->Quit();
    $word = null;
    die;
}
echo "doc opened";
try {
    $doc->ExportAsFixedFormat($doc_path . "test_image.pdf", 17, false, 0, 0, 0, 0, 7, true, true, 2, true, true, false);

} catch (com_exception $e) {
    $nl = "<br />";
    echo $e->getMessage() . $nl;
    echo $e->getCode() . $nl;
    echo $e->getTraceAsString();
    echo $e->getFile() . " LINE: " . $e->getLine();
    $word->Quit();
    $word = null;
    die;
}

echo "created pdf";
$word->Quit();
$word = null;
