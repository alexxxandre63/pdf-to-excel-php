
<?php
require __DIR__ . '/vendor/autoload.php';

use Smalot\PdfParser\Parser;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (!isset($_FILES['arquivo_pdf'])) {
    die('Nenhum arquivo enviado.');
}

$arquivo_temporario = $_FILES['arquivo_pdf']['tmp_name'];
$parser = new Parser();
$pdf = $parser->parseFile($arquivo_temporario);
$texto = $pdf->getText();

function extrair($texto, $regex) {
    preg_match($regex, $texto, $match);
    return $match[1] ?? '';
}

$dados = [
    'Nome' => extrair($texto, '/Nome:\s*(.*?)\s*CPF:/'),
    'CPF' => extrair($texto, '/CPF:\s*([\d\.\-]+)/'),
    'Endereço' => extrair($texto, '/Endereço:\s*(.*?)\s*Bairro:/'),
    'Bairro' => extrair($texto, '/Bairro:\s*(.*?)\s*Cidade:/'),
    'Cidade' => extrair($texto, '/Cidade:\s*(.*?)\s*UF:/'),
    'UF' => extrair($texto, '/UF:\s*(.*?)\s*CEP:/'),
    'CEP' => extrair($texto, '/CEP:\s*([\d\-]+)/'),
    'Email' => extrair($texto, '/E-mail:\s*(.*?)\s/'),
    'Placa' => extrair($texto, '/Placa:\s*([A-Z0-9]+)/'),
    'Modelo' => extrair($texto, '/Veículo:\s*(.*?)\s*Ce/'),
    'Ano Modelo' => extrair($texto, '/Ano Modelo:\s*(\d+)/'),
    'Ano Fabricação' => extrair($texto, '/Ano De Fabricação:\s*(\d+)/'),
    'Valor Devolvido' => extrair($texto, '/devolução no valor de R\$ ([\d,]+)/')
];

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$linha = 1;
foreach ($dados as $campo => $valor) {
    $sheet->setCellValue("A$linha", $campo);
    $sheet->setCellValue("B$linha", $valor);
    $linha++;
}

$arquivo_excel = 'dados_extraidos.xlsx';
$writer = new Xlsx($spreadsheet);
$writer->save($arquivo_excel);

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . $arquivo_excel . '"');
readfile($arquivo_excel);
unlink($arquivo_excel);
exit;
?>
