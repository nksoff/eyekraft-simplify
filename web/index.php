<?php

use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\JsonResponse;
use Symfony\Component\HttpFoundation\Response;

require('../vendor/autoload.php');

$app = new Silex\Application();
$app['debug'] = true;

$app->register(new Silex\Provider\MonologServiceProvider(), array(
    'monolog.logfile' => 'php://stderr',
));
$app->register(new Silex\Provider\TwigServiceProvider(), array(
    'twig.path' => __DIR__.'/views',
));


$app->get('/', function() use($app) {
    return $app['twig']->render('index.twig');
});

$app->post('/handle-checks', function(Request $request) use($app) {
    $app['monolog']->addDebug('checks');

    $fileInfo = $request->files->get('file');

    if(!$fileInfo || !$fileInfo->isValid()) {
        return new Response('Не удалось загрузить файл', 400);
    }

    $data = file($fileInfo->getPathname());

    //////////////////////////////////////////////////////////////
    $headerRow = null;
    $headerRowKey = null;
    foreach($data as $key=>$row) {
        if(preg_match('/-- -/', $row)) {
            $headerRow = $row;
            $headerRowKey = $key;
            break;
        }
    }

    if(!$headerRow) {
        return new Response('Не удалось найти заголовок', 400);
    }


    $headerSpaces = array_keys(
        array_filter(
            str_split($headerRow), function($char) {
                return $char === ' ';
            })
        );

    $columns = array();
    $prev = 0;
    foreach($headerSpaces as $spaceKey) {
        $columns[] = array($prev, $spaceKey);
        $prev = $spaceKey;
    }

    $headerNamesRows = array(
        iconv('cp1251', 'utf8', $data[$headerRowKey - 2]),
        iconv('cp1251', 'utf8', $data[$headerRowKey - 1]),
    );

    $headerNames = array();

    foreach($columns as $column) {
        $tmp = array();
        foreach($headerNamesRows as $row) {
            $tmp[] = trim(iconv_substr($row, $column[0], $column[1] - $column[0]));
        }

        $headerNames[] = trim(implode(' ', $tmp));
    }
    //////////////////////////////////////////////////////////////

    $rows = array_filter($data, function($row) {
        return preg_match('/\//', $row);
    });

    $rows = array_map(function($row) use($columns) {
        $res = array();

        foreach($columns as $column) {
            $res[] = trim(substr($row, $column[0],  $column[1] - $column[0]));
        }

        return $res;
    }, $rows);

    $excel = new PHPExcel;
    $excelList = $excel->setActiveSheetIndex(0);
    $excelList->setTitle('List');

    $rowNum = 1;
    $columnNum = 'A';

    $formatLikeSumColumns = array();

    foreach($headerNames as $name) {
        $excelList->setCellValue($columnNum . $rowNum, $name);

        if(preg_match('/сумма/iu', $name)) {
            $formatLikeSumColumns[] = $columnNum;
        }

        $columnNum++;
    }

    $rowNum++;

    foreach($rows as $row) {
        $columnNum = 'A';
        foreach($row as $rowColumn) {
            $type = PHPExcel_Cell_DataType::TYPE_STRING;
            if(in_array($columnNum, $formatLikeSumColumns)) {
                $type = PHPExcel_Cell_DataType::TYPE_NUMERIC;
                $rowColumn = (float) $rowColumn;
            }

            $excelList->setCellValueExplicit($columnNum . $rowNum, $rowColumn, $type);
            $columnNum++;
        }
        $rowNum++;
    }

    $fileName = md5('excel_' . time() . rand()) . '.xls';
    $filePath = 'data/' . $fileName;
    PHPExcel_IOFactory::createWriter($excel, 'Excel5')->save($filePath);

    $data = [
        'fileLink' => '/' . $filePath,
    ];
    return new JsonResponse($data);
});

$app->run();
