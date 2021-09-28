<?php
  require 'vendor/autoload.php';

  //読み込むファイルとシート名を指定
  //以降、worksheetオブジェクトからExcelファイルを操作
  $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
  $workbook = $reader->load("sample.xlsx");
  $worksheet = $workbook->getSheetByName('Sheet1');

  //保存用のオブジェクトを作成
  $ws = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($workbook);

  //行列データの取得 ※A1, B1 などセル番地を作って値を取得する
  foreach ($worksheet->getRowIterator() as $row) {
    foreach($worksheet->getColumnIterator() as $col) {
      echo $worksheet->getCell($col->getColumnIndex() .
      $row->getRowIndex())->getValue().PHP_EOL;
    }
  }

  //A列の値を取得する
  foreach ($worksheet->getRowIterator() as $row) {
    foreach($worksheet->getColumnIterator() as $col) {
      $col_index = $col->getColumnIndex();
      if ($col_index == 'A') {
        echo $worksheet->getCell($col->getColumnIndex() .
        $row->getRowIndex())->getValue().PHP_EOL;
      }
    }
  }

  //A列のデータを格納
  $dummy_array = array();
  foreach ($worksheet->getRowIterator() as $row) {
    foreach($worksheet->getColumnIterator() as $col) {
      $col_index = $col->getColumnIndex();
      if ($col_index == 'A') {
        $dumyy_array[] = $worksheet->getCell($col->getColumnIndex() .
        $row->getRowIndex())->getValue().PHP_EOL;
      }
    }
  }

  //予定出発時間を格納
  $STD_array = array();
  foreach ($worksheet->getRowIterator() as $row) {
    foreach($worksheet->getColumnIterator() as $col) {
      $col_index = $col->getColumnIndex();
      if ($col_index == 'D') {
        $STD_array[] = $worksheet->getCell($col->getColumnIndex() .
        $row->getRowIndex())->getValue().PHP_EOL;
      }
    }
  }
  //列名削除
  $STD_NAME = $STD_array[0];
  if (in_array($STD_NAME, $STD_array, true)) {
    array_shift($STD_array);
  }

  //出発時間を格納
  $ATD_array = array();
  foreach ($worksheet->getRowIterator() as $row) {
    foreach($worksheet->getColumnIterator() as $col) {
      $col_index = $col->getColumnIndex();
      if ($col_index == 'F') {
        $ATD_array[] = $worksheet->getCell($col->getColumnIndex() .
        $row->getRowIndex())->getValue().PHP_EOL;
      }
    }
  }
  //列名削除
  $ATD_NAME = $ATD_array[0];
  if (in_array($ATD_NAME, $ATD_array, true)) {
    array_shift($ATD_array);
  }

  //for文を使って時間を比較し、結果を新しい配列に格納する
  $counter = count($STD_array);
  $dep_result = array();
  for ($i=0; $i < $counter; $i++) {
    if ($STD_array[$i] < $ATD_array[$i]) {
      $dep_result[] = array('☓');
    }else {
      $dep_result[] = array('◯');
    }
  }
  //結果を反映
  array_unshift($dep_result, array('定時出発'));
  $worksheet->fromArray($dep_result, null, 'I1');

  //予定到着時間を格納
  $STA_array = array();
  foreach ($worksheet->getRowIterator() as $row) {
    foreach($worksheet->getColumnIterator() as $col) {
      $col_index = $col->getColumnIndex();
      if ($col_index == 'E') {
        $STA_array[] = $worksheet->getCell($col->getColumnIndex() .
        $row->getRowIndex())->getValue().PHP_EOL;
      }
    }
  }
  //列名削除
  $STA_NAME = $STA_array[0];
  if (in_array($STA_NAME, $STA_array, true)) {
    array_shift($STA_array);
  }

  //到着時間を格納
  $ATA_array = array();
  foreach ($worksheet->getRowIterator() as $row) {
    foreach($worksheet->getColumnIterator() as $col) {
      $col_index = $col->getColumnIndex();
      if ($col_index == 'G') {
        $ATA_array[] = $worksheet->getCell($col->getColumnIndex() .
        $row->getRowIndex())->getValue().PHP_EOL;
      }
    }
  }
  //列名削除
  $ATA_NAME = $ATA_array[0];
  if (in_array($ATA_NAME, $ATA_array, true)) {
    array_shift($ATA_array);
  }

  //for文を使って時間を比較し、結果を新しい配列に格納する
  $counter = count($STA_array);
  $arr_result = array();
  for ($i=0; $i < $counter; $i++) {
    if ($STA_array[$i] < $ATA_array[$i]) {
      $arr_result[] = array('☓');
    }else {
      $arr_result[] = array('◯');
    }
  }
  //結果を反映
  array_unshift($arr_result, array('定時到着'));
  $worksheet->fromArray($arr_result, null, 'J1');

  //乗客数を取得
  $PAX_array = array();
  foreach ($worksheet->getRowIterator() as $row) {
    foreach($worksheet->getColumnIterator() as $col) {
      $col_index = $col->getColumnIndex();
      if ($col_index == 'H') {
        $PAX_array[] = $worksheet->getCell($col->getColumnIndex() .
        $row->getRowIndex())->getValue().PHP_EOL;
      }
    }
  }
  //列名削除
  $PAX_NAME = $PAX_array[0];
  if (in_array($PAX_NAME, $PAX_array, true)) {
    array_shift($PAX_array);
  }

  //乗客数が40人以上であれば◎、そうでなければ☓を格納する
  $pax_result = array();
  foreach ($PAX_array as $i) {
    if ($i >= 40){
      $pax_result[] = array("◎");
    }else {
      $pax_result[] = array("☓");
    }
  }
  array_unshift($pax_result, array('40人以上'));

  //fromArrayを使って値を反映
  $worksheet->fromArray($pax_result, null, 'K1');

  //乗車率を計算
  $LOAD_FACTOR = array();
  $ave_LF = 0;
  foreach ($PAX_array as $i) {
    $result = ($i / 45) * 100;
    $ave_LF += $result;
    $result = number_format($result, 1) . '%';
    $LOAD_FACTOR[] = array($result);
  }
  array_unshift($LOAD_FACTOR, array('40名以上'));

  //fromArrayを使って値を反映
  $worksheet->fromArray($LOAD_FACTOR, null, 'L1');

  /////サマリーデータを反映/////

  //定時出発率
  $dep_count = 0;
  foreach ($dep_result as $i) {
    if ($i[0] == '◯') {
      $dep_count += 1;
    }
  }
  $dep_ave = $dep_count / (count($dep_result) - 1);
  $dep_ave = $dep_ave * 100 . '%';
  $worksheet->setCellValue('M1', '定時出発率');
  $worksheet->setCellValue('M2', $dep_ave);

  //定時到着率
  $arr_count = 0;
  foreach ($arr_result as $i) {
    if ($i[0] === '◯') {
      $arr_count += 1;
    }
  }
  $arr_ave = $arr_count / (count($arr_result) - 1);
  $arr_ave = $arr_ave * 100 . '%';
  $worksheet->setCellValue('N1', '定時到着率');
  $worksheet->setCellValue('N2', $arr_ave);

  //平均乗車率
  $ave_lf = $ave_LF / (count($LOAD_FACTOR) -1);
  $ave_lf = number_format($ave_lf, 1) . '%';
  $worksheet->setCellValue('O1', '平均乗車率');
  $worksheet->setCellValue('O2', $ave_lf);

  //路線ごとの運行数をカウント
  $route_array = [];
  foreach ($worksheet->getRowIterator() as $row) {
    foreach ($worksheet->getColumnIterator() as $col) {
      $col_index = $col->getColumnIndex();
      if ($col_index == 'C') {
        $route_array[] = $worksheet->getCell($col->getColumnIndex() .
        $row->getRowIndex())->getValue().PHP_EOL;
      }
    }
  }
  $route_array_header = $route_array[0];
  if (in_array($route_array_header, $route_array, true)) {
    array_shift($route_array);
  }
  $route_num = array_count_values($route_array);
  $RN = array();
  foreach ($route_num as $key => $val) {
    $RN[] = array($key, $val);
  }
  $route_num_header = array(['区間', '運行数']);
  $worksheet->fromArray($route_num_header, null, 'P1');
  $worksheet->fromArray($RN, null, 'P2');

  /////棒グラフを反映/////
  //グラフ生成に必要な機能の呼び出し
  use PhpOffice\PhpSpreadsheet\Chart\Chart;
  use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
  use PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;
  use PhpOffice\PhpSpreadsheet\Chart\Layout;
  use PhpOffice\PhpSpreadsheet\Chart\Legend;
  use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
  use PhpOffice\PhpSpreadsheet\Chart\Title;

  //X軸ラベルの指定  新宿-新潟～川越的場-新潟
  $X_LABELS = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING,
                    'Sheet1!$P$2:$P$5', null, 4), );

  //プロットデータの指定　17～12
  $X_DATA = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER,
                  'Sheet1!$Q$2:$Q$5', null, 4), );

  //データシリーズを用意する
  $data_series = new DataSeries(DataSeries::TYPE_BARCHART,     //plotTYpe
                                DataSeries::GROUPING_STANDARD, //plotGrouping
                                range(0, count($X_LABELS) - 1),//plotOrder
                                [],                            //plotLabels->今回はなし
                                $X_LABELS,                     //plotCategories
                                $X_DATA                        //plotValues
                                );

  //PlotAreaにデータシリーズを設定
  $plotArea = new PlotArea(null, array($data_series));

  //タイトルを設定
  $title = new Title('区間別運行数');

  //チャートオブジェクトを生成
  $chart = new Chart('sample_bar_chart', $title, null, $plotArea, true, 0, null, null);
  /*
  ↑↑の()左から 「name」「title」「legend」「plotArea」「plotVisibleOnly」「displayBlanksAs」
  「xAxisLabel」「yAxisLabel」それぞれに該当する引数を指定する。
  */

  //グラフを反映する位置を指定　M8～Q24の範囲にグラフを反映する
  $chart->setTopLeftPosition('M8');
  $chart->setBottomRightPosition('Q24');

  //ワークシートにチャートオブジェクトを追加
  $worksheet->addChart($chart);
  $ws->setIncludeCharts(true);

  //保存;
  $ws->save('result.xlsx');

 ?>
