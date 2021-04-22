<form name="form1" method="post" action=""><h1>成绩前10名学生导出功能</h1>
    <p>请输入需要导出的年级和班级，若只需要查询年级前10名，请留空班级信息</p>
    <h2>
        年级：<input style="width:120px;height:30px" type="number" name="grade" max="9" min="7" step="1"
                  placeholder="年级范围：7-9" required>
    </h2>
    <h2>
        班级：<input style="width:120px;height:30px" type="number" name="class_no" max="10" min="1" step="1"
                  placeholder="班级范围：1-10">
    </h2>
    <p>
        <input name="submit" type="submit" id="submit" value="提交并下载数据"/>
    </p>
</form>

<?php
include './lib/PHPExcel.php';
include './lib/PHPExcel/Writer/Excel2007.php';
include './MysqlDB.php';
error_reporting(0);
$config = array(
    'dbname'=>'visiting'
);
$con = MysqlDB::getInstance($config);

if (isset($_POST["submit"])) {
    $grade = $_POST[grade];
    $class_no = $_POST[class_no];

    if ($class_no != null) {
        $sql = <<<EOF
select indexing,
       name,
       sex,
       age,
       grade,
       class,
       entrance,
       ranking,
       total,
       term1_Chinese,
       term1_math,
       term1_English,
       term1_history,
       term1_geography,
       term1_chemistry,
       term2_Chinese,
       term2_math,
       term2_English,
       term2_history,
       term2_geography,
       term2_chemistry
from (select name,
             sex,
             age,
             grade,
             class,
             entrance,
             s.id                                                                                         as indexing,
             term1_Chinese + term1_math + term1_English + term1_history +
             term1_geography + term1_chemistry + term2_Chinese +
             term2_math + term2_English + term2_history +
             term2_geography + term2_chemistry                                                            as total,
             dense_rank() over (partition by grade,class order by term1_Chinese + term1_math + term1_English +
                                                                  term1_history +
                                                                  term1_geography + term1_chemistry + term2_Chinese +
                                                                  term2_math + term2_English + term2_history +
                                                                  term2_geography + term2_chemistry desc) as ranking,
             term1_Chinese,
             term1_math,
             term1_English,
             term1_history,
             term1_geography,
             term1_chemistry,
             term2_Chinese,
             term2_math,
             term2_English,
             term2_history,
             term2_geography,
             term2_chemistry
      from (select id,
                   Chinese   as term1_Chinese,
                   math      as term1_math,
                   English   as term1_English,
                   histories as term1_history,
                   geography as term1_geography,
                   chemistry as term1_chemistry
            from scores c1
            where term = 1) term1
               left join (select id,
                                 Chinese   as term2_Chinese,
                                 math      as term2_math,
                                 English   as term2_English,
                                 histories as term2_history,
                                 geography as term2_geography,
                                 chemistry as term2_chemistry
                          from scores c2
                          where term = 2) term2
                         on term1.id = term2.id
               right join students s
                          on term1.id = s.id) r
where ranking <= 10
  and grade = %d
  and class = %d;
EOF;
        $sql = sprintf($sql, $grade, $class_no);
    } else {
        $sql = <<<EOF
select indexing,
       name,
       sex,
       age,
       grade,
       class,
       entrance,
       ranking,
       total,
       term1_Chinese,
       term1_math,
       term1_English,
       term1_history,
       term1_geography,
       term1_chemistry,
       term2_Chinese,
       term2_math,
       term2_English,
       term2_history,
       term2_geography,
       term2_chemistry
from (select name,
             sex,
             age,
             grade,
             class,
             entrance,
             s.id                                                                                         as indexing,
             term1_Chinese + term1_math + term1_English + term1_history +
             term1_geography + term1_chemistry + term2_Chinese +
             term2_math + term2_English + term2_history +
             term2_geography + term2_chemistry                                                            as total,
             dense_rank() over (partition by grade order by term1_Chinese + term1_math + term1_English +
                                                                  term1_history +
                                                                  term1_geography + term1_chemistry + term2_Chinese +
                                                                  term2_math + term2_English + term2_history +
                                                                  term2_geography + term2_chemistry desc) as ranking,
             term1_Chinese,
             term1_math,
             term1_English,
             term1_history,
             term1_geography,
             term1_chemistry,
             term2_Chinese,
             term2_math,
             term2_English,
             term2_history,
             term2_geography,
             term2_chemistry
      from (select id,
                   Chinese   as term1_Chinese,
                   math      as term1_math,
                   English   as term1_English,
                   histories as term1_history,
                   geography as term1_geography,
                   chemistry as term1_chemistry
            from scores c1
            where term = 1) term1
               left join (select id,
                                 Chinese   as term2_Chinese,
                                 math      as term2_math,
                                 English   as term2_English,
                                 histories as term2_history,
                                 geography as term2_geography,
                                 chemistry as term2_chemistry
                          from scores c2
                          where term = 2) term2
                         on term1.id = term2.id
               right join students s
                          on term1.id = s.id) r
where ranking <= 10
  and grade = %d;
EOF;
        $sql = sprintf($sql, $grade);
    }

    $fileheader = array('id', 'name', 'sex', 'age', 'grade', 'class', 'entrance', 'ranking', 'total', 'term1_Chinese', 'term1_math', 'term1_English', 'term1_history', 'term1_geography', 'term1_chemistry', 'term2_Chinese', 'term2_math', 'term2_English', 'term2_history', 'term2_geography', 'term2_chemistry');

    $data = $con->getAll($sql, $fileheader);

    $objPHPExcel = new PHPExcel();
    $objSheet = $objPHPExcel->getActiveSheet();
    $objSheet->setTitle('Sheet1');
    $objSheet->fromArray($data);
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel2007');
    ob_end_clean();
    header("Pragma: public");
    header("Expires: 0");
    header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
    header("Content-Type:application/force-download");
    header("Content-Type:application/vnd.ms-excel");
    header("Content-Type:application/octet-stream");
    header("Content-Type:application/download");;
    header('Content-Disposition:attachment;filename="result.xlsx"');
    header('Cache-Control: max-age=0');
    header("Content-Transfer-Encoding:binary");
    $objWriter->save('php://output');

    //exportExcel($data, $fileheader);
}

//导出
/**
 * @throws PHPExcel_Writer_Exception
 * @throws PHPExcel_Exception
 * @throws PHPExcel_Reader_Exception
 */
function exportExcel($data, $fileheader){
    $excel = new PHPExcel();
    $objActSheet = $excel->getActiveSheet();
    $letter = array('A','B','C','D','E','F','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
    $excel->setActiveSheetIndex(0);
    $objActSheet->setTitle('Sheet1');
    for($i = 0;$i < count($fileheader);$i++) {
        //设置表头值
        $objActSheet->setCellValue("$letter[$i]1",$fileheader[$i]);
        //设置表头字体样式
        //$objActSheet->getStyle("$letter[$i]1")->getFont()->setName('微软雅黑');
        //设置表头字体大小
        //$objActSheet->getStyle("$letter[$i]1")->getFont()->setSize(12);
        //设置表头字体是否加粗
        $objActSheet->getStyle("$letter[$i]1")->getFont()->setBold(true);
        //设置表头文字垂直居中
        $objActSheet->getStyle("$letter[$i]1")->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        //设置文字上下居中
        $objActSheet->getStyle($letter[$i])->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
        //设置表头外的文字垂直居中
        $excel->setActiveSheetIndex(0)->getStyle($letter[$i])->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    }
    //单独设置入学时间列宽度为15
    $objActSheet->getColumnDimension('F')->setWidth(15);
    for ($i = 2;$i <= count($data) + 1;$i++) {
        $j = 0;
        foreach ($data[$i - 2] as $key=>$value) {
            $objActSheet->setCellValue("$letter[$j]$i",$value);
            $j++;
        }
        $objActSheet->getRowDimension($i)->setRowHeight('80px');
    }

    $objWriter = PHPExcel_IOFactory::createWriter($excel,'Excel2007');
    ob_end_clean();
    header("Pragma: public");
    header("Expires: 0");
    header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
    header("Content-Type:application/force-download");
    header("Content-Type:application/vnd.ms-excel");
    header("Content-Type:application/octet-stream");
    header("Content-Type:application/download");;
    header('Content-Disposition:attachment;filename="result.xlsx"');
    header('Cache-Control: max-age=0');
    header("Content-Transfer-Encoding:binary");
    $objWriter->save('php://output');
}

