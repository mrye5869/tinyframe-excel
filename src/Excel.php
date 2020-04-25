<?php
// +----------------------------------------------------------------------
// | zibi [ WE CAN DO IT MORE SIMPLE]
// +----------------------------------------------------------------------
// | Copyright (c) 2016-2019 http://xmzibi.com All rights reserved.
// +----------------------------------------------------------------------
// | Author：MrYe       <email：55585190@qq.com>
// +----------------------------------------------------------------------

namespace og\excel;

use og\error\ToolException;
use og\helper\Arr;

class Excel
{

    /**
     * 对应的excel类型
     * @var array
     */
    protected $types = [
        'excel5'        => 'Excel5',
        'excel2007'     => 'Excel2007',
    ];

    /**
     * excel类型
     * @var string
     */
   protected  $type;

    /**
     * 初始化
     * Excel constructor.
     * @param string $type
     * @param string $sdkPath
     * @throws ToolException
     */
   public function __construct($type = 'Excel2007', $sdkPath = '')
   {
        $this->setType($type);

        $this->setSdkPath($sdkPath);
   }

    /**
     * 设置skd路径
     * @param $sdkPath
     * @return $this
     */
    public function setSdkPath($sdkPath)
    {
        if($sdkPath) {
            //加载自定义sdk
            if(!is_file($sdkPath)) {
                //sdk文件不存在，抛出异常

                throw new ToolException('qrcode sdk file does not exist:'.$sdkPath);
            }

            include $sdkPath;

        } else {
            //加载微擎内置exce类库
            include  env('root_path').'framework/library/phpexcel/PHPExcel.php';
        }

        if(!class_exists('PHPExcel')) {
            //类不存在，抛出异常

            throw new ToolException('PHPExcel class does not exist');
        }

        return $this;
    }

    /**
     * 设置类型
     * @param $type
     * @return $this
     */
    public function setType($type)
    {
        if(in_array($type, $this->types)) {
            //设置类型
            $this->type = $type;
        }

        return $this;
    }

    /**
     * 设置单个sheet
     * @param \PHPExcel $objPHPExcel
     * @param $sheet
     * @param $data
     * @param array $fields
     * @param string $titleName
     */
    protected function sheetExport($objPHPExcel, $sheet, $data, $fields = [], $titleName = '')
    {

        if($sheet == 0) {
            //第一个不必创建，从0开始即可
            $ActiveSheet = $objPHPExcel->getSheet();

        } else {
            //创建sheet
            $ActiveSheet = $objPHPExcel->createSheet($sheet);
        }
        //表头字段
        $fields = $this->getFields($fields, $data);

        //设置表头
        $i = 0;
        foreach($fields as $key => $field) {
            $ck = $this->num2alpha($i ++) . '1';
            $ActiveSheet->setCellValue($ck, $field);
        }

        //设置表数据
        $newData = $this->getData($data);
        foreach($newData as $key => $val) {
            //设置数据
            $key = $key + 1;
            $num = $key + 1;
            $ii = 0;
            foreach($fields as $field => $value) {
                if(isset($val[$field])) {
                    $ck = $this->num2alpha($ii ++) . $num;
                    $ActiveSheet->setCellValue($ck, $val[$field]);
                }
            }
        }

        return $ActiveSheet->setTitle($this->getTitleName($sheet, $titleName));
    }

    /**
     * 导出excel数据报表
     * @param array $data
     * @param array $fields
     * @param string $titleName
     * @param string $fileName
     * @param string $author
     * @throws ToolException
     * @throws \PHPExcel_Reader_Exception
     * @throws \PHPExcel_Writer_Exception
     */
   public function export($data, $fields = [], $titleName = '', $fileName = '', $author = 'MrYe')
   {
       $dimension = Arr::dimension($data);
       if($dimension < 2 || $dimension > 3) {
           //data数据的数组维数只能在2-3之间

           throw new ToolException('There is no exportable data');
       }

       $objPHPExcel = new \PHPExcel();
       //创建人
       $objPHPExcel->getProperties()->setCreator($author)
            //最后修改人
            ->setLastModifiedBy($author)
            //关键字
            ->setKeywords("excel")
            //种类
            ->setCategory("result file");

        if($dimension == 2) {
            //给数组加维数
            $data = [0 => $data];
            $fields = [0 => $fields];
        }
        //开始生成excel
        try {

            $sheeti = 0;
            foreach ($data as $key => $item) {
                //sheeti逐个生成
                $field = isset($fields[$sheeti]) ? $fields[$sheeti] : false;
                $this->sheetExport($objPHPExcel, $sheeti, $item, $field, $titleName);

                $sheeti ++;
            }

        } catch (\Exception $exception) {
            //继续抛出异常
            throw new ToolException($sheeti.'st export error:'.$exception->getMessage());
        }

        header('Content-Type: application/vnd.ms-excel');
        header("Content-Disposition: attachment;filename=".$this->getFileName($fileName));
        header('Cache-Control: max-age=0');
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, $this->type);
        $objWriter->save('php://output');

        exit();
   }

    /**
     * 读取行数据
     * @param \PHPExcel_Worksheet $sheet
     * @param bool $isDocking
     * @return array
     * @throws \PHPExcel_Exception
     */
   protected function sheetRead($sheet, $isDocking)
   {

       $images = [];
       foreach ($sheet->getDrawingCollection() as $img) {
           //获取图片所在行和列
           list($column, $row)= \PHPExcel_Cell::coordinateFromString($img->getCoordinates());
           $images[$row][$column] = $img->getPath();
       }

       $data = [];
       foreach ($sheet->toArray(null, false, false, true) as $row => $item) {
           //读取数据
           if(!empty($images[$row])) {
               $item = array_merge($item, $images[$row]);
           }

           $data[$row] = $item;
       }

       if($isDocking) {
           //执行对应操作
           $field = array_shift($data);
           foreach ($data as $key => $val)
           {
               //行
               foreach ($val as $k => $v)
               {
                   //列
                   if ($field[$k]) {
                       //赋值
                       $result[$key][$field[$k]] = $v;
                   }
               }
           }

           return $result;

       } else {

           return $data;
       }

   }

    /**
     * 读取excel中的数据
     * @param string $path
     * @param int|array|string $sheeti
     * @param bool $isDocking
     * @return array
     * @throws ToolException
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     */
   public function read($path, $sheeti = 0, $isDocking = false)
   {

       if(!is_file($path)) {
           //读取的文件不存在

           throw new ToolException('file does not exist:'.$path);
       }
       $excel = \PHPExcel_IOFactory::load($path);


       if(is_array($sheeti)) {
           //数组
           foreach ($sheeti as $i) {
               $sheets[$i] = $excel->getSheet($i);
           }

       } elseif($sheeti == 'all') {
           //全部
           $sheets = $excel->getAllSheets();

       } else {
           //单个
           $sheeti = (int)$sheeti;
           $sheets = [$excel->getSheet($sheeti)];
       }

       $result = [];
       try {

           foreach ($sheets as $i => $sheet) {
               $result[$i] = $this->sheetRead($sheet, $isDocking);
           }

       } catch (\Exception $exception) {
           //继续抛出异常
           throw new ToolException($i.'st red error:'.$exception->getMessage());
       }

       return is_int($sheeti) ? Arr::first($result) : $result;
   }


    /**
     * 生成excel的列字段
     * @param $index
     * @param int $start
     * @return string
     */
    protected function num2alpha($index, $start = 65)
    {
        $str = '';
        if (floor($index / 26) > 0) {
            $str .= $this->num2alpha(floor($index / 26)-1);
        }

        return $str . chr($index % 26 + $start);
    }

    /**
     * 获取excel中的标题
     * @param int $sheet
     * @param string $titleName
     * @return string
     */
    protected function getTitleName($sheet = 0, $titleName)
    {

        $titleName = !empty($titleName) ? $titleName : '数据报表';

        return $sheet == 0 ? $titleName : $titleName.($sheet + 1);
    }

    /**
     *  获取导出的excel文件名称
     * @param $fileName
     * @param $ext
     * @return string
     */
    protected function getFileName($fileName)
    {
        if($this->type == 'Excel2007') {
            $ext = '.xlsx';
        } else {
            $ext = '.xls';
        }
        if(empty($fileName)) {
            //默认名称
            $fileName = date('Y-m-d').'数据报表'.'.'.$ext;

        } else {
            if (!strexists($fileName, $ext)) {

                $fileName .= $ext;
            }

        }

        return $fileName;
    }

    /**
     * 获取导出数据中的字段
     * @param $fields
     * @param array $data
     * @return mixed
     */
    protected function getFields($fields, $data = [])
    {
        if(empty($fields)) {
            foreach ($data as $key => $value) {
                foreach ($value as $field => $val) {
                    $fields[$field] = $field;
                }
                break;
            }
        }

        return $fields;
    }

    /**
     * 获取新的数据
     * @param array $data
     * @return array
     */
    protected function getData($data = [])
    {
        $result = [];
        foreach ($data as $key => $item) {
            foreach ($item as &$value) {
                if(is_numeric($value)) {
                    $value .= "\t";
                }
            }
            $result[] = $item;
        }
        unset($value);

        return $result;
    }

}