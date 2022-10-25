<?php
namespace Ezijing\Yii2Excel;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use yii\base\Component;
use yii\web\UploadedFile;

/**
 * excel 工具类.
 */
class Excel extends Component implements ExcelInterface
{
    /**
     * @var Spreadsheet
     */
    protected $spreadsheet;

    /**
     * 边框.
     *
     * @var array
     */
    protected $border;

    /**
     * 本地配置.
     */
    protected $config;

    /**
     * 值类型映射.
     *
     * @var string[]
     */
    protected $valueTypeMap = [
        'int', // int
        'string', // 字符串
        'date', // Y-H-d H:i:s
        'time', // 秒级时间戳
        'float', // 转为浮点型
        'function', // 函数
    ];

    /**
     * 文件类型.
     *
     * @var string
     */
    private $_fileType = 'Xlsx';

    public function init()
    {
        parent::init();

        $this->initSpreadsheet();
        $this->initLocalFileDir();
    }

    /**
     * 初始化spreadsheet.
     */
    protected function initSpreadsheet()
    {
        $this->spreadsheet = new Spreadsheet();
        $this->border = [
            'borders' => [
                //外边框
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
                //内边框
                'inside' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];
    }

    /**
     * 初始化保存到本地的文件夹.
     */
    protected function initLocalFileDir()
    {
        $this->config['local_file_address'] = $this->config['local_file_address'] ?? \Yii::getAlias('@webroot') . '/storage/excel';
        $this->pathExists($this->config['local_file_address'] ?? \Yii::getAlias('@webroot') . '/storage/excel');
    }

    private function pathExists($path)
    {
        $this->mkdirs($path);
    }

    private function mkdirs($dir, $mode = 0700)
    {
        if (is_dir($dir) || @mkdir($dir, $mode)) {
            return true;
        }

        if (! mkdir(dirname($dir), $mode)) {
            return false;
        }

        return @mkdir($dir, $mode);
    }

    /**
     * 单元格.
     * @return string[]
     */
    protected function cellMap()
    {
        return ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
    }

    public function exportExcelForASingleSheet(string $tableName, array $data = [])
    {
        Validator::validate($data, [
            'export_way' => 'Required|Int|IntIn:'.implode(',', [ExcelConstant::DOWNLOAD_TO_BROWSER_BY_TMP, ExcelConstant::SAVE_TO_A_LOCAL_DIRECTORY]),
            'enable_number' => 'Bool',
            'titles' => 'Required|Arr|>>>:请设置表头',
            'keys' => 'Required|Arr|ArrLen:' . count($data['titles'] ?? []),
            'data' => 'Arr',
            'value_type' => 'Arr',
            'value_type[*].key' => 'Required|Str',
            'value_type[*].type' => 'Required|Str',
        ]);
        $worksheet = $this->spreadsheet->getActiveSheet();
        // 表头 设置单元格内容
        $enableNumber = false;
        if (isset($data['enable_number']) && $data['enable_number']) {
            $enableNumber = true;
        }

        if ($enableNumber) {
            array_unshift($data['titles'], '序号');
            array_unshift($data['keys'], 'num');
        }

        // 设置工作表标题名称
        $worksheet->setTitle($tableName);
        $cellMap = $this->cellMap();
        $maxCell = $cellMap[count($data['titles']) - 1];
        $worksheet->getStyle('A1:' . $maxCell . 1)->applyFromArray(array_merge($this->border, [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER, // 水平居中对齐
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['argb' => 'fffbeeee'],
            ],
        ]));
        foreach ($data['titles'] as $key => $value) {
            if ($key == '序号') {
                $worksheet->getColumnDimension($cellMap[$key])->setWidth(10);
            } else {
                $worksheet->getColumnDimension($cellMap[$key])->setWidth(25);
            }

            $worksheet->setCellValueExplicitByColumnAndRow($key + 1, 1, $value, 's');
        }

        // 从第二行开始,填充表格数据
        $row = 2;
        // 序号
        $num = 1;

        $valueTypes = $data['value_type'] ?? [];

        foreach ($data['data'] as $item) {
            if ($enableNumber) {
                $item['num'] = $num;
            }
            $worksheet->getStyle("A{$row}:" . $maxCell . $row)->applyFromArray($this->border);
            // 从第一列设置并初始化数据
            foreach ($item as $i => $v) {
                $rowKey = array_search($i, $data['keys']);
                if ($rowKey === false) {
                    continue;
                }

                // 格式化值类型
                if ($valueTypes) {
                    $keys = array_column($valueTypes, 'key');
                    $valueTypes = array_combine($keys, $valueTypes);
                    if (isset($valueTypes[$i])) {
                        $v = $this->formatValue($i, $valueTypes[$i]['type'], $v, $valueTypes[$i]['func'] ?? null);
                    }
                }

                ++$rowKey;
                $worksheet->setCellValueExplicitByColumnAndRow($rowKey, $row, $v, 's');
            }
            ++$num;
            ++$row;
        }

        $fileName = $this->getFileName($tableName);

        return $this->downloadDistributor($data['export_way'], $fileName);
    }

    /**
     * 下载分发器.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws \Psr\Container\ContainerExceptionInterface
     * @throws \Psr\Container\NotFoundExceptionInterface
     * @return \Psr\Http\Message\ResponseInterface|string[]|void
     */
    protected function downloadDistributor(int $exportWay, string $fileName)
    {
        switch ($exportWay) {
            case ExcelConstant::SAVE_TO_A_LOCAL_DIRECTORY:
                return $this->saveToLocal($fileName);
            case ExcelConstant::DOWNLOAD_TO_BROWSER_BY_TMP:
            default:
                return $this->saveToBrowserByTmp($fileName);
        }
    }

    /**
     * 保存到临时文件再从浏览器自动下载到本地.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws \Psr\Container\ContainerExceptionInterface
     * @throws \Psr\Container\NotFoundExceptionInterface
     * @return \Psr\Http\Message\ResponseInterface
     */
    protected function saveToBrowserByTmp(string $fileName)
    {
        $localFileName = $this->getLocalUrl($fileName);
        $writer = IOFactory::createWriter($this->spreadsheet, $this->_fileType);

        // 保存到临时文件下
        $writer->save($localFileName);

        // 将文件转为字符串
        $content = file_get_contents($localFileName);

        // 删除临时文件
        unlink($localFileName);


        header("Content-type: application/octet-stream");
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('content-description', "attachment;filename={$fileName}");
        header("Content-type: application/octet-stream");
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        echo $content;
        exit();
    }

    /**
     * 保存到服务器本地.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @return string[]
     */
    protected function saveToLocal(string $fileName)
    {
        $url = $this->getLocalUrl($fileName);
        $writer = @IOFactory::createWriter($this->spreadsheet, $this->_fileType);
        $writer->save($url);
        $this->spreadsheet->disconnectWorksheets();
        unset($this->spreadsheet);

        return [
            'path' => $url,
            'filename' => $fileName,
        ];
    }

    /**
     * 获取保存到本地的excel地址.
     *
     * @return string
     */
    protected function getLocalUrl(string $fileName)
    {
        return $this->config['local_file_address'] . DIRECTORY_SEPARATOR . $fileName;
    }

    /**
     * 设置文件名.
     *
     * @return string
     */
    protected function getFileName(string $tableName)
    {
        return sprintf('%s_%s.xlsx', $tableName, date('Ymd_His'));
    }

    /**
     * 格式化值的类型.
     *
     * @param $key 键
     * @param $type 值类型
     * @param $value 值
     * @param null $func 回调函数
     * @return false|int|mixed|string
     */
    protected function formatValue($key, $type, $value, $func = null)
    {
        $typeMap = array_flip($this->valueTypeMap);
        if (! isset($typeMap[strtolower($type)])) {
            throw new ExcelException(\Yii::t('app', 'value_type.*.type error'), ErrorCode::PARAMETER_ERROR);
        }

        switch (strtolower($type)) {
            case 'int':
                $value = (int) $value;
                break;
            case 'string':
                $value = (string) trim($value);
                break;
            case 'date':
                $value = date('Y-m-d H:i:s', (int) $value);
                break;
            case 'time':
                $value = strtotime((string) $value);
                break;
            case 'function':
                if ($func) {
                    $value = $func($value);
                }
                break;
            default:
        }

        return $value;
    }


    public function exportExcelWithMultipleSheets(string $tableName, array $data = [])
    {
        Validator::validate($data, [
            'export_way' => 'Required|Int|IntIn:' . implode(',', [ExcelConstant::DOWNLOAD_TO_BROWSER_BY_TMP, ExcelConstant::SAVE_TO_A_LOCAL_DIRECTORY]),
            'sheets_params' => 'Required|Arr',
            'sheets_params[*].sheet_title' => 'Required|Str',
            'sheets_params[*].enable_number' => 'Required|Bool',
            'sheets_params[*].titles' => 'Required|Arr',
            'sheets_params[*].keys' => 'Required|Arr',
            'sheets_params[*].data' => 'Required|Arr',
            'sheets_params[*].value_type' => 'Arr',
            'sheets_params[*].value_type[*].key' => 'Str',
            'sheets_params[*].value_type[*].type' => 'Str',
        ]);

        $firstSheet = true;
        foreach ($data['sheets_params'] as $sheetParamsValue) {
            $enableNumber = false;
            if (isset($sheetParamsValue['enable_number']) && $sheetParamsValue['enable_number']) {
                $enableNumber = true;
            }

            if ($enableNumber) {
                array_unshift($sheetParamsValue['titles'], '序号');
                array_unshift($sheetParamsValue['keys'], 'num');
            }

            if ($firstSheet) {
                $worksheet = $this->spreadsheet->getActiveSheet();
            } else {
                $worksheet = $this->spreadsheet->createSheet();
            }
            $firstSheet = false;

            // 设置工作表名称
            $worksheet->setTitle($sheetParamsValue['sheet_title']);
            $cellMap = $this->cellMap();
            $maxCell = $cellMap[count($sheetParamsValue['titles']) - 1];
            $worksheet->getStyle('A1:' . $maxCell . 1)->applyFromArray(array_merge($this->border, [
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_CENTER, // 水平居中对齐
                ],
                'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'startColor' => ['argb' => 'fffbeeee'],
                ],
            ]));

            // 表头设置单元格内容
            foreach ($sheetParamsValue['titles'] as $titleKey => $titleValue) {
                // 设置列宽
                if ($titleKey == '序号') {
                    $worksheet->getColumnDimension($cellMap[$titleKey])->setWidth(10);
                } else {
                    $worksheet->getColumnDimension($cellMap[$titleKey])->setWidth(25);
                }
                $worksheet->setCellValueExplicitByColumnAndRow($titleKey + 1, 1, $titleValue, 's');
            }

            $row = 2;
            // 序号
            $num = 1;

            $valueTypes = $sheetParamsValue['value_type'] ?? [];

            foreach ($sheetParamsValue['data'] as $item) {
                if ($enableNumber) {
                    $item['num'] = $num;
                }
                $worksheet->getStyle("A{$row}:" . $maxCell . $row)->applyFromArray($this->border);
                // 从第一列设置并初始化数据
                foreach ($item as $i => $v) {
                    $rowKey = array_search($i, $sheetParamsValue['keys']);
                    if ($rowKey === false) {
                        continue;
                    }

                    // 格式化值类型
                    if ($valueTypes) {
                        $keys = array_column($valueTypes, 'key');
                        $valueTypes = array_combine($keys, $valueTypes);
                        if (isset($valueTypes[$i])) {
                            $v = $this->formatValue($i, $valueTypes[$i]['type'], $v, $valueTypes[$i]['func'] ?? null);
                        }
                    }

                    ++$rowKey;
                    $worksheet->setCellValueExplicitByColumnAndRow($rowKey, $row, $v, 's');
                }
                ++$num;
                ++$row;
            }
        }
        // 默认打开第一个sheet
        $this->spreadsheet->setActiveSheetIndex(0);
        $fileName = $this->getFileName($tableName);

        return $this->downloadDistributor($data['export_way'], $fileName);
    }

    public function importExcelForASingleSheet(array $data = []): array
    {
        Validator::validate($data, [
            'import_way' => 'Required|Int|IntIn:' . implode(',', [ExcelConstant::THE_LOCAL_IMPORT, ExcelConstant::BROWSER_IMPORT]),
            'file_path' => 'Str|Regexp:/(?:[x|X][l|L][s|S][x|X])$/',
            'titles' => 'Required|Arr',
            'titles[*]' => 'Required|Str',
            'keys' => 'Required|Arr',
            'keys[*]' => 'Required|Str',
            'value_type' => 'Arr',
            'value_type[*].key' => 'Str',
            'value_type[*].type' => 'Str',
        ]);

        // 强制titles和keys的数量保持一致
        if (count($data['titles']) != count($data['keys'])) {
            throw new ExcelException(\Yii::t('app', 'titles和keys要保持对应'), ErrorCode::PARAMETER_ERROR);
        }

        return $this->importDistributor($data);
    }

    /**
     * 导入分发器.
     *
     * @param array $data 数据
     * @throws \Psr\Container\ContainerExceptionInterface
     * @throws \Psr\Container\NotFoundExceptionInterface
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return array
     */
    protected function importDistributor(array $data)
    {
        switch ($data['import_way']) {
            case ExcelConstant::THE_LOCAL_IMPORT:
                return $this->importExcelForASingleSheetLocal($data);
            default:
                return $this->importExcelForASingleSheetBrowser($data);
        }
    }

    /**
     * 从浏览器导入单个sheet的excel.
     *
     * @param array $data 数据
     * @throws \Psr\Container\ContainerExceptionInterface
     * @throws \Psr\Container\NotFoundExceptionInterface
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return array
     */
    protected function importExcelForASingleSheetBrowser(array $data)
    {
        $file = $this->importFileRequetVerify();

        return $this->readExcelForASingleSheet($file->tempName, $data);
    }

    /**
     * 通过接口请求获取file.
     * @return UploadedFile|null
     */
    protected function importFileRequetVerify()
    {
        $request = \Yii::$app->getRequest();
        if (! $request->isPost) {
            throw new ExcelException(\Yii::t('app', '只接收POST请求'), ErrorCode::FOR_EXAMPLE_IMPORT_DATA);
        }
        $model = new UploadForm();
        $model->file = UploadedFile::getInstance($model, 'file');
        if (! $model->file) {
            throw new ExcelException(\Yii::t('app', sprintf('导入文件失败')), ErrorCode::FAILED_TO_IMPORT_FILES_PROCEDURE);
        }

        // 获取文件上传的临时文件
        if (! isset($model->file->tempName)) {
            throw new ExcelException(\Yii::t('app', '导入文件失败, 找不到临时目录！'), ErrorCode::FAILED_TO_IMPORT_FILES_PROCEDURE);
        }

        return $model->file;
    }


    /**
     * 从本地导入多个sheet的excel.
     *
     * @param array $data 数据
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return array
     */
    protected function importExcelForASingleSheetLocal(array $data)
    {
        $filePath = $data['file_path'];

        return $this->readExcelForASingleSheet($filePath, $data);
    }

    /**
     * 格式化单个sheet的excel.
     *
     * @param string $filePath 文件路径
     * @param array $data 数据
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return array
     */
    protected function readExcelForASingleSheet(string $filePath, array $data)
    {
        $objRead = IOFactory::createReader($this->_fileType);

        $inputFileType = IOFactory::identify($filePath);
        if (strtolower($inputFileType) != 'xlsx') {
            throw new ExcelException(\Yii::t('app', '只支持导入Xlsx文件'), ErrorCode::PARAMETER_ERROR );
        }

        if (! $objRead->canRead($filePath)) {
            throw new ExcelException(\Yii::t('app', '只支持导入Excel文件'), ErrorCode::FAILED_TO_IMPORT_FILES_PROCEDURE);
        }

        $objRead->setReadDataOnly(true);
        $objRead->setReadEmptyCells(false);

        $spreadsheet = @$objRead->load($filePath);

        if (empty($list = $spreadsheet->getSheet(0)->toArray())) {
            return [];
        }

        // 强制第一行必须是表头，无导入的数据
        if (count($list) <= 1) {
            return [];
        }

        // 获取表头、非正常格式不作处理
        if (empty($headers = $list[0])) {
            throw new ExcelException(\Yii::t('app', '无可导入的数据'), ErrorCode::FOR_EXAMPLE_IMPORT_DATA);
        }

        // 匹配自定义的titles和keys
        $titleMap = array_combine($data['titles'], $data['keys']);

        // 获取指定key的映射结果
        $keyMap = [];
        foreach ($headers as $headerIndex => $header) {
            if ($header && isset($titleMap[trim((string)$header)])) {
                $keyMap[$headerIndex] = $titleMap[trim((string)$header)];
            }
        }
        if (empty($keyMap)) {
            return [];
        }

        // 去除表头
        array_shift($list);

        // 初始化格式化的数据
        $formatData = [];

        $valueTypes = $data['value_type'] ?? [];

        // 开启携程，格式化数据
        foreach ($list as $index => &$item) {
            foreach ($keyMap as $keyIndex => $key) {
                $value = (string) ($item[$keyIndex] ?? '');
                // 格式化值类型
                if ($valueTypes) {
                    $keys = array_column($valueTypes, 'key');
                    $valueTypes = array_combine($keys, $valueTypes);
                    if (isset($valueTypes[$key])) {
                        $value = $this->formatValue($key, $valueTypes[$key]['type'], $value, $valueTypes[$key]['func'] ?? null);
                    }
                }
                $formatData[$index][$key] = $value;
            }
            if (empty(implode('', $formatData[$index]))) {
                unset($formatData[$index]);
            }
        }

        return $formatData;
    }

    public function import(array $data): array
    {
        Validator::validate($data, [
            'import_way' => 'Required|Int|IntIn:' . implode(',', [ExcelConstant::DOWNLOAD_TO_BROWSER_BY_TMP, ExcelConstant::SAVE_TO_A_LOCAL_DIRECTORY]),
            'file_path' => 'Str|Regexp:/(?:[x|X][l|L][s|S][x|X])$/',
        ]);

        switch ($data['import_way']) {
            case ExcelConstant::THE_LOCAL_IMPORT:
                $filePath = $data['file_path'];
                break;
            default:
                $file = $this->importFileRequetVerify();
                $filePath = $file['tmp_file'];
        }

        if (file_exists($filePath)) {
            $inputFileType = \PhpOffice\PhpSpreadsheet\IOFactory::identify($filePath);
            $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($inputFileType);
            $reader->setReadDataOnly(true);
            $spreadsheetReader = $reader->load($filePath);
            $sheets = $spreadsheetReader->getAllSheets();
            foreach ($sheets as $sheet) {
                $data[] = $sheet->toArray();
            }
            return $data;
        } else {
            throw new ExcelException(\Yii::t('app', '文件写入错误'), ErrorCode::FAILED_TO_IMPORT_FILES_PROCEDURE);
        }

    }
}
