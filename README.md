yii2-excel
=====================================

安装
----------------------------

### 3、配置文件 `main-local.php`
```php
<?php

declare(strict_types=1);

return [
    'components' => [
        'excel' => 'components\excel\Excel',
    ],
];
```

使用
---------------------------
### 1、导出单个sheet的excel到本地
```php
$tableName = 'test';
$data = [
    'export_way' => ExcelConstant::SAVE_TO_A_LOCAL_DIRECTORY, // 导出方式
    'enable_number' => true, // 是否开启序号
    'titles' => ['ID', '用户名', '部门', '职位'], // 设置表头
    'keys' => ['id', 'username', 'department', 'position'], // 设置表头标识，必须与要导出的数据的key对应
    // 要导出的数据
    'data' => [
        ['id' => '1', 'username' => '小明', 'department' => '运营部', 'position' => '产品运营'],
        ['id' => '2', 'username' => '小王', 'department' => '技术部', 'position' => 'PHP'],
    ],
    // 验证规则, 本地导入也适用
    'value_type' => [
        // 强转string
        ['key' => 'position', 'type' => 'string'],
        // 强转int
        ['key' => 'id', 'type' => 'int'],
        // 回调处理
        [
            'key' => 'department',
            'type' => 'function',
            'func' => function($value) {
                return (string) $value;
            },
        ],
    ],
];

$res = \Yii::$app->excel->exportExcelForASingleSheet($tableName, $data);
```

### 2、从浏览器导出单个sheet的excel
```php
<?php

declare(strict_types=1);

class ExcelController extends Controller
{
    /**
     * @var Excel
     */
    protected $excel;

    public function init()
    {
        $this->excel = \Yii::$app->excel;
    }

    public function download()
    {
        $tableName = 'test';
        $data = [
            'export_way' => ExcelConstant::DOWNLOAD_TO_BROWSER_BY_TMP,
            'enable_number' => false,
            'titles' => ['ID', '用户名', '部门', '职位'],
            'keys' => ['id', 'username', 'department', 'position'],
            'data' => [
                ['id' => '1', 'username' => '小明', 'department' => '运营部', 'position' => '产品运营'],
                ['id' => '2', 'username' => '小王', 'department' => '技术部', 'position' => 'PHP'],
            ],
            // 验证规则, 本地导入也适用
            'value_type' => [
                // 强转string
                ['key' => 'position', 'type' => 'string'],
                // 强转int
                ['key' => 'id', 'type' => 'int'],
                // 回调处理
                [
                    'key' => 'department',
                    'type' => 'function',
                    'func' => function($value) {
                        return (string) $value;
                    },
                ],
            ],
        ];

        return $this->excel->exportExcelForASingleSheet($tableName, $data);
    }
}
```

### 3、导出多个sheet的excel到本地
```php
$tableName = 'sheets';
$data = [
    'export_way' => ExcelConstant::SAVE_TO_A_LOCAL_DIRECTORY,
    'sheets_params' => [
        [
            'sheet_title' => '企业1',
            'enable_number' => true, // 是否开启序号
            'titles' => ['ID', '用户名', '部门', '职位'],
            'keys' => ['id', 'username', 'department', 'position'],
            'data' => [
                ['id' => '1', 'username' => '小明', 'department' => '运营部', 'position' => '产品运营'],
                ['id' => '2', 'username' => '小王', 'department' => '技术部', 'position' => 'PHP'],
            ],
            // 验证规则, 本地导入也适用
            'value_type' => [
                // 强转string
                ['key' => 'position', 'type' => 'string'],
                // 强转int
                ['key' => 'id', 'type' => 'int'],
                // 回调处理
                [
                    'key' => 'department',
                    'type' => 'function',
                    'func' => function($value) {
                        return (string) $value;
                    },
                ],
            ]
        ],
        [
            'sheet_title' => '企业2',
            'titles' => ['ID', '用户名', '部门', '职位'],
            'keys' => ['id', 'username', 'department', 'position'],
            'data' => [
                ['id' => '3', 'username' => '小李', 'department' => '运营部', 'position' => '产品运营'],
                ['id' => '4', 'username' => '小赵', 'department' => '技术部', 'position' => 'PHP'],
            ],
        ],
        [
            'sheet_title' => '部门',
            'enable_number' => false, // 是否开启序号
            'titles' => ['ID', '部门', '职位'],
            'keys' => ['id', 'department', 'position'],
            'data' => [
                ['id' => 1, 'department' => '运营部', 'position' => '产品运营'],
                ['id' => 2, 'department' => '技术部', 'position' => 'PHP'],
            ],
        ],
    ]
];

$res = \Yii::$app->excel->exportExcelWithMultipleSheets($tableName, $data);
print_r($res);
```

### 4、从浏览器导出多个sheet的excel
```php
<?php

class ExcelController extends Controller
{
    /**
     * @var Excel
     */
    protected $excel;

    public function init()
    {
        $this->excel = \Yii::$app->excel;
    }

    public function download()
    {
        $tableName = 'sheets';
        $data = [
            'export_way' => ExcelConstant::DOWNLOAD_TO_BROWSER_BY_TMP,
            'sheets_params' => [
                [
                    'sheet_title' => '企业1',
                    'enable_number' => true,
                    'titles' => ['ID', '用户名', '部门', '职位'],
                    'keys' => ['id', 'username', 'department', 'position'],
                    'data' => [
                        ['id' => '1', 'username' => '小明', 'department' => '运营部', 'position' => '产品运营'],
                        ['id' => '2', 'username' => '小王', 'department' => '技术部', 'position' => 'PHP'],
                    ],
                    // 验证规则, 本地导入也适用
                    'value_type' => [
                        // 强转string
                        ['key' => 'position', 'type' => 'string'],
                        // 强转int
                        ['key' => 'id', 'type' => 'int'],
                        // 回调处理
                        [
                            'key' => 'department',
                            'type' => 'function',
                            'func' => function($value) {
                                return (string) $value;
                            },
                        ],
                    ],
                ],
                [
                    'sheet_title' => '企业2',
                    'titles' => ['ID', '用户名', '部门', '职位'],
                    'keys' => ['id', 'username', 'department', 'position'],
                    'data' => [
                        ['id' => '3', 'username' => '小李', 'department' => '运营部', 'position' => '产品运营'],
                        ['id' => '4', 'username' => '小赵', 'department' => '技术部', 'position' => 'PHP'],
                    ],
                ],
                [
                    'sheet_title' => '部门',
                    'titles' => ['ID', '部门', '职位'],
                    'keys' => ['id', 'department', 'position'],
                    'data' => [
                        ['id' => 1, 'department' => '运营部', 'position' => '产品运营'],
                        ['id' => 2, 'department' => '技术部', 'position' => 'PHP'],
                    ],
                ],
            ]
        ];

        return $this->excel->exportExcelWithMultipleSheets($tableName, $data);
    }
}
```

### 5、导入单个sheet的excel

#### 测试数据
|ID|用户名|部门|职位|
|:---: |:----:|:----:|:-----:|
|1|小明|运营部|产品运营|
|2|小王|技术部|PHP|


#### (1) 本地导入
##### 代码
```php
$data = [
    // 带入方式
    'import_way' => ExcelConstant::THE_LOCAL_IMPORT,
    // 文件路径
    'file_path' => '/Users/ezijing/php_project/hyperf-demo/storage/excel/test_20220113_105250.xlsx',
    // 指定导入的title
    'titles' => ['部门', 'ID'],
    // 指定生成的key
    'keys' => ['position', 'id'],
];

$res = $this->excel->importExcelForASingleSheet($data);

print_r($res);
```
##### 导入结果
```
Array (
    [0] => Array
        (
            [id] => 1
            [position] => 运营部
        )

    [1] => Array
        (
            [id] => 2
            [position] => 技术部
        )

)
```

#### (2) 接口导入
##### 请求方式
> `POST`
##### 请求的类型
> `form-data`

##### 请求参数
|参数|类型|
|:---: |:----:|
|file|text|

##### 代码
```php
<?php

declare(strict_types=1);

namespace App\Controller;

use Ezijing\HyperfExcel\Core\Constants\ExcelConstant;
use Ezijing\HyperfExcel\Core\Services\Excel;
use Hyperf\HttpServer\Annotation\AutoController;

/**
 * @AutoController
 */
class ExcelController extends Controller
{
    /**
     * @var Excel
     */
    protected $excel;

    public function init()
    {
         $this->excel = \Yii::$app->excel;
    }

    public function import()
    {
        $data = [
            'import_way' => ExcelConstant::BROWSER_IMPORT,
            // 指定导入的title
            'titles' => ['部门', 'ID', '职位'],
            // 指定生成的key
            'keys' => ['department', 'id', 'position'],
            // 验证规则, 本地导入也适用
            'value_type' => [
                // 强转string
                ['key' => 'position', 'type' => 'string'],
                // 强转int
                ['key' => 'id', 'type' => 'int'],
                // 回调处理
                [
                    'key' => 'department',
                    'type' => 'function',
                    'func' => function($value) {
                        return (string) $value;
                    },
                ],
            ]
        ];

        return $this->excel->importExcelForASingleSheet($data);
    }
}

```

##### 导入结果
```json
[
    {
        "id": 1,
        "department": "运营部",
        "position": "产品运营"
    },
    {
        "id": 2,
        "department": "技术部",
        "position": "PHP"
    }
]
```
