<?php

namespace Ezijing\Yii2Excel;


class ErrorCode
{
    /**
     * @Message("处理Excel异常")
     */
    public const ERROR = 500;

    /**
     * @Message("Excel参数错误")
     */
    public const PARAMETER_ERROR = 4007;

    /**
     * @Message("导入文件失败")
     */
    public const FAILED_TO_IMPORT_FILES_PROCEDURE = 4008;

    /**
     * @Message("无可导入的数据")
     */
    public const FOR_EXAMPLE_IMPORT_DATA = 4009;
}
