<?php


namespace Ezijing\Yii2Excel;


class ExcelConstant
{
    /**
     * @Message("保存到本地")
     */
    public const SAVE_TO_A_LOCAL_DIRECTORY = 1;

    /**
     * @Message("缓存到临时文件并下载到浏览器")
     */
    public const DOWNLOAD_TO_BROWSER_BY_TMP = 2;

    /**
     * @Message("本地导入")
     */
    public const THE_LOCAL_IMPORT = 1;

    /**
     * @Message("浏览器导入")
     */
    public const BROWSER_IMPORT = 2;
}
