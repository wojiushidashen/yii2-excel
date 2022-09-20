<?php


namespace Ezijing\Yii2Excel;


use yii\base\Model;
use yii\web\UploadedFile;

class UploadForm extends Model
{
    /**
     * @var UploadedFile file attribute
     */
    public $file;

    /**
     * @return array the validation rules.
     */
    public function rules()
    {
        return [
            [['file'], 'file', 'extensions' => 'xlsx'],
        ];
    }

    public function formName()
    {
        return '';
    }
}
