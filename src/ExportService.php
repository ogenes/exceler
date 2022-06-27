<?php
/**
 * User: ogenes <ogenes.yi@gmail.com>
 * Date: 2022/6/18
 */

namespace Ogenes\Exceler;

use PhpOffice\PhpSpreadsheet\Chart\Properties;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;

class ExportService extends Base
{
    /**
     * @var array
     */
    protected $headerFont = [
        'name' => '微软雅黑',
        'color' => ['argb' => '5E5E5E'],
        'bold' => true,
        'size' => 8
    ];
    
    /**
     * @var array[]
     */
    protected $headerBorders = [
        'allBorders' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => ['argb' => '00595959'],
        ],
    ];
    
    /**
     * @var array
     */
    protected $headerAlignment = [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
    ];
    
    /**
     * @var array
     */
    protected $headerFill = [];
    
    /**
     * @var array
     */
    protected $font = [
        'name' => '微软雅黑',
        'color' => ['argb' => '00595959'],
        'size' => 9
    ];
    
    /**
     * @var array[]
     */
    protected $borders = [
        'allBorders' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => ['argb' => '00595959'],
        ],
    ];
    
    /**
     * @var array
     */
    protected $alignment = [
        'horizontal' => Alignment::HORIZONTAL_LEFT,
        'vertical' => Alignment::VERTICAL_CENTER,
        'wrapText' => false,
    ];
    
    /**
     * @var array
     */
    protected $fill = [];
    
    /**
     * @var string
     */
    protected $formatCode = Properties::FORMAT_CODE_GENERAL;
    
    /**
     * @var bool
     */
    protected $freezeHeader = false;
    
    /**
     * @var bool
     */
    protected $autoFilter = false;
    
    /**
     * @var int
     */
    protected $width = 30;
    /**
     * @var string
     */
    protected $unit = '';
    
    /**
     * @var string
     */
    protected $filename = '';
    /**
     * @var string
     */
    protected $filepath = './file/';
    
    
    /**
     * @desc 设置字体
     * [
     * 'name' => '微软雅黑', //字体名称
     * 'color' => ['argb' => '00595959'], //颜色
     * 'size' => 9 // 字号
     * ]
     *
     * @param array $font
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setStyleFont(array $font): self
    {
        $this->font = array_merge($this->font, $font);
        return $this;
    }
    
    /**
     * @desc 设置边框
     * [
     * 'allBorders' => [
     * 'borderStyle' => Border::BORDER_THIN,
     * 'color' => ['argb' => '00595959'],
     * ],
     * ]
     *
     * @param array $borders
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setStyleBorders(array $borders): self
    {
        $this->borders = array_merge($this->borders, $borders);
        return $this;
    }
    
    /**
     * @desc 设置内容区 alignment
     *  [
     * 'horizontal' => Alignment::HORIZONTAL_CENTER, //水平居中
     * 'vertical' => Alignment::VERTICAL_CENTER, //垂直居中
     * 'wrapText' => false, // 是否自动换行
     * ]
     *
     * @param array $alignment
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setStyleAlignment(array $alignment): self
    {
        $this->alignment = array_merge($this->alignment, $alignment);
        return $this;
    }
    
    /**
     * @desc 设置内容区填充
     *
     * @param array $fill
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/19
     */
    public function setStyleFill(array $fill): self
    {
        $this->fill = array_merge($this->fill, $fill);
        return $this;
    }
    
    /**
     * @desc 设置内容区单元格格式
     *
     * @param string $formatCode
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setStyleFormatCode(string $formatCode): self
    {
        $this->formatCode = $formatCode;
        return $this;
    }
    
    /**
     * @desc 设置默认跨度， 默认 30
     *
     * @param float $width
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setStyleWidth(float $width): self
    {
        $width > 0 && $this->width = $width;
        return $this;
    }
    
    /**
     * @desc 设置通用单位
     * 有效单位有 pt (points), px (pixels), pc (pica), in (inches), cm (centimeters) and mm (millimeters).
     *
     * @param string $unit
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setStyleUnit(string $unit): self
    {
        $this->unit = $unit;
        return $this;
    }
    
    /**
     * @desc 设置表头字体
     * [
     * 'name' => '微软雅黑', //字体名称
     * 'color' => ['argb' => '5E5E5E'], // 字体颜色
     * 'bold' => true, // 加粗
     * 'size' => 8 // 字号
     * ]
     *
     * @param array $font
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setStyleHeaderFont(array $font): self
    {
        $this->headerFont = array_merge($this->headerFont, $font);
        return $this;
    }
    
    /**
     * @desc 设置表头边框
     * [
     * 'allBorders' => [
     * 'borderStyle' => Border::BORDER_THIN, //边框，参考 PhpOffice\PhpSpreadsheet\Style\Border
     * 'color' => ['argb' => '00595959'], // 颜色
     * ],
     * ]
     *
     * @param array $borders
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setStyleHeaderBorders(array $borders): self
    {
        $this->headerBorders = array_merge($this->headerBorders, $borders);
        return $this;
    }
    
    /**
     * @desc 设置表头的alignment
     *
     * [
     * 'horizontal' => Alignment::HORIZONTAL_CENTER, //水平居中
     * 'vertical' => Alignment::VERTICAL_CENTER, //垂直居中
     * 'wrapText' => false, // 是否自动换行
     * ]
     *
     * @param array $alignment
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setStyleHeaderAlignment(array $alignment): self
    {
        $this->headerAlignment = array_merge($this->headerAlignment, $alignment);
        return $this;
    }
    
    /**
     * @desc 设置表头填充
     *
     * @param array $fill
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/19
     */
    public function setStyleHeaderFill(array $fill): self
    {
        $this->headerFill = array_merge($this->headerFill, $fill);
        return $this;
    }
    
    /**
     * @desc 是否冻结首行
     *
     * @param bool $freeze
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/20
     */
    public function setFreezeHeader(bool $freeze): self
    {
        $this->freezeHeader = $freeze;
        return $this;
    }
    
    /**
     * @desc 自动筛选
     *
     * @param bool $autoFilter
     * @return $this
     *
     * @author: ogenes <ogenes.yi@gmail.com>
     * @date: 2022/6/20
     */
    public function setAutoFilter(bool $autoFilter): self
    {
        $this->autoFilter = $autoFilter;
        return $this;
    }
    
    /**
     * @desc 导出文件名， 不需要后缀， 仅支持导出 xlsx 格式文件
     *
     * @param string $filename
     * @return $this
     *
     * @author: ogenes<ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setFilename(string $filename): self
    {
        $this->filename = $filename;
        return $this;
    }
    
    public function getFilename(): string
    {
        if (empty($this->filename)) {
            $this->filename = md5(time() . random_int(1000, 9999));
        }
        
        return $this->filename;
    }
    
    /**
     * @desc 存储路径， 需要有写入权限
     *
     * @param string $filepath
     * @return $this
     *
     * @author: ogenes<ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setFilepath(string $filepath): self
    {
        $this->filepath = $filepath;
        return $this;
    }
    
    public function getFilepath(): string
    {
        return $this->filepath;
    }
}