<?php
namespace CSO\Excel2Html;

use CSO\Excel2Html\Exceptions\SheetNotFoudException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Font;

class HtmlConverter {
    //properties
    protected int $styleOption;
    protected float $scale;
    protected string $html;
    protected Worksheet $worksheet;
    /**
     * columns to read from Spreadsheet
     * @var string[]|null
     */
    protected array|null $columns;

    //ctor
    private function __construct(Worksheet $worksheet, int $styleOption = 0, array|null $columns = null, float $scale = 1.0){
        $this->styleOption = $styleOption;
        $this->columns = $columns;
        $this->scale = $scale;
        $this->worksheet = $worksheet;
        $this->html = '';
    }

    //statics
    public static function fromSpreadsheet(Spreadsheet $spreadsheet, int $styleOption = 0, string|null $worksheetName = null, array|null $columns = null, float $scale = 1.0): HtmlConverter{
        $worksheet = $spreadsheet->getSheetByName($worksheetName);
        if (is_null($worksheet)){
            throw new SheetNotFoudException();
        }
        return self::fromWorksheet($worksheet, $styleOption, $columns, $scale);
    }
    public static function fromWorksheet(Worksheet $worksheet, int $styleOption = 0, array|null $columns = null, float $scale = 1.0): HtmlConverter{
        $instance = new self($worksheet, $styleOption, $columns, $scale);
        return $instance;
    }
    public static function fromFilepath(string $filePath, int $styleOption = 0, string|null $worksheetName = null, array|null $columns = null, float $scale = 1.0): HtmlConverter{
        $spreadsheet = HtmlConverter::getSpreadsheetFromFilepath($filePath);
        return self::fromSpreadsheet($spreadsheet, $styleOption, $worksheetName, $columns, $scale);
    }

    //static helpers
    private static function getSpreadsheetFromFilepath(string $filePath): Spreadsheet{
        // init spreadsheet
        $inputFileType = IOFactory::identify($filePath);
        /**  Create a new Reader of the type that has been identified  **/
        $reader = IOFactory::createReader($inputFileType);
        /**  Load $inputFileName to a Spreadsheet Object  **/
        return $reader->load($filePath);
    }

    //methods
    public function getHtml(): string{
        if(!(is_null($this->html) || $this->html == '')){
            return $this->html;
        }
        // Get the highest row and column numbers referenced in the worksheet
        $tblWidth = $this->getTableWidth();
        //Get first data row
        $row = $this->getFirstDataRowIndex();
        $highestRow = $this->worksheet->getHighestDataRow();
        ob_start();
        ?>
        <table class="cso-excel-table" style="border-collapse: collapse; <?php echo $this->styleOption & StyleOptions::TABLE_SIZE_FIXED ? 'width:100%; table-layout: fixed; max-width: '.$tblWidth.'px;' : '' ?>">
            <colgroup>
                <?php foreach ($this->columns as $col): ?>
                    <?php
                        $colWidth = '';
                        if (($this->styleOption & StyleOptions::COLUMN_WIDTH_PROPORTIONAL) && ($this->styleOption & StyleOptions::WITH_COLUMN_WIDTH)){
                            $colWidth = 'width: '.($this->getColumnWidthInt($col) / $tblWidth * 100).'%;';
                        }
                        else if ($this->styleOption & StyleOptions::WITH_COLUMN_WIDTH){
                            $colWidth = 'width: '.$this->getColumnWidthInt($col).'px;';
                        }
                    ?>
                    <col style="<?php echo $colWidth ?>"></col>
                <?php endforeach; ?>
            </colgroup>
            <thead>
                <tr>
                    <?php foreach ($this->columns as $col): ?>
                        <th>
                            <?php echo $this->getCellValue($this->worksheet->getCell($col . $row)) ?>
                        </th>
                    <?php endforeach; ?>
                </tr>
            </thead>
            <tbody>
                <?php for ($row = $row + 1; $row <= $highestRow; ++$row): ?>
                    <tr excel-row="<?php echo $row ?>" style="height: <?php echo $this->getRowHeight($row) ?>px;">
                        <?php for ($i = 0; $i < count($this->columns);): ?>
                            <?php
                                $col = $this->columns[$i];
                                $cell = $this->worksheet->getCell($col . $row);
                                $colspan = $this->getColSpan($cell);
                                
                                $value = $this->getCellValue($cell);
                                
                                if(is_null($value))
                                    $value = '<br/>';
                                $tag = '<text>';
                                $closeTag = '</text>';
                                $font = $cell->getStyle()->getFont();
                                if ($font->getBold()){
                                    $tag .= '<b>';
                                    $closeTag = '</b>' . $closeTag;
                                }
                                else if ($font->getItalic()){
                                    $tag .= '<i>';
                                    $closeTag = '</i>' . $closeTag;
                                }
                                else if ($font->getSubscript()){
                                    $tag .= '<sub>';
                                    $closeTag = '</sub>' . $closeTag;
                                }
                                else if ($font->getSuperscript()){
                                    $tag .= '<sup>';
                                    $closeTag = '</sup>' . $closeTag;
                                }
                            ?>
                            <td colspan="<?php echo $colspan ?>" excel-col="<?php echo $cell->getColumn() ?>" excel-cell-range="<?php echo $cell->getMergeRange() ?>" style="background: <?php echo $this->getBackground($cell) ?>; <?php echo $this->getBorder($cell) ?>; white-space: <?php echo !$cell->getStyle()->getAlignment()->getWrapText() ? 'nowrap' : ''; ?>;">
                                <?php if ($cell->hasHyperlink()): ?>
                                    <a href="<?php echo $cell->getHyperlink()->getUrl() ?>" style="text-decoration: <?php echo $this->getTextDecoration($cell) ?>; color: #<?php echo $font->getColor()->getRGB()?>; font-size: <?php echo $this->getFontSizePt($cell) ?>pt; text-align: <?php echo $this->getAlignment($cell) ?>;">
                                        <text>
                                            <?php echo $tag ?>
                                                <?php echo $value?>
                                            <?php echo $closeTag ?>
                                        </text>
                                    </a>
                                <?php else: ?>
                                    <span style="display: block; width: 100%; color: #<?php echo $font->getColor()->getRGB()?>; font-size: <?php echo $this->getFontSizePt($cell) ?>pt; text-align: <?php echo $this->getAlignment($cell) ?>;">
                                        <text>
                                            <?php echo $tag ?>
                                                <?php echo $value?>
                                            <?php echo $closeTag ?>
                                        </text>
                                    </span>
                                <?php endif; ?>
                            </td>
                            <?php $i += $colspan; ?>
                        <?php endfor; ?>
                    </tr>
                <?php endfor; ?>
            </tbody>
        </table>
        <?php
        $this->html = ob_get_clean();
        
        return $this->html;
    }

    //private methods
    private function getColumnWidthInt(string $colName): float{
        $width = $this->worksheet->getColumnDimension($colName)->getWidth('px');
        if($width < 0){
            $width = $this->worksheet->getDefaultColumnDimension()->getWidth('px');
        }
        if($this->worksheet->getColumnDimension($colName)->getCollapsed()){
            $width = 0;
        }
        $colScale = 1;
        if($this->scale > 1.0){
            $colScale = $this->scale * 0.95;
        }
        return $width * $colScale;
    }
    private function getTableWidth(): float|int{
        $colWidth = 0;
        foreach ($this->columns as $col){
            $colWidth += $this->getColumnWidthInt($col);
        }
        return $colWidth;
    }
    private function getFirstDataRowIndex(): int{
        $highestRow = $this->worksheet->getHighestDataRow(); // e.g. 10
        $minCol = min($this->columns);
        $maxCol = max($this->columns);
        $row = -1;
        $value = null;
        do {
            ++$row;
            $array = $this->worksheet->rangeToArray($minCol.$row.':'.$maxCol.$row);
            $value = trim(implode($array));
        } while ($row <= $highestRow && is_null($value));
        ++$row;
        return $row;
    }
    private function getCellValue(Cell $cell): string{
        if($cell->isFormula()){
            return trim($cell->getOldCalculatedValue());
        }
        return trim($cell->getFormattedValue());
    }
    private function getRowHeight(int $rowNumber): float{
        $height = $this->worksheet->getRowDimension($rowNumber)->getRowHeight('px');
        if($height < 0){
            $height = $this->worksheet->getDefaultRowDimension()->getRowHeight('px');
            if($height < 0){
                $height = $this->worksheet->getRowDimension($rowNumber)->setRowHeight(10, 'pt')->getRowHeight('px');
            }
        }
        if($this->worksheet->getRowDimension($rowNumber)->getZeroHeight()){
            $height = 0;
        }
        return $height * $this->scale;
    }
    private function getColSpan(Cell $cell): int{
        if ($cell->isInMergeRange()){
            $mergeRange = explode(':', $cell->getMergeRange());
            $mergeRange[1] = substr($mergeRange[1], 0, strlen($mergeRange[1]) - strlen(filter_var($mergeRange[1], FILTER_SANITIZE_NUMBER_INT)));
            $mergeRange[0] = substr($mergeRange[0], 0, strlen($mergeRange[0]) - strlen(filter_var($mergeRange[0], FILTER_SANITIZE_NUMBER_INT)));
            
            $lastCell = array_search($mergeRange[1], $this->columns);
            $lastNotInCells = false;
            if ($lastCell === false){
                $lastNotInCells = true;
                $lastCell = count($this->columns);
                for($i = count($this->columns) - 1; $i >= 0; $i--){
                    if ((strlen($this->columns[$i]) === strlen($mergeRange[1]) && strcmp($this->columns[$i], $mergeRange[1]) <= 0) || (strlen($this->columns[$i]) < strlen($mergeRange[1]))){
                        $lastCell = $i;
                        if (strcmp($this->columns[$i], $mergeRange[1]) < 0)
                            $lastCell++;
                        break;
                    }
                    if( $i > count($this->columns))
                        $i = count($this->columns);
                }
            }
            $firstCell =  array_search($mergeRange[0], $this->columns);
            if ($firstCell === false){
                $firstCell = 0;
                for($i = 0; $i < count($this->columns); $i++){
                    if ((strlen($this->columns[$i]) === strlen($mergeRange[0]) && strcmp($this->columns[$i], $mergeRange[0]) >= 0) || (strlen($this->columns[$i]) > strlen($mergeRange[0]))){
                        $firstCell = $i;
                        if (strcmp($this->columns[$i], $mergeRange[0]) > 0)
                            $firstCell--;
                        break;
                    }
                }
                if( $i < 0)
                    $i = 0;
            }
            if($lastCell <= $firstCell){
                return 1;
            }
            
            if(!$lastNotInCells) 
                $lastCell+=1;
            return $lastCell - $firstCell;
        }
        return 1;
    }
    private function getBackground(Cell $cell): string{
        if ($cell->getStyle()->getFill()->getFillType() == null || $cell->getStyle()->getFill()->getFillType() === Fill::FILL_NONE){
            return 'transparent';
        }
        else if ($cell->getStyle()->getFill()->getFillType() === Fill::FILL_SOLID) {
            return '#'. $cell->getStyle()->getFill()->getStartColor()->getRGB();
        }
        return 'linear-gradient('
                . $cell->getStyle()->getFill()->getRotation()
                .'deg, #'
                . $cell->getStyle()->getFill()->getStartColor()->getRGB()
                .' 0%, #'
                . $cell->getStyle()->getFill()->getEndColor()->getRGB()
                .' 100%)';
    }

    private function getBorder(Cell $cell): string{
        $styles = [
            Border::BORDER_MEDIUM => 'solid 3px',
            Border::BORDER_NONE => 'none',
            Border::BORDER_THIN => 'solid 1px',
            Border::BORDER_THICK => 'solid 5px',
        ];
        if ($cell->isInMergeRange()){
            $borders = $cell->getStyle()->getBorders();
            $result = 'border-left: '. $styles[$borders->getLeft()->getBorderStyle()] .' #'. $borders->getLeft()->getColor()->getRGB() .'; /*'.$borders->getLeft()->getBorderStyle().'*/'//. $borders->getRight();
               .'border-bottom: '. $styles[$borders->getBottom()->getBorderStyle()] .' #'. $borders->getBottom()->getColor()->getRGB() .'; /*'.$borders->getBottom()->getBorderStyle().'*/'//. $borders->getRight();
               .'border-top: '. $styles[$borders->getTop()->getBorderStyle()] .' #'. $borders->getTop()->getColor()->getRGB() .'; /*'.$borders->getTop()->getBorderStyle().'*/';//. $borders->getRight();

            $mergeRange = explode(':', $cell->getMergeRange());
            $endcell = $this->worksheet->getCell($mergeRange[1]);
            $endBorders = $endcell->getStyle()->getBorders();
            $result .= 'border-right: '. $styles[$endBorders->getRight()->getBorderStyle()] .' #'. $endBorders->getRight()->getColor()->getRGB() .'; /*'.$borders->getRight()->getBorderStyle().'*/';//. $borders->getRight();
            return $result;
        }
        $borders = $cell->getStyle()->getBorders();
        return 'border-right: '. $styles[$borders->getRight()->getBorderStyle()] .' #'. $borders->getRight()->getColor()->getRGB() .'; /*'.$borders->getRight()->getBorderStyle().'*/'//. $borders->getRight();
               .'border-left: '. $styles[$borders->getLeft()->getBorderStyle()] .' #'. $borders->getLeft()->getColor()->getRGB() .'; /*'.$borders->getLeft()->getBorderStyle().'*/'//. $borders->getRight();
               .'border-bottom: '. $styles[$borders->getBottom()->getBorderStyle()] .' #'. $borders->getBottom()->getColor()->getRGB() .'; /*'.$borders->getBottom()->getBorderStyle().'*/'//. $borders->getRight();
               .'border-top: '. $styles[$borders->getTop()->getBorderStyle()] .' #'. $borders->getTop()->getColor()->getRGB() .'; /*'.$borders->getTop()->getBorderStyle().'*/';//. $borders->getRight();
    }
    private function getAlignment(Cell $cell): string{
        $alignment = $cell->getStyle()->getAlignment()->getHorizontal();
        if ($alignment === Alignment::HORIZONTAL_GENERAL){
            $alignment = Alignment::HORIZONTAL_LEFT;
        }
        return Alignment::HORIZONTAL_ALIGNMENT_FOR_HTML[$alignment];
    }
    private function getFontSizePt(Cell $cell): float{
        return $cell->getStyle()->getFont()->getSize() * $this->scale;
    }

    private function getTextDecoration(Cell $cell): string{
        if ($cell->getStyle()->getFont()->getUnderline() === Font::UNDERLINE_SINGLE){
            return 'underline solid';
        }
        else{
            return 'none';
        }
    }
}

class StyleOptions{
    const WITH_COLUMN_WIDTH         = 0b001;
    const COLUMN_WIDTH_PROPORTIONAL = 0b010;
    const TABLE_SIZE_FIXED          = 0b100;
}