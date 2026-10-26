param([string]$ProjectRoot = 'C:\DataMerge')
$ErrorActionPreference = 'Stop'
[void][Reflection.Assembly]::LoadFrom((Join-Path $ProjectRoot 'libs\Aspose.Cells.dll'))
$nativeLicense = New-Object Aspose.Cells.License
$nativeLicense.SetLicense((Join-Path $ProjectRoot 'libs\Aspose.Total.NET.lic'))
$sourceBook = New-Object Aspose.Cells.Workbook
$resultBook = New-Object Aspose.Cells.Workbook
try {
    $sourceBook.Settings.Date1904 = $true
    $source = $sourceBook.Worksheets[0]
    $source.Name = '原始明细'
    $source.Cells.Merge(0, 0, 1, 6)
    $source.Cells['A1'].PutValue('保留原始标题')
    $source.Cells['C7'].PutValue('工号')
    $source.Cells['F7'].PutValue('天数')
    $source.Cells['C8'].PutValue('001', $false)
    $source.Cells['F8'].PutValue(20.125)
    $style = $source.Cells['F8'].GetStyle()
    $style.Custom = '0.0000'
    $source.Cells['F8'].SetStyle($style)
    $source.Cells['E8'].PutValue([datetime]::new(2026, 10, 23))
    $dateStyle = $source.Cells['E8'].GetStyle()
    $dateStyle.Custom = 'yyyy-mm-dd'
    $source.Cells['E8'].SetStyle($dateStyle)
    $source.Cells.SetRowHeight(7, 31.0)
    $source.Cells.SetColumnWidth(4, 23.0)
    $source.Cells.HideColumn(3)
    $source.Cells.HideRow(11)
    $source.Cells['Z50'].PutValue('解析区域之外的尾注')
    $source.PageSetup.PrintArea = 'A1:Z50'
    $support = $sourceBook.Worksheets.Add('辅助')
    $support.Cells['A1'].PutValue(2)
    $support.Cells['A2'].PutValue(4)
    $support.IsVisible = $false
    $source.Cells['G8'].Formula = '=辅助!A1+1'
    $chartIndex = $source.Charts.Add([Aspose.Cells.Charts.ChartType]::Column, 12, 0, 22, 7)
    [void]$source.Charts[$chartIndex].NSeries.Add('辅助!A1:A2', $true)
    $dependencies = @($source.Cells['G8'].GetPrecedents())
    if ($dependencies.Count -ne 1 -or $dependencies[0].SheetName -ne '辅助') { throw '原表依赖提取失败' }

    $target = $resultBook.Worksheets[0]
    $target.Name = '目标'
    $target.Cells['A1'].PutValue('模板已有内容')
    [void]$resultBook.Worksheets.Add('源_原始明细')
    [void]$resultBook.Worksheets.Add('源_辅助')
    $target.Cells['D5'].Formula = "=INDEX('源_原始明细'!F:F,MATCH(`"001`",'源_原始明细'!C:C,0))"
    $target.Cells['D6'].Formula = "='源_原始明细'!G8"

    # Native copy contract used by template_pipeline: rename the in-memory
    # sources, precreate all destinations, and refer to same-named sheets.
    $source.Name = '源_原始明细'
    $support.Name = '源_辅助'
    $options = New-Object Aspose.Cells.CopyOptions
    $options.CopyNames = $true
    $options.ReferToSheetWithSameName = $true
    $options.ReferToDestinationSheet = $true
    $resultBook.Worksheets['源_原始明细'].Copy($source, $options)
    $resultBook.Worksheets['源_辅助'].Copy($support, $options)
    $resultBook.Worksheets['源_原始明细'].Cells['E8'].PutValue($source.Cells['E8'].DateTimeValue)
    $resultBook.CalculateFormula()
    $copied = $resultBook.Worksheets['源_原始明细']
    if ($copied.Cells['A1'].StringValue -ne '保留原始标题') { throw '标题未保留' }
    if ($copied.Cells.MergedCells.Count -ne 1) { throw '合并区域未保留' }
    if ($copied.Cells['C8'].Value -cne '001') { throw '文本主键类型未保留' }
    if ($copied.Cells['F8'].DoubleValue -ne 20.125) { throw '原始数值改变' }
    if ($copied.Cells['F8'].GetStyle().Custom -ne '0.0000') { throw '原始格式改变' }
    if ($copied.Cells.GetRowHeight(7) -ne 31.0) { throw '行高改变' }
    if ($copied.Cells.GetColumnWidth(4) -ne 23.0) { throw '列宽改变' }
    if (-not $copied.Cells.IsColumnHidden(3)) { throw '隐藏列改变' }
    if (-not $copied.Cells.IsRowHidden(11)) { throw '隐藏行改变' }
    if ($copied.Cells['Z50'].StringValue -ne '解析区域之外的尾注') { throw '只复制了解析区域，丢失整张表尾注' }
    if ($copied.PageSetup.PrintArea -notmatch 'Z\$?50') { throw '打印区域改变' }
    if ($copied.Charts.Count -ne 1 -or $copied.Charts[0].NSeries[0].Values -notmatch '源_辅助') { throw '图表数据源未保留' }
    if ($target.Cells['D5'].DoubleValue -ne 20.125) { throw ('真实行列坐标查找失败: ' + $target.Cells['D5'].StringValue + ' / ' + $target.Cells['D5'].Formula + ' key=' + $copied.Cells['C8'].Value + ' type=' + $copied.Cells['C8'].Value.GetType().FullName) }
    if ($target.Cells['D6'].DoubleValue -ne 3) { throw '跨源表依赖失效' }
    if ($target.Cells['A1'].StringValue -ne '模板已有内容') { throw '模板内容改变' }
    if ($resultBook.Worksheets.ExternalLinks.Count -ne 0) { throw ('复制产生了新的外部链接: ' + $copied.Cells['G8'].Formula + ' / ' + $resultBook.Worksheets.ExternalLinks[0].DataSource) }
    if ($copied.Cells['E8'].DateTimeValue.Date -ne [datetime]::new(2026, 10, 23)) { throw '1904/1900日期系统转换失败' }
    $roundTripPath = Join-Path ([IO.Path]::GetTempPath()) ('datamerge-native-' + [guid]::NewGuid().ToString('N') + '.xlsx')
    $reopenedBook = $null
    try {
        $resultBook.Save($roundTripPath)
        $reopenedBook = [Aspose.Cells.Workbook]::new($roundTripPath)
        $reopened = $reopenedBook.Worksheets['源_原始明细']
        if ($reopened.Cells['Z50'].StringValue -ne '解析区域之外的尾注') { throw '保存后整表内容丢失' }
        if ($reopened.Cells['C8'].Value -cne '001' -or $reopened.Cells['F8'].GetStyle().Custom -ne '0.0000') { throw '保存后坐标、类型或格式改变' }
        if ($reopened.Cells.MergedCells.Count -ne 1 -or $reopened.Charts.Count -ne 1) { throw '保存后合并单元格或图表丢失' }
    } finally {
        if ($null -ne $reopenedBook) { $reopenedBook.Dispose() }
        if (Test-Path -LiteralPath $roundTripPath) { Remove-Item -LiteralPath $roundTripPath }
    }
    Write-Output 'PASS: native original-sheet copy, true coordinates, formats, dependencies, formula calculation and date systems'
} finally {
    $resultBook.Dispose()
    $sourceBook.Dispose()
}
