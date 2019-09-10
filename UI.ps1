#Global variables
$script:SelectedRegister = $null
$script:PersonalRepository = "C:\Users\selyuto\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB"
#Code
Function Open-File ($Filter, $MultipleSelectionFlag)
{
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = $Filter
    if ($MultipleSelectionFlag -eq $true) {$OpenFileDialog.Multiselect = $true}
    if ($MultipleSelectionFlag -eq $false) {$OpenFileDialog.Multiselect = $false}
    $DialogResult = $OpenFileDialog.ShowDialog()
    if ($DialogResult -eq "OK") {return $OpenFileDialog.FileNames} else {return $null}
}

Function Show-MessageBox ()
{ 
    param($Message, $Title, [ValidateSet("OK", "OKCancel", "YesNo")]$Type)
    Add-Type –AssemblyName System.Windows.Forms 
    if ($Type -eq "OK") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title")}  
    if ($Type -eq "OKCancel") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title",[System.Windows.Forms.MessageBoxButtons]::OKCancel)}
    if ($Type -eq "YesNo") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title",[System.Windows.Forms.MessageBoxButtons]::YesNo)}
}

Function MainForm ()
{
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    #Главное окно
    $UpdateRegisterForm = New-Object System.Windows.Forms.Form
    $UpdateRegisterForm.Padding = New-Object System.Windows.Forms.Padding(0,0,0,0)
    $UpdateRegisterForm.ShowIcon = $false
    $UpdateRegisterForm.AutoSize = $true
    $UpdateRegisterForm.Text = "Настройки"
    $UpdateRegisterForm.AutoSizeMode = "GrowAndShrink"
    $UpdateRegisterForm.WindowState = "Normal"
    $UpdateRegisterForm.SizeGripStyle = "Hide"
    $UpdateRegisterForm.ShowInTaskbar = $true
    $UpdateRegisterForm.StartPosition = "CenterScreen"
    $UpdateRegisterForm.MinimizeBox = $false
    $UpdateRegisterForm.MaximizeBox = $false
    #Tab elelement
    $ScriptMainWindowTabControl = New-object System.Windows.Forms.TabControl
    $ScriptMainWindowTabControl.Location = New-Object System.Drawing.Size(5,5)
    $ScriptMainWindowTabControl.Size = New-Object System.Drawing.Size(700,720) #width,height
    $UpdateRegisterForm.Controls.Add($ScriptMainWindowTabControl)
    #Table QA settings tab
    $Peter1Tab = New-Object System.Windows.Forms.TabPage
    $Peter1Tab.Text = "Питер_1”
    $ScriptMainWindowTabControl.Controls.Add($Peter1Tab)
    #Run table QA tab
    $Peter2Tab = New-Object System.Windows.Forms.TabPage
    $Peter2Tab.Text = "Питер_2”
    $ScriptMainWindowTabControl.Controls.Add($Peter2Tab)
    #TOOLTIP
    $ToolTip = New-Object System.Windows.Forms.ToolTip
    #Кнопка обзор для файла
    $UpdateRegisterFormBrowseFileButton = New-Object System.Windows.Forms.Button
    $UpdateRegisterFormBrowseFileButton.Location = New-Object System.Drawing.Point(10,10) #x,y
    $UpdateRegisterFormBrowseFileButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $UpdateRegisterFormBrowseFileButton.Text = "Обзор..."
    $UpdateRegisterFormBrowseFileButton.TabStop = $false
    $UpdateRegisterFormBrowseFileButton.Add_Click({
    $script:SelectedRegister = Open-File -Filter "All files (*.*)| *.*" -MultipleSelectionFlag $false
        if ($script:SelectedRegister -ne $null) {
            $UpdateRegisterFormBrowseButtonFileLabel.Text = "Указанный файл: $(Split-Path -Path $script:SelectedRegister -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($UpdateRegisterFormBrowseButtonFileLabel, $script:SelectedRegister)
        } else {
            $UpdateRegisterFormBrowseButtonFileLabel.Text = "Выберите файл, который необходимо обработать"
        }
    })
    $Peter1Tab.Controls.Add($UpdateRegisterFormBrowseFileButton)
    #Поле к кнопке Обзор для файла
    $UpdateRegisterFormBrowseButtonFileLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormBrowseButtonFileLabel.Location =  New-Object System.Drawing.Point(95,14) #x,y
    $UpdateRegisterFormBrowseButtonFileLabel.Width = 725
    $UpdateRegisterFormBrowseButtonFileLabel.Text = "Выберите файл, который необходимо обработать"
    $UpdateRegisterFormBrowseButtonFileLabel.TextAlign = "TopLeft"
    $Peter1Tab.Controls.Add($UpdateRegisterFormBrowseButtonFileLabel)
    #Удалить пустые строки по коду товара
    $DeleteEmptiStringsCheckbox = New-Object System.Windows.Forms.CheckBox
    $DeleteEmptiStringsCheckbox.Width = 350
    $DeleteEmptiStringsCheckbox.Text = "Удалить пустые строки по коду товара"
    $DeleteEmptiStringsCheckbox.Location = New-Object System.Drawing.Point(10,45) #x,y
    $DeleteEmptiStringsCheckbox.Enabled = $true
    $DeleteEmptiStringsCheckbox.Checked = $true
    $DeleteEmptiStringsCheckbox.Add_CheckStateChanged({
        if ($DeleteEmptiStringsCheckbox.Checked -eq $true) {$UpdateRegisterFormItemCodeLabel.Enabled = $true; $UpdateRegisterFormInputItemCode.Enabled = $true} else {$UpdateRegisterFormItemCodeLabel.Enabled = $false; $UpdateRegisterFormInputItemCode.Enabled = $false}
    })
    $Peter1Tab.Controls.Add($DeleteEmptiStringsCheckbox)
    #Надпись "Код товара"
    $UpdateRegisterFormItemCodeLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormItemCodeLabel.Location = New-Object System.Drawing.Point(10,74) #x,y
    $UpdateRegisterFormItemCodeLabel.Width = 75
    $UpdateRegisterFormItemCodeLabel.Text = "Код товара:"
    $UpdateRegisterFormItemCodeLabel.TextAlign = "TopLeft"
    $Peter1Tab.Controls.Add($UpdateRegisterFormItemCodeLabel)
    #Поле ввода столбца кода товара
    $UpdateRegisterFormInputItemCode = New-Object System.Windows.Forms.TextBox 
    $UpdateRegisterFormInputItemCode.Location = New-Object System.Drawing.Point(85,71) #-3x,y
    $UpdateRegisterFormInputItemCode.Width = 25
    $UpdateRegisterFormInputItemCode.Text = "I"
    $Peter1Tab.Controls.Add($UpdateRegisterFormInputItemCode)
    #Округлить скидку и сократить цены до двух дробных разрядов без округления
    $TruncatePricesAndRoundDiscountCheckbox = New-Object System.Windows.Forms.CheckBox
    $TruncatePricesAndRoundDiscountCheckbox.Width = 500
    $TruncatePricesAndRoundDiscountCheckbox.Text = "Округлить скидку и сократить цены до двух дробных разрядов без округления"
    $TruncatePricesAndRoundDiscountCheckbox.Location = New-Object System.Drawing.Point(10,105) #x,y
    $TruncatePricesAndRoundDiscountCheckbox.Enabled = $true
    $TruncatePricesAndRoundDiscountCheckbox.Checked = $false
    $TruncatePricesAndRoundDiscountCheckbox.Add_CheckStateChanged({
        if ($TruncatePricesAndRoundDiscountCheckbox.Checked -eq $true) {
                $UpdateRegisterFormBlackPriceLabel.Enabled = $true
                $UpdateRegisterFormInputBlackPrice.Enabled = $true
                $UpdateRegisterFormRedPriceLabel.Enabled = $true
                $UpdateRegisterFormInputRedPrice.Enabled = $true
                $UpdateRegisterFormDiscountLabel.Enabled = $true
                $UpdateRegisterFormInputDiscount.Enabled = $true
            } else {
                $UpdateRegisterFormBlackPriceLabel.Enabled = $false
                $UpdateRegisterFormInputBlackPrice.Enabled = $false
                $UpdateRegisterFormRedPriceLabel.Enabled = $false
                $UpdateRegisterFormInputRedPrice.Enabled = $false
                $UpdateRegisterFormDiscountLabel.Enabled = $false
                $UpdateRegisterFormInputDiscount.Enabled = $false
            }
    })
    $Peter1Tab.Controls.Add($TruncatePricesAndRoundDiscountCheckbox)
    #Надпись "Черная цена"
    $UpdateRegisterFormBlackPriceLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormBlackPriceLabel.Location = New-Object System.Drawing.Point(10,134) #x,y
    $UpdateRegisterFormBlackPriceLabel.Width = 75
    $UpdateRegisterFormBlackPriceLabel.Text = "Черная цена:"
    $UpdateRegisterFormBlackPriceLabel.TextAlign = "TopLeft"
    $Peter1Tab.Controls.Add($UpdateRegisterFormBlackPriceLabel)
    #Поле ввода столбца черной цены
    $UpdateRegisterFormInputBlackPrice = New-Object System.Windows.Forms.TextBox 
    $UpdateRegisterFormInputBlackPrice.Location = New-Object System.Drawing.Point(85,131) #-3x,y
    $UpdateRegisterFormInputBlackPrice.Width = 25
    $UpdateRegisterFormInputBlackPrice.Text = "M"
    $Peter1Tab.Controls.Add($UpdateRegisterFormInputBlackPrice)
    #Надпись "Красная цена"
    $UpdateRegisterFormRedPriceLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormRedPriceLabel.Location = New-Object System.Drawing.Point(10,163) #x,y
    $UpdateRegisterFormRedPriceLabel.Width = 75
    $UpdateRegisterFormRedPriceLabel.Text = "Красн. цена:"
    $UpdateRegisterFormRedPriceLabel.TextAlign = "TopLeft"
    $Peter1Tab.Controls.Add($UpdateRegisterFormRedPriceLabel)
    #Поле ввода столбца красной цены
    $UpdateRegisterFormInputRedPrice = New-Object System.Windows.Forms.TextBox 
    $UpdateRegisterFormInputRedPrice.Location = New-Object System.Drawing.Point(85,160) #-3x,y
    $UpdateRegisterFormInputRedPrice.Width = 25
    $UpdateRegisterFormInputRedPrice.Text = "N"
    $Peter1Tab.Controls.Add($UpdateRegisterFormInputRedPrice)
    #Надпись "Скидка"
    $UpdateRegisterFormDiscountLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormDiscountLabel.Location = New-Object System.Drawing.Point(10,192) #x,y
    $UpdateRegisterFormDiscountLabel.Width = 75
    $UpdateRegisterFormDiscountLabel.Text = "Скидка:"
    $UpdateRegisterFormDiscountLabel.TextAlign = "TopLeft"
    $Peter1Tab.Controls.Add($UpdateRegisterFormDiscountLabel)
    #Поле ввода столбца скидка
    $UpdateRegisterFormInputDiscount = New-Object System.Windows.Forms.TextBox 
    $UpdateRegisterFormInputDiscount.Location = New-Object System.Drawing.Point(85,189) #-3x,y
    $UpdateRegisterFormInputDiscount.Width = 25
    $UpdateRegisterFormInputDiscount.Text = "O"
    $Peter1Tab.Controls.Add($UpdateRegisterFormInputDiscount)
    #Удалить лишние столбцы
    $DeleteRedundantColumnsCheckbox = New-Object System.Windows.Forms.CheckBox
    $DeleteRedundantColumnsCheckbox.Width = 500
    $DeleteRedundantColumnsCheckbox.Text = "Удалить лишние столбцы"
    $DeleteRedundantColumnsCheckbox.Location = New-Object System.Drawing.Point(10,223) #x,y
    $DeleteRedundantColumnsCheckbox.Enabled = $true
    $DeleteRedundantColumnsCheckbox.Checked = $false
    $DeleteRedundantColumnsCheckbox.Add_CheckStateChanged({})
    $Peter1Tab.Controls.Add($DeleteRedundantColumnsCheckbox)
    #Удалить лишние столбцы
    $AddColumnBCheckbox = New-Object System.Windows.Forms.CheckBox
    $AddColumnBCheckbox.Width = 500
    $AddColumnBCheckbox.Text = "Добавить столбец /В/"
    $AddColumnBCheckbox.Location = New-Object System.Drawing.Point(10,248) #x,y
    $AddColumnBCheckbox.Enabled = $true
    $AddColumnBCheckbox.Checked = $true
    $AddColumnBCheckbox.Add_CheckStateChanged({})
    $Peter1Tab.Controls.Add($AddColumnBCheckbox)
    #Удалить лишние столбцы
    $DeleteAllButPeterCheckbox = New-Object System.Windows.Forms.CheckBox
    $DeleteAllButPeterCheckbox.Width = 500
    $DeleteAllButPeterCheckbox.Text = "Убрать все города, кроме Санкт-Петербурга"
    $DeleteAllButPeterCheckbox.Location = New-Object System.Drawing.Point(10,273) #x,y
    $DeleteAllButPeterCheckbox.Enabled = $true
    $DeleteAllButPeterCheckbox.Checked = $true
    $DeleteAllButPeterCheckbox.Add_CheckStateChanged({})
    $Peter1Tab.Controls.Add($DeleteAllButPeterCheckbox)
    #Кнопка Начать
    $UpdateRegisterFormApplyButton = New-Object System.Windows.Forms.Button
    $UpdateRegisterFormApplyButton.Location = New-Object System.Drawing.Point(10,660) #x,y
    $UpdateRegisterFormApplyButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $UpdateRegisterFormApplyButton.Text = "Начать"
    $UpdateRegisterFormApplyButton.Enabled = $true
    $UpdateRegisterFormApplyButton.Add_Click({
        $TextInMessage = "Не указаны следующие параметры:`r`n"
        $ErrorPresent = $false
        if ($script:SelectedRegister -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан путь к файлу."}
        if ($DeleteEmptiStringsCheckbox.Checked -eq $true -and $UpdateRegisterFormInputItemCode.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец, в котором содержатся коды товаров."}
        if ($TruncatePricesAndRoundDiscountCheckbox.Checked -eq $true -and $UpdateRegisterFormInputBlackPrice.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец, в котором содержится черная цена товаров."}
        if ($TruncatePricesAndRoundDiscountCheckbox.Checked -eq $true -and $UpdateRegisterFormInputRedPrice.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец, в котором содержится красная цена товаров."}
        if ($TruncatePricesAndRoundDiscountCheckbox.Checked -eq $true -and $UpdateRegisterFormInputDiscount.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец, в котором содержится скидка на товары."}
        if ($ErrorPresent -eq $true) {
            Show-MessageBox -Message $TextInMessage -Title "Невозможно начать" -Type OK
        } else {
            if ((Show-MessageBox -Message "Перед началом данной операции убедитесь в том, что у вас нет открытых Word и Excel документов.`r`nВо время работы скрипт закроет все Word и Excel документы, не сохраняя их, что может привести к потере данных!`r`nПродолжить?" -Title "Подтвердите действие" -Type YesNo) -eq "Yes") {
                $ExcelApp = New-Object -ComObject Excel.Application
                $ExcelApp.Visible = $true
                $Workbook = $ExcelApp.Workbooks.Open($script:PersonalRepository)
                $Workbook = $ExcelApp.Workbooks.Open($script:SelectedRegister)
                $Worksheet = $Workbook.Worksheets.Item(1)
                $ExcelApp.Run("$(Split-Path $script:PersonalRepository -Leaf)!ParseRegister", "$($UpdateRegisterFormInputItemCode.Text)", "$($UpdateRegisterFormInputBlackPrice.Text)", "$($UpdateRegisterFormInputRedPrice.Text)", "$($UpdateRegisterFormInputDiscount.Text)")
                $UpdateRegisterForm.Close()
            }
        }
    })
    $Peter1Tab.Controls.Add($UpdateRegisterFormApplyButton)
    #Кнопка закрыть
    $UpdateRegisterFormCancelButton = New-Object System.Windows.Forms.Button
    $UpdateRegisterFormCancelButton.Location = New-Object System.Drawing.Point(100,660) #x,y
    $UpdateRegisterFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $UpdateRegisterFormCancelButton.Text = "Закрыть"
    $UpdateRegisterFormCancelButton.Add_Click({
        $UpdateRegisterForm.Close()
    })
    $Peter1Tab.Controls.Add($UpdateRegisterFormCancelButton)
    if ($DeleteEmptiStringsCheckbox.Checked -eq $true) {$UpdateRegisterFormItemCodeLabel.Enabled = $true; $UpdateRegisterFormInputItemCode.Enabled = $true} else {$UpdateRegisterFormItemCodeLabel.Enabled = $false; $UpdateRegisterFormInputItemCode.Enabled = $false}
    if ($TruncatePricesAndRoundDiscountCheckbox.Checked -eq $true) {
        $UpdateRegisterFormBlackPriceLabel.Enabled = $true
        $UpdateRegisterFormInputBlackPrice.Enabled = $true
        $UpdateRegisterFormRedPriceLabel.Enabled = $true
        $UpdateRegisterFormInputRedPrice.Enabled = $true
        $UpdateRegisterFormDiscountLabel.Enabled = $true
        $UpdateRegisterFormInputDiscount.Enabled = $true
     } else {
        $UpdateRegisterFormBlackPriceLabel.Enabled = $false
        $UpdateRegisterFormInputBlackPrice.Enabled = $false
        $UpdateRegisterFormRedPriceLabel.Enabled = $false
        $UpdateRegisterFormInputRedPrice.Enabled = $false
        $UpdateRegisterFormDiscountLabel.Enabled = $false
        $UpdateRegisterFormInputDiscount.Enabled = $false
     }
    $UpdateRegisterForm.ShowDialog()
}

MainForm | Out-Null
