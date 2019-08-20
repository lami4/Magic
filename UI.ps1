#Global variables
$script:SelectedRegister = $null
$script:PersonalRepository = "C:\Users\Tsedik\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB"
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
    $UpdateRegisterForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
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
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormBrowseFileButton)
    #Поле к кнопке Обзор для файла
    $UpdateRegisterFormBrowseButtonFileLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormBrowseButtonFileLabel.Location =  New-Object System.Drawing.Point(95,14) #x,y
    $UpdateRegisterFormBrowseButtonFileLabel.Width = 725
    $UpdateRegisterFormBrowseButtonFileLabel.Text = "Выберите файл, который необходимо обработать"
    $UpdateRegisterFormBrowseButtonFileLabel.TextAlign = "TopLeft"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormBrowseButtonFileLabel)
    #Надпись "Код товара"
    $UpdateRegisterFormItemCodeLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormItemCodeLabel.Location = New-Object System.Drawing.Point(10,46) #x,y
    $UpdateRegisterFormItemCodeLabel.Width = 75
    $UpdateRegisterFormItemCodeLabel.Text = "Код товара:"
    $UpdateRegisterFormItemCodeLabel.TextAlign = "TopLeft"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormItemCodeLabel)
    #Поле ввода столбца кода товара
    $UpdateRegisterFormInputItemCode = New-Object System.Windows.Forms.TextBox 
    $UpdateRegisterFormInputItemCode.Location = New-Object System.Drawing.Point(85,43) #-3x,y
    $UpdateRegisterFormInputItemCode.Width = 25
    $UpdateRegisterFormInputItemCode.Text = "I"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormInputItemCode)
    #Надпись "Черная цена"
    $UpdateRegisterFormBlackPriceLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormBlackPriceLabel.Location = New-Object System.Drawing.Point(10,74) #x,y
    $UpdateRegisterFormBlackPriceLabel.Width = 75
    $UpdateRegisterFormBlackPriceLabel.Text = "Черная цена:"
    $UpdateRegisterFormBlackPriceLabel.TextAlign = "TopLeft"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormBlackPriceLabel)
    #Поле ввода столбца черной цены
    $UpdateRegisterFormInputBlackPrice = New-Object System.Windows.Forms.TextBox 
    $UpdateRegisterFormInputBlackPrice.Location = New-Object System.Drawing.Point(85,71) #-3x,y
    $UpdateRegisterFormInputBlackPrice.Width = 25
    $UpdateRegisterFormInputBlackPrice.Text = "M"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormInputBlackPrice)
    #Надпись "Красная цена"
    $UpdateRegisterFormRedPriceLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormRedPriceLabel.Location = New-Object System.Drawing.Point(10,102) #x,y
    $UpdateRegisterFormRedPriceLabel.Width = 75
    $UpdateRegisterFormRedPriceLabel.Text = "Красн. цена:"
    $UpdateRegisterFormRedPriceLabel.TextAlign = "TopLeft"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormRedPriceLabel)
    #Поле ввода столбца красной цены
    $UpdateRegisterFormInputRedPrice = New-Object System.Windows.Forms.TextBox 
    $UpdateRegisterFormInputRedPrice.Location = New-Object System.Drawing.Point(85,99) #-3x,y
    $UpdateRegisterFormInputRedPrice.Width = 25
    $UpdateRegisterFormInputRedPrice.Text = "N"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormInputRedPrice)
    #Надпись "Скидка"
    $UpdateRegisterFormDiscountLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormDiscountLabel.Location = New-Object System.Drawing.Point(10,130) #x,y
    $UpdateRegisterFormDiscountLabel.Width = 75
    $UpdateRegisterFormDiscountLabel.Text = "Скидка:"
    $UpdateRegisterFormDiscountLabel.TextAlign = "TopLeft"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormDiscountLabel)
    #Поле ввода столбца скидка
    $UpdateRegisterFormInputDiscount = New-Object System.Windows.Forms.TextBox 
    $UpdateRegisterFormInputDiscount.Location = New-Object System.Drawing.Point(85,127) #-3x,y
    $UpdateRegisterFormInputDiscount.Width = 25
    $UpdateRegisterFormInputDiscount.Text = "O"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormInputDiscount)
    #Кнопка Начать
    $UpdateRegisterFormApplyButton = New-Object System.Windows.Forms.Button
    $UpdateRegisterFormApplyButton.Location = New-Object System.Drawing.Point(10,184) #x,y
    $UpdateRegisterFormApplyButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $UpdateRegisterFormApplyButton.Text = "Начать"
    $UpdateRegisterFormApplyButton.Enabled = $true
    $UpdateRegisterFormApplyButton.Add_Click({
        $TextInMessage = "Не указаны следующие параметры:`r`n"
        $ErrorPresent = $false
        if ($script:SelectedRegister -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан путь к файлу."}
        if ($UpdateRegisterFormInputItemCode.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец, в котором содержатся коды товаров."}
        if ($UpdateRegisterFormInputBlackPrice.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец, в котором содержится черная цена товаров."}
        if ($UpdateRegisterFormInputRedPrice.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец, в котором содержится красная цена товаров."}
        if ($UpdateRegisterFormInputDiscount.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец, в котором содержится скидка на товары."}
        if ($ErrorPresent -eq $true) {
            Show-MessageBox -Message $TextInMessage -Title "Невозможно начать" -Type OK
        } else {
            if ((Show-MessageBox -Message "Перед началом данной операции убедитесь в том, что у вас нет открытых Word и Excel документов.`r`nВо время работы скрипт закроет все Word и Excel документы, не сохраняя их, что может привести к потере данных!`r`nПродолжить?" -Title "Подтвердите действие" -Type YesNo) -eq "Yes") {
                $ExcelApp = New-Object -ComObject Excel.Application
                $ExcelApp.Visible = $true
                $Workbook = $ExcelApp.Workbooks.Open($script:PersonalRepository)
                $Workbook = $ExcelApp.Workbooks.Open($script:SelectedRegister)
                $Worksheet = $Workbook.Worksheets.Item(1)
                $ExcelApp.Run('PERSONAL.XLSB!ParseRegister', "$($UpdateRegisterFormInputItemCode.Text)", "$($UpdateRegisterFormInputBlackPrice.Text)", "$($UpdateRegisterFormInputRedPrice.Text)", "$($UpdateRegisterFormInputDiscount.Text)")
                $UpdateRegisterForm.Close()
            }
        }
    })
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormApplyButton)
    #Кнопка закрыть
    $UpdateRegisterFormCancelButton = New-Object System.Windows.Forms.Button
    $UpdateRegisterFormCancelButton.Location = New-Object System.Drawing.Point(100,184) #x,y
    $UpdateRegisterFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $UpdateRegisterFormCancelButton.Text = "Закрыть"
    $UpdateRegisterFormCancelButton.Add_Click({
        $UpdateRegisterForm.Close()
    })
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormCancelButton)
    $UpdateRegisterForm.ShowDialog()
}

MainForm | Out-Null
