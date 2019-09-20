#Global variables
$script:SelectedRegister = $null
$script:Piter1RegisterFile = $null
$script:Piter2RegisterFile = $null
$script:PersonalRepository = ""
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
    #Файл Питер_1
    $Piter1GroupBox = New-Object System.Windows.Forms.GroupBox
    $Piter1GroupBox.Location = New-Object System.Drawing.Point(10,10) #x,y165
    $Piter1GroupBox.Size = New-Object System.Drawing.Point(670,120) #width,height
    $Piter1GroupBox.Text = "Файл Питер_1"
    $Peter2Tab.Controls.Add($Piter1GroupBox)
    #Файл Питер_2
    $Piter2GroupBox = New-Object System.Windows.Forms.GroupBox
    $Piter2GroupBox.Location = New-Object System.Drawing.Point(10,145) #x,y
    $Piter2GroupBox.Size = New-Object System.Drawing.Point(670,175) #width,height
    $Piter2GroupBox.Text = "Файл Питер_2"
    $Peter2Tab.Controls.Add($Piter2GroupBox)
    #Файл Доп. настройки
    $AdditionalSettingsGroupBox = New-Object System.Windows.Forms.GroupBox
    $AdditionalSettingsGroupBox.Location = New-Object System.Drawing.Point(10,335) #x,y
    $AdditionalSettingsGroupBox.Size = New-Object System.Drawing.Point(670,230) #width,height
    $AdditionalSettingsGroupBox.Text = "Доп. настройки"
    $Peter2Tab.Controls.Add($AdditionalSettingsGroupBox)

    #Удалить пустые строки по коду товара
    $DeleteEmptiStringsCheckboxPiter2 = New-Object System.Windows.Forms.CheckBox
    $DeleteEmptiStringsCheckboxPiter2.Width = 350
    $DeleteEmptiStringsCheckboxPiter2.Text = "Удалить пустые строки по коду товара в файле Питер_2"
    $DeleteEmptiStringsCheckboxPiter2.Location = New-Object System.Drawing.Point(10,20) #x,y 10,45
    $DeleteEmptiStringsCheckboxPiter2.Enabled = $true
    $DeleteEmptiStringsCheckboxPiter2.Checked = $false
    $DeleteEmptiStringsCheckboxPiter2.Add_CheckStateChanged({
        if ($DeleteEmptiStringsCheckboxPiter2.Checked -eq $true) {$UpdateRegisterFormItemCodeLabelPiter2.Enabled = $true; $UpdateRegisterFormInputItemCodePiter2.Enabled = $true} else {$UpdateRegisterFormItemCodeLabelPiter2.Enabled = $false; $UpdateRegisterFormInputItemCodePiter2.Enabled = $false}
    })
    $AdditionalSettingsGroupBox.Controls.Add($DeleteEmptiStringsCheckboxPiter2)
    #Надпись "Код товара"
    $UpdateRegisterFormItemCodeLabelPiter2 = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormItemCodeLabelPiter2.Location = New-Object System.Drawing.Point(10,49) #x,y 10,74
    $UpdateRegisterFormItemCodeLabelPiter2.Width = 75
    $UpdateRegisterFormItemCodeLabelPiter2.Text = "Код товара:"
    $UpdateRegisterFormItemCodeLabelPiter2.TextAlign = "TopLeft"
    $AdditionalSettingsGroupBox.Controls.Add($UpdateRegisterFormItemCodeLabelPiter2)
    #Поле ввода столбца кода товара
    $UpdateRegisterFormInputItemCodePiter2 = New-Object System.Windows.Forms.TextBox 
    $UpdateRegisterFormInputItemCodePiter2.Location = New-Object System.Drawing.Point(85,46) #-3x,y 85,71
    $UpdateRegisterFormInputItemCodePiter2.Width = 25
    $UpdateRegisterFormInputItemCodePiter2.Text = "I"
    $AdditionalSettingsGroupBox.Controls.Add($UpdateRegisterFormInputItemCodePiter2)
    #Округлить скидку и сократить цены до двух дробных разрядов без округления
    $TruncatePricesAndRoundDiscountCheckboxPiter2 = New-Object System.Windows.Forms.CheckBox
    $TruncatePricesAndRoundDiscountCheckboxPiter2.Width = 600
    $TruncatePricesAndRoundDiscountCheckboxPiter2.Text = "Округлить скидку и сократить цены до двух дробных разрядов без округления в файле Питер_2"
    $TruncatePricesAndRoundDiscountCheckboxPiter2.Location = New-Object System.Drawing.Point(10,80) #x,y 10,105
    $TruncatePricesAndRoundDiscountCheckboxPiter2.Enabled = $true
    $TruncatePricesAndRoundDiscountCheckboxPiter2.Checked = $false
    $TruncatePricesAndRoundDiscountCheckboxPiter2.Add_CheckStateChanged({
        if ($TruncatePricesAndRoundDiscountCheckboxPiter2.Checked -eq $true) {
                $UpdateRegisterFormBlackPriceLabelPiter2.Enabled = $true
                $UpdateRegisterFormInputBlackPricePiter2.Enabled = $true
                $UpdateRegisterFormRedPriceLabelPiter2.Enabled = $true
                $UpdateRegisterFormInputRedPricePiter2.Enabled = $true
                $UpdateRegisterFormDiscountLabelPiter2.Enabled = $true
                $UpdateRegisterFormInputDiscountPiter2.Enabled = $true
            } else {
                $UpdateRegisterFormBlackPriceLabelPiter2.Enabled = $false
                $UpdateRegisterFormInputBlackPricePiter2.Enabled = $false
                $UpdateRegisterFormRedPriceLabelPiter2.Enabled = $false
                $UpdateRegisterFormInputRedPricePiter2.Enabled = $false
                $UpdateRegisterFormDiscountLabelPiter2.Enabled = $false
                $UpdateRegisterFormInputDiscountPiter2.Enabled = $false
            }
    })
    $AdditionalSettingsGroupBox.Controls.Add($TruncatePricesAndRoundDiscountCheckboxPiter2)
    #Надпись "Черная цена"
    $UpdateRegisterFormBlackPriceLabelPiter2 = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormBlackPriceLabelPiter2.Location = New-Object System.Drawing.Point(10,109) #x,y 10,134
    $UpdateRegisterFormBlackPriceLabelPiter2.Width = 75
    $UpdateRegisterFormBlackPriceLabelPiter2.Text = "Черная цена:"
    $UpdateRegisterFormBlackPriceLabelPiter2.TextAlign = "TopLeft"
    $AdditionalSettingsGroupBox.Controls.Add($UpdateRegisterFormBlackPriceLabelPiter2)
    #Поле ввода столбца черной цены
    $UpdateRegisterFormInputBlackPricePiter2 = New-Object System.Windows.Forms.TextBox 
    $UpdateRegisterFormInputBlackPricePiter2.Location = New-Object System.Drawing.Point(85,106) #-3x,y 85,131
    $UpdateRegisterFormInputBlackPricePiter2.Width = 25
    $UpdateRegisterFormInputBlackPricePiter2.Text = "M"
    $AdditionalSettingsGroupBox.Controls.Add($UpdateRegisterFormInputBlackPricePiter2)
    #Надпись "Красная цена"
    $UpdateRegisterFormRedPriceLabelPiter2 = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormRedPriceLabelPiter2.Location = New-Object System.Drawing.Point(10,138) #x,y 10,163
    $UpdateRegisterFormRedPriceLabelPiter2.Width = 75
    $UpdateRegisterFormRedPriceLabelPiter2.Text = "Красн. цена:"
    $UpdateRegisterFormRedPriceLabelPiter2.TextAlign = "TopLeft"
    $AdditionalSettingsGroupBox.Controls.Add($UpdateRegisterFormRedPriceLabelPiter2)
    #Поле ввода столбца красной цены
    $UpdateRegisterFormInputRedPricePiter2 = New-Object System.Windows.Forms.TextBox 
    $UpdateRegisterFormInputRedPricePiter2.Location = New-Object System.Drawing.Point(85,135) #-3x,y 85,160
    $UpdateRegisterFormInputRedPricePiter2.Width = 25
    $UpdateRegisterFormInputRedPricePiter2.Text = "N"
    $AdditionalSettingsGroupBox.Controls.Add($UpdateRegisterFormInputRedPricePiter2)
    #Надпись "Скидка"
    $UpdateRegisterFormDiscountLabelPiter2 = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormDiscountLabelPiter2.Location = New-Object System.Drawing.Point(10,167) #x,y 10,192
    $UpdateRegisterFormDiscountLabelPiter2.Width = 75
    $UpdateRegisterFormDiscountLabelPiter2.Text = "Скидка:"
    $UpdateRegisterFormDiscountLabelPiter2.TextAlign = "TopLeft"
    $AdditionalSettingsGroupBox.Controls.Add($UpdateRegisterFormDiscountLabelPiter2)
    #Поле ввода столбца скидка
    $UpdateRegisterFormInputDiscountPiter2 = New-Object System.Windows.Forms.TextBox 
    $UpdateRegisterFormInputDiscountPiter2.Location = New-Object System.Drawing.Point(85,164) #-3x,y 85,189
    $UpdateRegisterFormInputDiscountPiter2.Width = 25
    $UpdateRegisterFormInputDiscountPiter2.Text = "O"
    $AdditionalSettingsGroupBox.Controls.Add($UpdateRegisterFormInputDiscountPiter2)
    #Удалить лишние столбцы
    $DeleteRedundantColumnsCheckboxPiter2 = New-Object System.Windows.Forms.CheckBox
    $DeleteRedundantColumnsCheckboxPiter2.Width = 500
    $DeleteRedundantColumnsCheckboxPiter2.Text = "Удалить лишние столбцы в файле Питер_2"
    $DeleteRedundantColumnsCheckboxPiter2.Location = New-Object System.Drawing.Point(10,198) #x,y 10,223
    $DeleteRedundantColumnsCheckboxPiter2.Enabled = $true
    $DeleteRedundantColumnsCheckboxPiter2.Checked = $false
    $DeleteRedundantColumnsCheckboxPiter2.Add_CheckStateChanged({})
    $AdditionalSettingsGroupBox.Controls.Add($DeleteRedundantColumnsCheckboxPiter2)
    
    #Кнопка обзор для файла Питер_1
    $Piter1FileBrowseFileButton = New-Object System.Windows.Forms.Button
    $Piter1FileBrowseFileButton.Location = New-Object System.Drawing.Point(10,20) #x,y
    $Piter1FileBrowseFileButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $Piter1FileBrowseFileButton.Text = "Обзор..."
    $Piter1FileBrowseFileButton.TabStop = $false
    $Piter1FileBrowseFileButton.Add_Click({
    $script:Piter1RegisterFile = Open-File -Filter "All files (*.*)| *.*" -MultipleSelectionFlag $false
        if ($script:Piter1RegisterFile -ne $null) {
            $Piter1FileBrowseButtonLabel.Text = "Указанный файл: $(Split-Path -Path $script:Piter1RegisterFile -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($Piter1FileBrowseButtonLabel, $script:Piter1RegisterFile)
        } else {
            $Piter1FileBrowseButtonLabel.Text = "Выберите файл Питер_1"
        }
    })
    $Piter1GroupBox.Controls.Add($Piter1FileBrowseFileButton)
    #Поле к кнопке Обзор для файла Питер_1
    $Piter1FileBrowseButtonLabel = New-Object System.Windows.Forms.Label
    $Piter1FileBrowseButtonLabel.Location =  New-Object System.Drawing.Point(95,24) #x,y
    $Piter1FileBrowseButtonLabel.Width = 725
    $Piter1FileBrowseButtonLabel.Text = "Выберите файл Питер_1"
    $Piter1FileBrowseButtonLabel.TextAlign = "TopLeft"
    $Piter1GroupBox.Controls.Add($Piter1FileBrowseButtonLabel)
    #Надпись "Код товара"
    $Piter1ItemCodeLabel = New-Object System.Windows.Forms.Label
    $Piter1ItemCodeLabel.Location = New-Object System.Drawing.Point(10,60) #x,y
    $Piter1ItemCodeLabel.Width = 75
    $Piter1ItemCodeLabel.Text = "Код товара:"
    $Piter1ItemCodeLabel.TextAlign = "TopLeft"
    $Piter1GroupBox.Controls.Add($Piter1ItemCodeLabel)
    #Поле ввода Код товара
    $Piter1ItemCodeInput = New-Object System.Windows.Forms.TextBox 
    $Piter1ItemCodeInput.Location = New-Object System.Drawing.Point(85,57) #-3x,y
    $Piter1ItemCodeInput.Width = 25
    $Piter1ItemCodeInput.Text = "I"
    $Piter1GroupBox.Controls.Add($Piter1ItemCodeInput)
    #Надпись "/B/"
    $Piter1ColumnBLabel = New-Object System.Windows.Forms.Label
    $Piter1ColumnBLabel.Location = New-Object System.Drawing.Point(10,89) #x,y
    $Piter1ColumnBLabel.Width = 75
    $Piter1ColumnBLabel.Text = "/B/:"
    $Piter1ColumnBLabel.TextAlign = "TopLeft"
    $Piter1GroupBox.Controls.Add($Piter1ColumnBLabel)
    #Поле ввода столбца "/B/"
    $Piter1ColumnBInput = New-Object System.Windows.Forms.TextBox 
    $Piter1ColumnBInput.Location = New-Object System.Drawing.Point(85,86) #-3x,y
    $Piter1ColumnBInput.Width = 25
    $Piter1ColumnBInput.Text = ""
    $Piter1GroupBox.Controls.Add($Piter1ColumnBInput)
    #Надпись "Верн. наименование"
    $Piter1CorrectNameLabel = New-Object System.Windows.Forms.Label
    $Piter1CorrectNameLabel.Location = New-Object System.Drawing.Point(10,118) #x,y
    $Piter1CorrectNameLabel.Width = 75
    $Piter1CorrectNameLabel.Text = "Верн. наим.:"
    $Piter1CorrectNameLabel.TextAlign = "TopLeft"
    #$Piter1GroupBox.Controls.Add($Piter1CorrectNameLabel)
    #Поле ввода столбца скидка
    $Piter1CorrectNameInput = New-Object System.Windows.Forms.TextBox 
    $Piter1CorrectNameInput.Location = New-Object System.Drawing.Point(85,115) #-3x,y
    $Piter1CorrectNameInput.Width = 25
    $Piter1CorrectNameInput.Text = ""
    #$Piter1GroupBox.Controls.Add($Piter1CorrectNameInput)

    #Кнопка обзор для файла Питер_2
    $Piter2FileBrowseFileButton = New-Object System.Windows.Forms.Button
    $Piter2FileBrowseFileButton.Location = New-Object System.Drawing.Point(10,20) #x,y
    $Piter2FileBrowseFileButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $Piter2FileBrowseFileButton.Text = "Обзор..."
    $Piter2FileBrowseFileButton.TabStop = $false
    $Piter2FileBrowseFileButton.Add_Click({
    $script:Piter2RegisterFile = Open-File -Filter "All files (*.*)| *.*" -MultipleSelectionFlag $false
        if ($script:Piter2RegisterFile -ne $null) {
            $Piter2FileBrowseButtonLabel.Text = "Указанный файл: $(Split-Path -Path $script:Piter2RegisterFile -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($Piter2FileBrowseButtonLabel, $script:Piter2RegisterFile)
        } else {
            $Piter2FileBrowseButtonLabel.Text = "Выберите файл Питер_2"
        }
    })
    $Piter2GroupBox.Controls.Add($Piter2FileBrowseFileButton)
    #Поле к кнопке Обзор для файла Питер_2
    $Piter2FileBrowseButtonLabel = New-Object System.Windows.Forms.Label
    $Piter2FileBrowseButtonLabel.Location =  New-Object System.Drawing.Point(95,24) #x,y
    $Piter2FileBrowseButtonLabel.Width = 725
    $Piter2FileBrowseButtonLabel.Text = "Выберите файл Питер_2"
    $Piter2FileBrowseButtonLabel.TextAlign = "TopLeft"
    $Piter2GroupBox.Controls.Add($Piter2FileBrowseButtonLabel)
    #Надпись "Код товара"
    $Piter2ItemCodeLabel = New-Object System.Windows.Forms.Label
    $Piter2ItemCodeLabel.Location = New-Object System.Drawing.Point(10,60) #x,y
    $Piter2ItemCodeLabel.Width = 75
    $Piter2ItemCodeLabel.Text = "Код товара:"
    $Piter2ItemCodeLabel.TextAlign = "TopLeft"
    $Piter2GroupBox.Controls.Add($Piter2ItemCodeLabel)
    #Поле ввода Код товара
    $Piter2ItemCodeInput = New-Object System.Windows.Forms.TextBox 
    $Piter2ItemCodeInput.Location = New-Object System.Drawing.Point(85,57) #-3x,y
    $Piter2ItemCodeInput.Width = 25
    $Piter2ItemCodeInput.Text = "I"
    $Piter2GroupBox.Controls.Add($Piter2ItemCodeInput)
    #Надпись "/B/"
    $Piter2ColumnBLabel= New-Object System.Windows.Forms.Label
    $Piter2ColumnBLabel.Location = New-Object System.Drawing.Point(10,89) #x,y
    $Piter2ColumnBLabel.Width = 75
    $Piter2ColumnBLabel.Text = "/B/:"
    $Piter2ColumnBLabel.TextAlign = "TopLeft"
    $Piter2GroupBox.Controls.Add($Piter2ColumnBLabel)
    #Поле ввода столбца "/B/"
    $Piter2ColumnBInput = New-Object System.Windows.Forms.TextBox 
    $Piter2ColumnBInput.Location = New-Object System.Drawing.Point(85,86) #-3x,y
    $Piter2ColumnBInput.Width = 25
    $Piter2ColumnBInput.Text = ""
    $Piter2GroupBox.Controls.Add($Piter2ColumnBInput)
    #Добавить колонку /Верное наименование/
    $Piter2CorrectNameCheckbox = New-Object System.Windows.Forms.CheckBox
    $Piter2CorrectNameCheckbox.Width = 500
    $Piter2CorrectNameCheckbox.Text = "Добавить столбец /Верное наименование/"
    $Piter2CorrectNameCheckbox.Location = New-Object System.Drawing.Point(10,118) #x,y
    $Piter2CorrectNameCheckbox.Enabled = $true
    $Piter2CorrectNameCheckbox.Checked = $true
    $Piter2CorrectNameCheckbox.Add_CheckStateChanged({})
    $Piter2GroupBox.Controls.Add($Piter2CorrectNameCheckbox)
    #Добавить колонку /Верное наименование/
    $Piter2RemoveCitiesCheckbox = New-Object System.Windows.Forms.CheckBox
    $Piter2RemoveCitiesCheckbox.Width = 500
    $Piter2RemoveCitiesCheckbox.Text = "Удалить все города, кроме Санкт-Петербурга"
    $Piter2RemoveCitiesCheckbox.Location = New-Object System.Drawing.Point(10,143) #x,y
    $Piter2RemoveCitiesCheckbox.Enabled = $true
    $Piter2RemoveCitiesCheckbox.Checked = $true
    $Piter2RemoveCitiesCheckbox.Add_CheckStateChanged({})
    $Piter2GroupBox.Controls.Add($Piter2RemoveCitiesCheckbox)
    #Надпись "Верн. наименование"
    $Piter2CorrectNameLabel = New-Object System.Windows.Forms.Label
    $Piter2CorrectNameLabel.Location = New-Object System.Drawing.Point(10,118) #x,y
    $Piter2CorrectNameLabel.Width = 75
    $Piter2CorrectNameLabel.Text = "Верн. наим.:"
    $Piter2CorrectNameLabel.TextAlign = "TopLeft"
    #$Piter2GroupBox.Controls.Add($Piter2CorrectNameLabel)
    #Поле ввода столбца "Верн. наименование"
    $Piter2CorrectNameInput = New-Object System.Windows.Forms.TextBox 
    $Piter2CorrectNameInput.Location = New-Object System.Drawing.Point(85,115) #-3x,y
    $Piter2CorrectNameInput.Width = 25
    $Piter2CorrectNameInput.Text = ""
    #$Piter2GroupBox.Controls.Add($Piter2CorrectNameInput)
    #Кнопка Начать
    $Piter2TabApplyButton = New-Object System.Windows.Forms.Button
    $Piter2TabApplyButton.Location = New-Object System.Drawing.Point(10,660) #x,y
    $Piter2TabApplyButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $Piter2TabApplyButton.Text = "Начать"
    $Piter2TabApplyButton.Enabled = $true
    $Piter2TabApplyButton.Add_Click({
        $TextInMessage = "Не указаны следующие параметры:`r`n"
        $ErrorPresent = $false
        if ($script:Piter1RegisterFile -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан путь к файлу Питер_1."}
        if ($script:Piter2RegisterFile -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан путь к файлу Питер_2."}
        if ($Piter1ItemCodeInput.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец 'Код товара' для файла Питер_1"}
        if ($Piter1ColumnBInput.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец '/B/' для файла Питер_1"}
        #if ($Piter1CorrectNameInput.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец 'Правильное наименование' для файла Питер_1"}
        if ($Piter2ItemCodeInput.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец 'Код товара' для файла Питер_2"}
        if ($Piter2ColumnBInput.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец '/B/' для файла Питер_2"}
        #if ($Piter2CorrectNameInput.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец 'Правильное наименование' для файла Питер_2"}
        if ($DeleteEmptiStringsCheckboxPiter2.Checked -eq $true -and $UpdateRegisterFormInputItemCodePiter2.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец, в котором содержатся коды товаров."}
        if ($TruncatePricesAndRoundDiscountCheckboxPiter2.Checked -eq $true -and $UpdateRegisterFormInputBlackPricePiter2.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец, в котором содержится черная цена товаров."}
        if ($TruncatePricesAndRoundDiscountCheckboxPiter2.Checked -eq $true -and $UpdateRegisterFormInputRedPricePiter2.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец, в котором содержится красная цена товаров."}
        if ($TruncatePricesAndRoundDiscountCheckboxPiter2.Checked -eq $true -and $UpdateRegisterFormInputDiscountPiter2.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан столбец, в котором содержится скидка на товары."}
        if ($ErrorPresent -eq $true) {
            Show-MessageBox -Message $TextInMessage -Title "Невозможно начать" -Type OK
        } else {
            if ((Show-MessageBox -Message "Перед началом данной операции убедитесь в том, что у вас нет открытых Word и Excel документов.`r`nВо время работы скрипт закроет все Word и Excel документы, не сохраняя их, что может привести к потере данных!`r`nПродолжить?" -Title "Подтвердите действие" -Type YesNo) -eq "Yes") {
                $ExcelAppPiterTwo = New-Object -ComObject Excel.Application
                $ExcelAppPiterTwo.Visible = $true
                $Workbook = $ExcelAppPiterTwo.Workbooks.Open($script:PersonalRepository)
                if ($DeleteEmptiStringsCheckboxPiter2.Checked -eq $true) {$PiterTwoDeleteEmptyRowsFlag = "true"} else {$PiterTwoDeleteEmptyRowsFlag = "false"}
                if ($TruncatePricesAndRoundDiscountCheckboxPiter2.Checked -eq $true) {$NormalizePricesAndDiscountsFlag = "true"} else {$NormalizePricesAndDiscountsFlag = "false"}
                if ($DeleteRedundantColumnsCheckboxPiter2.Checked -eq $true) {$RemoveRedundantColumnsFlag = "true"} else {$RemoveRedundantColumnsFlag = "false"}
                if ($Piter2CorrectNameCheckbox.Checked -eq $true) {$PiterTwoAddBColumnFlag = "true"} else {$PiterTwoAddBColumnFlag = "false"}
                if ($Piter2RemoveCitiesCheckbox.Checked -eq $true) {$PiterTwoDeleteCitiesFlag = "true"} else {$PiterTwoDeleteCitiesFlag = "false"}
                $ExcelAppPiterTwo.Run("$(Split-Path $script:PersonalRepository -Leaf)!CopyDataToAnotherFile", "$script:Piter1RegisterFile", "$($Piter1ItemCodeInput.Text)", "$($Piter1ColumnBInput.Text)", "$script:Piter2RegisterFile", "$($Piter2ItemCodeInput.Text)", "$($Piter2ColumnBInput.Text)", "$PiterTwoAddBColumnFlag", "$PiterTwoDeleteCitiesFlag", "$PiterTwoDeleteEmptyRowsFlag", "$($UpdateRegisterFormInputItemCodePiter2.Text)", "$NormalizePricesAndDiscountsFlag", "$($UpdateRegisterFormInputBlackPricePiter2.Text)", "$($UpdateRegisterFormInputRedPricePiter2.Text)", "$($UpdateRegisterFormInputDiscountPiter2.Text)", "$RemoveRedundantColumnsFlag")
                #$UpdateRegisterForm.Close()
            }
        }
    })
    $Peter2Tab.Controls.Add($Piter2TabApplyButton)
    #Кнопка закрыть
    $Piter2TabCancelButton = New-Object System.Windows.Forms.Button
    $Piter2TabCancelButton.Location = New-Object System.Drawing.Point(100,660) #x,y
    $Piter2TabCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $Piter2TabCancelButton.Text = "Закрыть"
    $Piter2TabCancelButton.Add_Click({
        $UpdateRegisterForm.Close()
    })
    $Peter2Tab.Controls.Add($Piter2TabCancelButton)
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
    #Добавить колонку /B/
    $AddColumnBCheckbox = New-Object System.Windows.Forms.CheckBox
    $AddColumnBCheckbox.Width = 500
    $AddColumnBCheckbox.Text = "Добавить столбец /В/"
    $AddColumnBCheckbox.Location = New-Object System.Drawing.Point(10,248) #x,y
    $AddColumnBCheckbox.Enabled = $true
    $AddColumnBCheckbox.Checked = $true
    $AddColumnBCheckbox.Add_CheckStateChanged({})
    $Peter1Tab.Controls.Add($AddColumnBCheckbox)
    #Удалить все города, кроме Санкт-Петербурга
    $DeleteAllButPeterCheckbox = New-Object System.Windows.Forms.CheckBox
    $DeleteAllButPeterCheckbox.Width = 500
    $DeleteAllButPeterCheckbox.Text = "Удалить все города, кроме Санкт-Петербурга"
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
                if ($DeleteEmptiStringsCheckbox.Checked -eq $true) {$DeleteEmptyRowsFlag = "true"} else {$DeleteEmptyRowsFlag = "false"}
                if ($TruncatePricesAndRoundDiscountCheckbox.Checked -eq $true) {$NormalizePricesAndDiscountsFlag = "true"} else {$NormalizePricesAndDiscountsFlag = "false"}
                if ($DeleteRedundantColumnsCheckbox.Checked -eq $true) {$RemoveRedundantColumnsFlag = "true"} else {$RemoveRedundantColumnsFlag = "false"}
                if ($AddColumnBCheckbox.Checked -eq $true) {$AddBColumnFlag = "true"} else {$AddBColumnFlag = "false"}
                if ($DeleteAllButPeterCheckbox.Checked -eq $true) {$DeleteCitiesFlag = "true"} else {$DeleteCitiesFlag = "false"}
                $ExcelApp.Run("$(Split-Path $script:PersonalRepository -Leaf)!ParseRegister", "$($UpdateRegisterFormInputItemCode.Text)", "$($UpdateRegisterFormInputBlackPrice.Text)", "$($UpdateRegisterFormInputRedPrice.Text)", "$($UpdateRegisterFormInputDiscount.Text)", $DeleteEmptyRowsFlag, $NormalizePricesAndDiscountsFlag, $RemoveRedundantColumnsFlag, $AddBColumnFlag, $DeleteCitiesFlag)
                #$UpdateRegisterForm.Close()
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
    if ($DeleteEmptiStringsCheckboxPiter2.Checked -eq $true) {$UpdateRegisterFormItemCodeLabelPiter2.Enabled = $true; $UpdateRegisterFormInputItemCodePiter2.Enabled = $true} else {$UpdateRegisterFormItemCodeLabelPiter2.Enabled = $false; $UpdateRegisterFormInputItemCodePiter2.Enabled = $false}
    if ($TruncatePricesAndRoundDiscountCheckboxPiter2.Checked -eq $true) {
        $UpdateRegisterFormBlackPriceLabelPiter2.Enabled = $true
        $UpdateRegisterFormInputBlackPricePiter2.Enabled = $true
        $UpdateRegisterFormRedPriceLabelPiter2.Enabled = $true
        $UpdateRegisterFormInputRedPricePiter2.Enabled = $true
        $UpdateRegisterFormDiscountLabelPiter2.Enabled = $true
        $UpdateRegisterFormInputDiscountPiter2.Enabled = $true
     } else {
        $UpdateRegisterFormBlackPriceLabelPiter2.Enabled = $false
        $UpdateRegisterFormInputBlackPricePiter2.Enabled = $false
        $UpdateRegisterFormRedPriceLabelPiter2.Enabled = $false
        $UpdateRegisterFormInputRedPricePiter2.Enabled = $false
        $UpdateRegisterFormDiscountLabelPiter2.Enabled = $false
        $UpdateRegisterFormInputDiscountPiter2.Enabled = $false
     }
    $UpdateRegisterForm.ShowDialog()
}

MainForm | Out-Null
