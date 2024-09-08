Add-Type -AssemblyName System.Windows.Forms

$DocumentsPath = [environment]::GetFolderPath('MyDocuments')
$ModelsPath = Join-Path $DocumentsPath "Modelos Personalizados do Office"

$FileDialog = New-Object System.Windows.Forms.OpenFileDialog
$FileDialog.Title = "Selecione um modelo do Word"
$FileDialog.InitialDirectory = $ModelsPath

if ($FileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $modelSelect = $FileDialog.FileName


    $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveDialog.Filter = "Documentos do Word (*.docx)|*.docx"
    $SaveDialog.Title = "Salvar documento do Word"
    $SaveDialog.FileName = "NovoDocumento.docx"

    if ($SaveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $SavePath = $SaveDialog.FileName

        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $doc = $word.Documents.Add($modelSelect)

        $doc.SaveAs([ref]$SavePath)
    }
    else {
        Write-Host "Operação de salvar cancelada."
    }
}
else {
    Write-Host "Operação de seleção de modelo cancelada"
}