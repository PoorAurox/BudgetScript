function Select-File {
    param([string]$Directory = $PWD)
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = [System.Windows.Forms.OpenFileDialog]::new()
  
    $dialog.InitialDirectory = (Resolve-Path $Directory).Path
    $dialog.RestoreDirectory = $true
  
    $result = $dialog.ShowDialog()
  
    if($result -eq [System.Windows.Forms.DialogResult]::OK){
      return $dialog.FileName
    }
  }

$month = Read-Host 'Enter worksheet name in Budget'
$budgetXLSX = Import-Excel -Path "$home\Documents\Finances\Budget\One Year Spending Plan - 2023.xlsx" -WorksheetName $month -StartRow 2 -StartColumn 2

$amex = Import-Csv -Path (Select-File)

$diff = $budgetXLSX | Where {$_.Owner -eq 'Amex' -and $null -ne $_.Date} | Compare-Object -ReferenceObject $amex -Property {$_.Amount -as [double]} -PassThru