function Select-File 
{
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
function Update-ExcelDates
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline)]
        [Object[]]$Group
    )
    Process
    {
        foreach($singleGroup in $Group)
        {
            $singleGroup.Group | Format-Table Date, Amount, Description | Write-Output
            if($PSCmdlet.ShouldProcess("Update date in budget", "Are you sure you want to update the date in the budget to $($singleGroup[1].Date)","Update date"))
            {
                
            }
        }
    }
}


$month = Read-Host 'Enter worksheet name in Budget'
$budgetXLSX = Import-Excel -Path "$home\Documents\Finances\Budget\One Year Spending Plan - 2023.xlsx" -WorksheetName $month -StartRow 2 -StartColumn 2

$amex = Import-Csv -Path (Select-File)

$diff = $budgetXLSX | Where-Object {$_.Owner -eq 'Amex' -and $null -ne $_.Date} | Compare-Object -ReferenceObject $amex -Property {$_.Amount -as [double]}, {$_.Date -as [DateTime]} -PassThru

$diff | Group-Object {$_.Amount -as [Double]} | Where-Object Count -gt 1 | Foreach-Object { 
    $_.Group | Format-Table Date, Amount, Description | Write-Output
    if($PSCmdlet.ShouldProcess())
}

$diff | Format-table Date, Description, Amount, 'Card Member', SideIndicator