[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
$exitCode
function Set-Cancel
{
     $form.close()
     $global:exitCode1 = 1
}
function Set-RunScript
{
     $form.Dispose()
     $global:exitCode1 = 0
}
function Set-AllObjectProperties_OnResize
{
    $label.Width = $form.Width -10
}
function Set-DMLWarningForm
{
    $form= New-Object Windows.Forms.Form
    $form.Width = 500
    $form.add_ResizeEnd({Set-AllObjectProperties_OnResize})

    $label = New-Object Windows.Forms.Label
    $label.Width = $form.Width -10
    $label.Top = 30
    $label.Height = 100
    $label.Text = "Reminder:  Get the correct DML files for this release and put them in place before continuing."

    $button1 = New-Object Windows.Forms.Button
    $button1.Width = 150

    $button1.Top = 0
    $button1.text = "Cancel script"
    $button1.add_click({Set-Cancel})

    $button2 = New-Object Windows.Forms.Button
    $button2.Width = 150    
    $button2.Top = 0
    $button2.Left = $form.Width - 200
    $button2.text = "Continue script"
    $button2.add_click({Set-RunScript})
    
    $form.controls.add($label)
    $form.controls.add($button1)
    $form.controls.add($button2)
    $form.Modal = $true

    $form.ShowDialog()
    return $exitCode
}
#export-modulemember -function Set-DMLWarningForm