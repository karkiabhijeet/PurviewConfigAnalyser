
function Show-PurviewConfigAnalyserGUI {
    [CmdletBinding()]
    param(
        [string]$MasterControlPath = (Join-Path $PSScriptRoot "..\..\config\MasterControlBooks\ControlBook_Reference.csv"),
        [string]$MasterPropertyPath = (Join-Path $PSScriptRoot "..\..\config\MasterControlBooks\ControlBook_Property_Reference.csv"),
        [string]$OutputPath = (Join-Path $PSScriptRoot "..\..\config")
    )
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
    } catch {
        Write-Host "❌ Could not load Windows Forms assemblies: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
    try {
        $controls = Import-Csv $MasterControlPath
        $properties = Import-Csv $MasterPropertyPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading control book CSVs: $($_.Exception.Message)", "Error", 'OK', 'Error')
        return
    }
    $form = New-Object Windows.Forms.Form
    $form.Text = "Custom Configuration Creator"
    $form.Size = New-Object Drawing.Size(900, 650)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $label = New-Object Windows.Forms.Label
    $label.Text = "Select controls and configure properties."
    $label.AutoSize = $true
    $label.Location = New-Object Drawing.Point(10,10)
    $form.Controls.Add($label)
    $tree = New-Object Windows.Forms.TreeView
    $tree.Location = New-Object Drawing.Point(10,40)
    $tree.Size = New-Object Drawing.Size(350,500)
    $tree.CheckBoxes = $true
    $form.Controls.Add($tree)
    $capabilities = $controls | Group-Object Capability
    foreach ($cap in $capabilities) {
        $parent = $tree.Nodes.Add($cap.Name)
        foreach ($ctrl in $cap.Group) {
            $node = $parent.Nodes.Add("$($ctrl.ControlID) - $($ctrl.Control)")
            $node.Tag = $ctrl.ControlID
        }
    }
    # Panel for dynamic property controls
    $propertyPanel = New-Object Windows.Forms.Panel
    $propertyPanel.Location = New-Object Drawing.Point(370,40)
    $propertyPanel.Size = New-Object Drawing.Size(500,500)
    $propertyPanel.AutoScroll = $true
    $form.Controls.Add($propertyPanel)
    $propertyInputs = @{}

    function Update-PropertyPanel {
        $propertyPanel.Controls.Clear()
        $propertyInputs.Clear()
        $y = 10
        $checkedControls = @()
        foreach ($cap in $tree.Nodes) {
            foreach ($ctrl in $cap.Nodes) {
                if ($ctrl.Checked) { $checkedControls += $ctrl.Tag }
            }
        }
        foreach ($cid in $checkedControls) {
            $props = $properties | Where-Object { $_.ControlID -eq $cid }
            if ($props) {
                $labelHeader = New-Object Windows.Forms.Label
                $labelHeader.Text = "ControlID: $cid"
                $labelHeader.Font = New-Object Drawing.Font('Segoe UI',10,[Drawing.FontStyle]::Bold)
                $labelHeader.Location = New-Object Drawing.Point(10, $y)
                $labelHeader.Size = New-Object Drawing.Size(400, 25)
                $propertyPanel.Controls.Add($labelHeader)
                $y += 30
                foreach ($p in $props) {
                    $isRequired = ($p.MustConfigure -eq $true -or $p.MustConfigure -eq 'true' -or [string]::IsNullOrWhiteSpace($p.DefaultValue))
                    $labelText = if ($isRequired) { "* $($p.Properties)" } else { $p.Properties }
                    $label = New-Object Windows.Forms.Label
                    $label.Text = $labelText
                    $label.Location = New-Object Drawing.Point(20, $y)
                    $label.Size = New-Object Drawing.Size(200, 25)
                    $propertyPanel.Controls.Add($label)
                    $tb = New-Object Windows.Forms.TextBox
                    $tb.Location = New-Object Drawing.Point(230, $y)
                    $tb.Size = New-Object Drawing.Size(250, 25)
                    $tb.Text = $p.DefaultValue
                    $propertyPanel.Controls.Add($tb)
                    $propertyInputs["$($cid)|$($p.Properties)"] = @{ TextBox = $tb; Required = $isRequired }
                    $y += 35
                }
            }
        }
        if ($y -eq 10) {
            $labelNone = New-Object Windows.Forms.Label
            $labelNone.Text = 'Check controls to configure their properties.'
            $labelNone.Location = New-Object Drawing.Point(10, $y)
            $labelNone.Size = New-Object Drawing.Size(400, 25)
            $propertyPanel.Controls.Add($labelNone)
        } else {
            $noteLabel = New-Object Windows.Forms.Label
            $noteLabel.Text = '* Required property'
            $noteLabel.ForeColor = [System.Drawing.Color]::Red
            $noteLabel.Location = New-Object Drawing.Point(10, $y)
            $noteLabel.Size = New-Object Drawing.Size(400, 20)
            $propertyPanel.Controls.Add($noteLabel)
        }
    }

    # Update property panel when a control is checked/unchecked
    $tree.Add_AfterCheck({ Update-PropertyPanel })
    # Also update when the form loads
    Update-PropertyPanel
    $saveBtn = New-Object Windows.Forms.Button
    $saveBtn.Text = "Save Configuration"
    $saveBtn.Location = New-Object Drawing.Point(370,560)
    $saveBtn.Size = New-Object Drawing.Size(200,40)
    $saveBtn.Add_Click({
        $checked = @()
        foreach ($cap in $tree.Nodes) {
            foreach ($ctrl in $cap.Nodes) {
                if ($ctrl.Checked) { $checked += $ctrl.Tag }
            }
        }
        if (-not $checked) {
            [Windows.Forms.MessageBox]::Show("Select at least one control.","Validation",'OK','Warning')
            return
        }
        $ctrlOut = @()
        $propOut = @()
        $allValid = $true
        $missingProps = @()
        foreach ($cid in $checked) {
            $ctrl = $controls | Where-Object { $_.ControlID -eq $cid }
            $ctrlOut += [PSCustomObject]@{ Capability=$ctrl.Capability; ControlID=$ctrl.ControlID; Control=$ctrl.Control }
            $props = $properties | Where-Object { $_.ControlID -eq $cid }
            foreach ($p in $props) {
                $key = "$($cid)|$($p.Properties)"
                if ($propertyInputs.ContainsKey($key)) {
                    $val = $propertyInputs[$key].TextBox.Text
                    $isRequired = $propertyInputs[$key].Required
                } else {
                    $val = $p.DefaultValue
                    $isRequired = ($p.MustConfigure -eq $true -or $p.MustConfigure -eq 'true' -or [string]::IsNullOrWhiteSpace($p.DefaultValue))
                }
                if ($isRequired -and [string]::IsNullOrWhiteSpace($val)) {
                    $allValid = $false
                    $missingProps += "$($p.Properties) (ControlID: $cid)"
                }
                $propOut += [PSCustomObject]@{ ControlID=$cid; Properties=$p.Properties; DefaultValue=$val; MustConfigure=$p.MustConfigure }
            }
        }
        if (-not $allValid) {
            [Windows.Forms.MessageBox]::Show("Please fill all required properties before saving:`n`n" + ($missingProps -join "`n"),"Validation",'OK','Warning')
            return
        }
        $configName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter configuration name:","Config Name","MyCustomConfig")
        if (-not $configName) { return }
        $ctrlOut | Export-Csv (Join-Path $OutputPath "ControlBook_${configName}_Config.csv") -NoTypeInformation
        $propOut | Export-Csv (Join-Path $OutputPath "ControlBook_Property_${configName}_Config.csv") -NoTypeInformation
        [Windows.Forms.MessageBox]::Show("Configuration saved!","Success",'OK','Information')
        $form.Close()
    })
    $form.Controls.Add($saveBtn)
    $cancelBtn = New-Object Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Location = New-Object Drawing.Point(580,560)
    $cancelBtn.Size = New-Object Drawing.Size(120,40)
    $cancelBtn.Add_Click({ $form.Close() })
    $form.Controls.Add($cancelBtn)
    $form.ShowDialog()
}
