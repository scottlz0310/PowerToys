Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "PhotoGeo Preview Handler Test Host"
$form.Width = 800
$form.Height = 600
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = "Preview should appear below:"
$label.Dock = "Top"
$form.Controls.Add($label)

$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = "Fill"
$panel.BackColor = [System.Drawing.Color]::LightGray
$form.Controls.Add($panel)

$form.Add_Shown({
        $hwnd = $panel.Handle
        $rect = $panel.ClientRectangle
        $exePath = "C:\workspace\PowerToys\x64\Debug\PowerToys.PhotoGeoPreviewHandler.exe"

        # Create a dummy image file if it doesn't exist
        $imagePath = "C:\workspace\PowerToys\test_image.jpg"
        if (-not (Test-Path $imagePath)) {
            # Create a simple bitmap
            $bmp = New-Object System.Drawing.Bitmap(400, 300)
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            $g.Clear([System.Drawing.Color]::Blue)
            $g.DrawString("Test Image", [System.Drawing.Font]::new("Arial", 20), [System.Drawing.Brushes]::White, 10, 10)
            $bmp.Save($imagePath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
            $g.Dispose()
            $bmp.Dispose()
        }

        $arguments = @(
            "`"$imagePath`"",
            $hwnd.ToString("X"),
            $rect.Left,
            $rect.Right,
            $rect.Top,
            $rect.Bottom
        )

        Write-Host "Launching: $exePath"
        Write-Host "Args: $arguments"

        $process = Start-Process -FilePath $exePath -ArgumentList $arguments -PassThru

        $form.Add_FormClosing({
                # Kill by name to be sure, as the process object might be stale or insufficient if multiple launched
                Get-Process -Name "PowerToys.PhotoGeoPreviewHandler" -ErrorAction SilentlyContinue | Stop-Process -Force
            })
    })

[System.Windows.Forms.Application]::Run($form)
