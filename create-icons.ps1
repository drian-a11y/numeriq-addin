# Create simple placeholder icons for the add-in
Add-Type -AssemblyName System.Drawing

function Create-Icon {
    param($size, $outputPath)
    
    $bitmap = New-Object System.Drawing.Bitmap($size, $size)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    
    # Fill with blue background
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 120, 212))
    $graphics.FillRectangle($brush, 0, 0, $size, $size)
    
    # Add white "N" text
    $font = New-Object System.Drawing.Font("Arial", ($size * 0.6), [System.Drawing.FontStyle]::Bold)
    $textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center
    $rect = New-Object System.Drawing.RectangleF(0, 0, $size, $size)
    $graphics.DrawString("N", $font, $textBrush, $rect, $format)
    
    # Save
    $bitmap.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    
    # Cleanup
    $graphics.Dispose()
    $bitmap.Dispose()
    $brush.Dispose()
    $textBrush.Dispose()
    $font.Dispose()
}

# Create icons
Write-Host "Creating icon files..." -ForegroundColor Green
Create-Icon 16 "assets\icon-16.png"
Create-Icon 32 "assets\icon-32.png"
Create-Icon 64 "assets\icon-64.png"
Create-Icon 80 "assets\icon-80.png"
Write-Host "Icons created successfully!" -ForegroundColor Green
