# Generate Ribbon button PNG icons using inline C#
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$resourceDir = Join-Path (Split-Path $scriptDir) "Resources"

if (-not (Test-Path $resourceDir)) {
    New-Item -ItemType Directory -Path $resourceDir -Force | Out-Null
}

$source = @"
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

public class IconGenerator
{
    public static void Generate(string path, int size)
    {
        using (var bmp = new Bitmap(size, size))
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);

            float penWidth = Math.Max(1f, size / 10f);
            float margin = size / 8f;

            // Chain links (dark blue)
            using (var pen = new Pen(Color.FromArgb(51, 102, 153), penWidth))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;

                float linkW = size * 0.35f;
                float linkH = size * 0.35f;
                float linkY = margin + size * 0.1f;

                // Left chain link
                g.DrawEllipse(pen, margin, linkY, linkW, linkH);
                // Right chain link
                g.DrawEllipse(pen, size * 0.5f, linkY, linkW, linkH);
            }

            // Break mark (red X)
            using (var redPen = new Pen(Color.FromArgb(200, 50, 50), penWidth * 1.5f))
            {
                redPen.StartCap = LineCap.Round;
                redPen.EndCap = LineCap.Round;

                float x1 = size * 0.3f, y1 = size * 0.6f;
                float x2 = size * 0.7f, y2 = size * 0.9f;
                g.DrawLine(redPen, x1, y1, x2, y2);
                g.DrawLine(redPen, x2, y1, x1, y2);
            }

            bmp.Save(path, ImageFormat.Png);
        }
    }
}
"@

Add-Type -TypeDefinition $source -ReferencedAssemblies System.Drawing

[IconGenerator]::Generate((Join-Path $resourceDir "RemoveLinks_16.png"), 16)
Write-Host "Generated: Resources\RemoveLinks_16.png" -ForegroundColor Green

[IconGenerator]::Generate((Join-Path $resourceDir "RemoveLinks_32.png"), 32)
Write-Host "Generated: Resources\RemoveLinks_32.png" -ForegroundColor Green

Write-Host ""
Write-Host "Done!" -ForegroundColor Cyan
