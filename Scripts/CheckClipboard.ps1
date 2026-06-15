# Check clipboard for zero-width spaces (U+200B)
$text = Get-Clipboard
if ($text) {
    Write-Host "Clipboard text:"
    Write-Host $text
    Write-Host ""
    Write-Host "Length: $($text.Length)"
    $zws = [char]0x200B
    $count = 0
    foreach ($c in $text.ToCharArray()) {
        if ($c -eq $zws) { $count++ }
    }
    Write-Host "Zero-width spaces: $count"
    if ($count -eq 0) {
        Write-Host "CLEAN"
    } else {
        Write-Host "DIRTY"
    }
} else {
    Write-Host "Clipboard is empty or not text"
}
