# Fix emoji encoding in Code.gs
$filePath = "Code.gs"
$content = [System.IO.File]::ReadAllText($filePath, [System.Text.Encoding]::UTF8)

# Replace corrupted emojis with proper Unicode escape sequences
$replacements = @{
    'â–¶ï¸' = "`u{25B6}`u{FE0F}"  # Play button
    'ðŸ"¥' = "`u{1F525}"  # Fire
    'ðŸ"' = "`u{1F50D}"  # Magnifying glass
    'ðŸ"§' = "`u{1F4E7}"  # Email
    'ðŸ"‹' = "`u{1F4CB}"  # Clipboard
    'ðŸ"Š' = "`u{1F4CA}"  # Bar chart
    'ðŸ"¦' = "`u{1F4E6}"  # Package
    'ðŸ"„' = "`u{1F504}"  # Counterclockwise arrows
    'â°' = "`u{23F0}"  # Alarm clock
    'ðŸ›'' = "`u{1F6D1}"  # Stop sign
    'ðŸŽ¯' = "`u{1F3AF}"  # Direct hit
    'ðŸ'¾' = "`u{1F4BE}"  # Floppy disk
    'ðŸ"ˆ' = "`u{1F4C8}"  # Chart increasing
    'ðŸ'°' = "`u{1F4B0}"  # Money bag
    'ðŸ"' = "`u{1F4C1}"  # File folder
    'ðŸ"…' = "`u{1F4C5}"  # Calendar
    'ðŸ"‚' = "`u{1F4C2}"  # Open file folder
    'ðŸ"¬' = "`u{1F52C}"  # Microscope
    'âš™ï¸' = "`u{2699}`u{FE0F}"  # Gear
    'ðŸ""' = "`u{1F513}"  # Unlocked padlock
    'ðŸ•'' = "`u{1F550}"  # Clock 1
    'ðŸ§¹' = "`u{1F9F9}"  # Broom
    'âš ï¸' = "`u{26A0}`u{FE0F}"  # Warning
    'â€¢' = "`u{2022}"  # Bullet
}

foreach ($key in $replacements.Keys) {
    $content = $content -replace [regex]::Escape($key), $replacements[$key]
}

# Write back with UTF-8 encoding (no BOM)
$utf8 = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($filePath, $content, $utf8)

Write-Host "Emojis fixed in Code.gs!"
