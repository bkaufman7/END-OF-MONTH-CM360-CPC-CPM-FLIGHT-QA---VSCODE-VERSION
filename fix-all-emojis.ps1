# Fix all emojis in Code.gs
$filePath = "Code.gs"
$content = [System.IO.File]::ReadAllText($filePath, [System.Text.Encoding]::UTF8)

# Replace all corrupted emoji patterns with proper emojis
$content = $content -replace '`u\{26A0\}`u\{FE0F\}', '⚠️'
$content = $content -replace '`u\{2022\}', '•'
$content = $content -replace 'ðŸŸ¨', '🟨'
$content = $content -replace 'â‰¥', '≥'
$content = $content -replace 'â€¦', '…'
$content = $content -replace 'â€"', '—'
$content = $content -replace 'âŒ', '❌'
$content = $content -replace 'Ã—', '×'
$content = $content -replace 'âœ•', '✕'
$content = $content -replace 'â€™', "'"
$content = $content -replace 'â€œ', '"'
$content = $content -replace 'â€', '"'
$content = $content -replace 'ðŸŸ¥', '🟥'
$content = $content -replace 'ðŸš¨', '🚨'
$content = $content -replace 'âš ï¸', '⚠️'
$content = $content -replace 'ðŸŸ¦', '🟦'
$content = $content -replace 'ðŸŸ©', '🟩'
$content = $content -replace 'âœ…', '✅'
$content = $content -replace 'â³', '⏳'

# Write back with UTF-8 encoding (no BOM)
$utf8 = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($filePath, $content, $utf8)

Write-Host "All emojis fixed!"
