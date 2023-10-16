







$repo = "rcalexterneuzen/rc-migration-street"
$filename = "rc-migration-street.zip"
$releases = "https://api.github.com/repos/$repo/releases"

$tag = (Invoke-WebRequest $releases | ConvertFrom-Json)[0].tag_name

$download = "https://github.com/$repo/releases/download/$tag/$file"
$name = $file.Split(".")[0]
$zip = "$name-$tag.zip"
$dir = "$name-$tag"

Invoke-WebRequest $download -Out $zip

Expand-Archive $zip -Force

Remove-Item $name -Recurse -Force -ErrorAction SilentlyContinue 

Move-Item $dir\$name -Destination $name -Force

Remove-Item $zip -Force
Remove-Item $dir -Recurse -Force

