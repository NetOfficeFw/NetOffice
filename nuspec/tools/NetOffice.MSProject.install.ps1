param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("MSProjectApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}

$ref = $project.Object.References.Item("MSHTMLApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
