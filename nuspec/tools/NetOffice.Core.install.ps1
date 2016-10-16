param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("NetOffice")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}

$ref = $project.Object.References.Item("OfficeApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}

$ref = $project.Object.References.Item("VBIDEApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
