param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("PowerPointApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
