param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("VisioApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
