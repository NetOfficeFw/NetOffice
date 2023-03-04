param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("AccessApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
