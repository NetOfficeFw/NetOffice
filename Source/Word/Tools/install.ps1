param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("WordApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
