param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("MSComctlLibApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
