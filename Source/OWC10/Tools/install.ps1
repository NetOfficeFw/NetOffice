param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("OWC10Api")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
