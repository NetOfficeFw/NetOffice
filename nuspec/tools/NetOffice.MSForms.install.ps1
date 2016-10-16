param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("MSFormsApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
