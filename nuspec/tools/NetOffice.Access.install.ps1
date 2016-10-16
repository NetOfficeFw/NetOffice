param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("AccessApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}

$ref = $project.Object.References.Item("DAOApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}

$ref = $project.Object.References.Item("ADODBApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}

$ref = $project.Object.References.Item("OWC10Api")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}

$ref = $project.Object.References.Item("MSDATASRCApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}

$ref = $project.Object.References.Item("MSComctlLibApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
