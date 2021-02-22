$registryPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
$Name = "DisableMergeInstance"
$value = "1"

New-ItemProperty -Path $registryPath -Name $name -Value $value `
    -PropertyType DWORD -Force | Out-Null



