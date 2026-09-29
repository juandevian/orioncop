param([Parameter(Mandatory)][string]$Key)
$mask = [Text.Encoding]::UTF8.GetBytes("OrionPlus-IBC")
$data = [Text.Encoding]::UTF8.GetBytes($Key)
for ($i = 0; $i -lt $data.Length; $i++) { $data[$i] = $data[$i] -bxor $mask[$i % $mask.Length] }
[Convert]::ToBase64String($data)
