Set-StrictMode -Version 3.0


$Z_BLANC = '　'
$UNDER_SCORE = '_' 
function combined_name {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)][String]$first_name,
        [Parameter(Mandatory)][String]$last_name,
        #デフォルト引数
        [String]$blanc = $Z_BLANC
    )
    $sb = New-Object System.Text.StringBuilder

    #副作用処理  StringBuilderならちょっと速いらしい。要素数が少ないから意味ないかも。
    @($first_name, $blanc ,$last_name) |
    ForEach-Object{[void] $sb.Append($_)}

    return $sb.ToString()
}

function one_liner {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)][PSCustomObject[]]$name_list
    )
    $names = foreach ($name in $name_list){
        $name.replace($Z_BLANC, "")
    }
    return $names -join $UNDER_SCORE
}

exit 0