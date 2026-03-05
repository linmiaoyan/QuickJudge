
function naps2.console {
    # 定义NAPS2.Console.exe的完整路径（适配含空格的路径）
    $naps2ConsolePath = "D:\Download\NAPS2\App\NAPS2.Console.exe"
    
    # 检查文件是否存在，避免调用失败
    if (-not (Test-Path -Path $naps2ConsolePath -PathType Leaf)) {
        Write-Error "错误：未找到NAPS2控制台程序，请检查路径是否正确！路径：$naps2ConsolePath"
        return
    }
    # 调用程序并传递参数
    & $naps2ConsolePath @args
}
