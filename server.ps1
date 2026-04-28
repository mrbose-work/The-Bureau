$port = 8080
$listener = New-Object -TypeName System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Prefixes.Add("http://127.0.0.1:$port/")
$listener.Start()
Write-Host "Listening on port $port..."

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        $localPath = "$pwd" + $request.Url.LocalPath.Replace('/', '\')
        if ($localPath.EndsWith('\')) {
            $localPath += "index.html"
        }
        
        if (Test-Path $localPath -PathType Leaf) {
            try {
                $content = [System.IO.File]::ReadAllBytes($localPath)
                $response.ContentLength64 = $content.Length
                
                if ($localPath.EndsWith(".html")) { $response.ContentType = "text/html" }
                elseif ($localPath.EndsWith(".css")) { $response.ContentType = "text/css" }
                elseif ($localPath.EndsWith(".js")) { $response.ContentType = "application/javascript" }
                elseif ($localPath.EndsWith(".png")) { $response.ContentType = "image/png" }
                
                $response.OutputStream.Write($content, 0, $content.Length)
            } catch {
                $response.StatusCode = 500
            }
        } else {
            $response.StatusCode = 404
        }
        $response.Close()
    }
} finally {
    $listener.Stop()
}
