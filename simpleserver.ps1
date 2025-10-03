# Создаём HTTP-листенер
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:80/")
$listener.Start()
Write-Host "Сервер запущен: http://localhost:80/"

while ($true) {
    # Ждём входящего запроса
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response

    # Ответ
    $responseString = "Hello from PowerShell!"
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseString)

    # Настройка заголовков
    $response.ContentLength64 = $buffer.Length
    $response.ContentType = "text/plain"

    # Запись в поток
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Close()

    Write-Host "Запрос обработан: $($request.HttpMethod) $($request.Url.AbsolutePath)"
}