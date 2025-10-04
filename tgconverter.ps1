# WTFPL ChatGPT-4-turbo

# === НАСТРОЙКИ ===
$BotToken = "8125458267:AAHE7iOtZRFHvhAurP2aJnxSsL-Kn7Iu0Jo"  # ← токен Telegram-бота
$Port = 80

# === ПУТИ ===
$DownloadDir = "$PSScriptRoot\downloads"
$ConvertedDir = "$PSScriptRoot\converted"
$ConverterScript = "$PSScriptRoot\convert_to_pdf.ps1"

# Создание директорий при необходимости
New-Item -ItemType Directory -Force -Path $DownloadDir, $ConvertedDir | Out-Null

# === HTTP LISTENER ===
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$Port/")
$listener.Start()
Write-Host "🚀 Telegram Webhook Listener started on port $Port"

while ($true) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response

    try {
        if ($request.HasEntityBody) {
            $reader = [System.IO.StreamReader]::new($request.InputStream)
            $jsonText = $reader.ReadToEnd()
            $reader.Close()

            if (-not $jsonText) {
                throw "❌ Получено пустое тело запроса"
            }

            $json = $jsonText | ConvertFrom-Json
        } else {
            throw "❌ Запрос не содержит тела (HasEntityBody = false)"
        }

        $chatId = $json.message.chat.id
        $fileId = $json.message.document.file_id
        $fileName = $json.message.document.file_name

        if (-not $fileId) {
		    # Пользователь прислал не документ
		    $text = "Please send .doc .docx .ppt .pptx"

		    Invoke-RestMethod -Uri "https://api.telegram.org/bot$BotToken/sendMessage" `
		                      -Method Post `
		                      -ContentType "application/json" `
		                      -Body (@{
		                          chat_id = $json.message.chat.id
		                          text    = $text
		                      } | ConvertTo-Json -Depth 2)

		    $response.StatusCode = 200
		    $response.Close()
		    continue
		}

        Write-Host "📥 Получен файл: $fileName от пользователя $chatId"

        # 1. Получаем путь к файлу
        $getFileUrl = "https://api.telegram.org/bot$BotToken/getFile?file_id=$fileId"
        $fileInfo = Invoke-RestMethod -Uri $getFileUrl -Method Get
        $filePath = $fileInfo.result.file_path

        # 2. Скачиваем файл
        $downloadedFile = Join-Path $DownloadDir $fileName
        $fileUrl = "https://api.telegram.org/file/bot$BotToken/$filePath"
        Invoke-WebRequest -Uri $fileUrl -OutFile $downloadedFile
        Write-Host "💾 Скачан: $downloadedFile"

        # 3. Конвертация .docx → .pdf через внешний скрипт
        $convertedFile = Join-Path $ConvertedDir ([IO.Path]::GetFileNameWithoutExtension($fileName) + ".pdf")

        Write-Host "⚙️ Запуск конвертера: $ConverterScript"
        & powershell -ExecutionPolicy Bypass -File $ConverterScript -InputFile $downloadedFile -OutputFile $convertedFile

        if (-not (Test-Path $convertedFile)) {
            throw "❌ Конвертация не удалась — файл $convertedFile не найден"
        }

        Write-Host "✅ Конвертировано: $convertedFile"

        # 4. Отправляем PDF обратно пользователю (через HttpClient — работает в PowerShell 5.1)
        Add-Type -AssemblyName "System.Net.Http"

        $sendFileUrl = "https://api.telegram.org/bot$BotToken/sendDocument"

        $httpClient = New-Object System.Net.Http.HttpClient
        $content = New-Object System.Net.Http.MultipartFormDataContent

        # Добавляем Chat ID
        $chatIdContent = New-Object System.Net.Http.StringContent($chatId)
        $content.Add($chatIdContent, "chat_id")

        # Добавляем PDF файл
        $fileStream = [System.IO.File]::OpenRead($convertedFile)
		$fileContent = New-Object System.Net.Http.StreamContent($fileStream)
		$fileContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/pdf")
		
		# Назначаем имя вручную
		$contentDisposition = New-Object System.Net.Http.Headers.ContentDispositionHeaderValue("form-data")
		$contentDisposition.Name = '"document"'
		$contentDisposition.FileName = '"document.pdf"'  # ← Жёстко заданное имя
		$fileContent.Headers.ContentDisposition = $contentDisposition

		# Добавляем в запрос
		$content.Add($fileContent)
        # Отправка
        $tgResponse = $httpClient.PostAsync($sendFileUrl, $content).Result

        if ($tgResponse.IsSuccessStatusCode) {
            Write-Host "📤 Отправлен файл $convertedFile пользователю $chatId"
        } else {
            Write-Host "❌ Ошибка при отправке файла: $($tgResponse.StatusCode) $($tgResponse.ReasonPhrase)"
        }

        $response.StatusCode = 200
        $response.Close()
    } catch {
        Write-Host "❌ Ошибка: $_"

        # Пытаемся отправить сообщение об ошибке в Telegram
        try {
            $errorMessage = "Error converting file: `"$fileName`"."

            Invoke-RestMethod -Uri "https://api.telegram.org/bot$BotToken/sendMessage" `
                              -Method Post `
                              -ContentType "application/json" `
                              -Body (@{
                                  chat_id = $chatId
                                  text    = $errorMessage
                              } | ConvertTo-Json -Depth 2)
        } catch {
            Write-Host "⚠️ Не удалось отправить сообщение об ошибке пользователю: $_"
        }

        try {
            $response.StatusCode = 200
            $response.Close()
        } catch {}
    }
}


