Add-Type -AssemblyName System.Speech
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class Keyboard {
        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);
        public const int VK_LCONTROL = 0xA2; // Left Ctrl
    }
"@

# Define log file path
$logFile = Join-Path -Path $PSScriptRoot -ChildPath "log.txt"

# Log function (file only, with console option)
function Write-Log($message, [string]$ConsoleMessage) {
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] $message"
    Add-Content -Path $logFile -Value $logMessage
    if ($ConsoleMessage) {
        Write-Host $ConsoleMessage
    }
}

# API setup - Should be set as environment variable GrokAPIKey
$apiKey = $env:GrokAPIKey
$url = "https://api.x.ai/v1/chat/completions"
$Grokory = New-Object System.Speech.Synthesis.SpeechSynthesizer
$headers = @{
    "Content-Type" = "application/json"
    "Authorization" = "Bearer $apiKey"
}

# Clear log file
if (Test-Path $logFile) {
    Clear-Content $logFile
}

Write-Log "Initializing System.Speech Recognizer..."
$recognizer = New-Object System.Speech.Recognition.SpeechRecognitionEngine
if (-not $recognizer) {
    Write-Log "Failed to create recognizer."
    exit
}

Write-Log "Loading dictation grammar..."
$grammar = New-Object System.Speech.Recognition.DictationGrammar
$recognizer.LoadGrammar($grammar)
Write-Log "Dictation grammar loaded."

Write-Log "Setting audio input..."
try {
    $recognizer.SetInputToDefaultAudioDevice()
    Write-Log "Audio input set."
}
catch {
    Write-Log "No microphone detected: $_"
    Write-Host "Error: No microphone detected. Please connect a microphone and try again."
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    if ($recognizer) {
        $recognizer.Dispose()
    }
    exit
}

Write-Log "Setting initial silence timeout to 10 seconds..."
$recognizer.InitialSilenceTimeout = [TimeSpan]::FromSeconds(10)

Write-Log "Initial recognizer state: $($recognizer.RecognizerState)"

# Display startup message in console
Write-Host "Starting speech recognition loop. Press Left Ctrl to speak to Grok, release for response. Esc to quit..."

Write-Log "Speech recognition loop started."

# Main loop
$continue = $true
while ($continue) {
    try {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq [ConsoleKey]::Escape) {
                Write-Log "Esc key pressed, exiting..."
                $continue = $false
                continue
            }
        }

        # Check Left Ctrl state
        if ([Keyboard]::GetAsyncKeyState([Keyboard]::VK_LCONTROL) -lt 0) {  # Key is down
            Write-Log "Left Ctrl pressed, starting recognition..."
            $result = $recognizer.Recognize([TimeSpan]::FromSeconds(5))
            if ($result) {
                $speech = $result.Text
                Write-Log "Recognized: $speech" -ConsoleMessage "You: $speech"

                # Mock response if no API key is used
                if (-not $apiKey) {
                    Write-Log "No API key set, using mock response."
                    $mockResponse = "Hi, this is Grok! Hello world! Create an environment variable named GrokAPIKey with your API key, or add it directly to line 26 in the script to connect to me."
                    Write-Log "Grok response: $mockResponse" -ConsoleMessage "Grok: $mockResponse"
                    $Grokory.Speak($mockResponse)
                } else {
                    $body = @{
                        messages = @(
                            @{
                                role = "system"
                                content = "You are Grok, a helpful AI assistant."
                            },
                            @{
                                role = "user"
                                content = $speech
                            }
                        )
                        model = "grok-2-latest"
                        stream = $false
                        temperature = 0
                    } | ConvertTo-Json -Depth 10

                    Write-Log "Sending to xAI API: $speech"
                    $ProgressPreference = 'SilentlyContinue'
                    $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
                    $ProgressPreference = 'Continue'
                    $grokText = $response.choices[0].message.content
                    Write-Log "Grok response: $grokText" -ConsoleMessage "Grok: $grokText"
                    $Grokory.Speak($grokText)
                }
            } else {
                Write-Log "No speech detected."
            }
            # Wait for Left Ctrl release
            while ([Keyboard]::GetAsyncKeyState([Keyboard]::VK_LCONTROL) -lt 0) {
                Start-Sleep -Milliseconds 100
            }
            # Clear input buffer
            while ([Console]::KeyAvailable) { $null = [Console]::ReadKey($true) }
        }
        Start-Sleep -Milliseconds 100  # Reduce CPU usage
    }
    catch {
        Write-Log "Error: $_"
        if ($_ -like "*You can purchase credits*") {
            Write-Log "Spoke: You are out of API credits" -ConsoleMessage "Grok: You are out of API credits"
            $Grokory.Speak("You are out of API credits")
        } elseif ($_ -like "*The remote name could not be resolved*") {
            Write-Log "Spoke: API connection failed. Check your internet or the API URL." -ConsoleMessage "Grok: API connection failed. Check your internet or the API URL."
            $Grokory.Speak("API connection failed. Check your internet or the API URL.")
        } else {
            Write-Log "Spoke: An error has occurred, check the console output for details" -ConsoleMessage "Grok: An error has occurred, check the console output for details"
            Write-Host "Error: $_"
            $Grokory.Speak("An error has occurred, check the console output for details")
        }
        # Clear input buffer after error to prevent lockup
        while ([Console]::KeyAvailable) { $null = [Console]::ReadKey($true) }
    }
}

# Cleanup after loop ends (via Esc)
if ($recognizer) {
    Write-Log "Disposing recognizer..."
    $recognizer.Dispose()
}
if ($Grokory) {
    Write-Log "Disposing Grokory synthesizer..."
    $Grokory.Dispose()
}
