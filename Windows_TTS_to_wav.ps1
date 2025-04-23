# Load speech system
Add-Type -AssemblyName System.Speech

# What should TTS say?
do {
    $textToSpeak = Read-Host "What should TTS say"
    # Check if the user entered anything. Otherwise ask again
    if ([string]::IsNullOrWhiteSpace($textToSpeak)) {
        Write-Host "Bruh... Just tell me what to say"
        $validInput = $false
    }
    else {
        $validInput = $true
    }
}
while (-not $validInput)

# Folder name "TTS Sounds"
$subfolderName = "TTS Sounds"

# Combines place where the script is and sub-folder name to $ttsFolder
$ttsFolder = Join-Path $PSScriptRoot $subfolderName

# Use $ttsFolder to create a sub-folder .../TTS Sounds
# And check if it's already created
if (-not (Test-Path $ttsFolder)) {
    New-Item -Path $ttsFolder -ItemType Directory | Out-Null
}

# Ask for the file name
do {
    $fileName = Read-Host "Enter the TTS file name"
    # Check if the user entered anything
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        Write-Host "What do I name this file!?"
        $validInput = $false
    }
    else {
        $validInput = $true
    }
}
while (-not $validInput)

# Makes "TTS Sounds/$fileName".wav
$wavPath = Join-Path $ttsFolder ($fileName + ".wav")

# Makes $speaker variable do TTS stuff
$speaker = New-Object System.Speech.Synthesis.SpeechSynthesizer

# Get installed voices from your system
$availableVoices = $speaker.GetInstalledVoices() | ForEach-Object { $_.VoiceInfo }

# Display voices starting from 1
Write-Host "`nAvailable Voices:`n"
for ($i = 1; $i -le $availableVoices.Count; $i++) {
    $index = $i - 1
    Write-Host "[$i] $($availableVoices[$index].Name)"
}

# Voice selection and preview
# And error checking and correcting
# If you just press Enter it'll select voice number 1
# Makes it quick to go through the program
do {
    $voiceIndex = Read-Host "Enter the number of the voice you want to use (1 to $($availableVoices.Count))"
    if ([string]::IsNullOrWhiteSpace($voiceIndex)) {
        $voiceIndex = 1
        $selectedVoice = $availableVoices[$voiceIndex - 1].Name
        $speaker.SelectVoice($selectedVoice)
        $validInput = $true
    }

    # Allow float numbers, round them, and reject non-numerics characters
    elseif ($voiceIndex -match '^\d+(\.\d+)?$') {
            $voiceIndex = [int][Math]::Round($voiceIndex)

        # Check if it's within selection range
        if ($voiceIndex -ge 1 -and $voiceIndex -le $availableVoices.Count) {
            $selectedVoice = $availableVoices[$voiceIndex - 1].Name
            $speaker.SelectVoice($selectedVoice)

            # Voice preview
            Write-Host "`nPreviewing voice: $selectedVoice"
            $speaker.Speak("Hello this is a voice preview")

            # Confirm voice
            $confirmation = Read-Host "Do you want to use this voice? (Y/N)"
            if ([string]::IsNullOrWhiteSpace($confirmation)) {
                $validInput = $true
            }
            elseif ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
                    Write-Host "Using voice: $selectedVoice"
                    $validInput = $true
            }
            else {
                Write-Host "Let's try a different voice"
                $validInput = $false
            }
        }
        else {
            Write-Host "Please choose voice from 1 to $($availableVoices.Count)"
            $validInput = $false
        }
    }
    else {
        Write-Host "Invalid input. Choose voice from 1 to $($availableVoices.Count)"
        $validInput = $false
    }
}
while (-not $validInput)

# Apply selected voice
$speaker.SelectVoice($selectedVoice)
Write-Host "Selected voice: $selectedVoice"

# Set volume
# Check for errors and round floats
# If you just press Enter it'll set volume to 100%
# Makes it quick to go through the program
do {
    $volume = Read-Host "`nSet volume (0 to 100)"
    if ([string]::IsNullOrWhiteSpace($volume)) {
        $volume = 100
        $validVolume = $true
    }

    # Round and convert only if it's a number
    elseif ($volume -match '^\d+(\.\d+)?$') {
            $volume = [int][Math]::Round($volume)

            # Clamp within 0 to 100
            if ($volume -gt 100) { $volume = 100 }
            elseif ($volume -lt 0) { $volume = 0 }
            $validVolume = $true
    }
    else {
        Write-Host "Invalid input. Set volume between 0 and 100"
        $validVolume = $false
    }
}
while (-not $validVolume)

# Apply volume
$speaker.Volume = [int]$volume
Write-Host "Volume: $volume"

# Set speech rate
# Check for errors and round floats
# If you just press Enter it'll set speech rate to 0
# Makes it quick to go through the program
do {
    $rate = Read-Host "`nSet speech rate (-10 to 10)"
    if ([string]::IsNullOrWhiteSpace($rate)) {
        $rate = 0
        $validRate = $true
    }

    # Allow negative or positive numbers
    # Round them if they're floats
    elseif ($rate -match '^[-+]?\d+(\.\d+)?$') {
            $rate = [int][Math]::Round($rate)

            # Clamp within -10 to 10
            if ($rate -gt 10) { $rate = 10 }
            elseif ($rate -lt -10) { $rate = -10 }
            $validRate = $true
    }
    else {
        Write-Host "Invalid input. Select speech rate between -10 and 10"
        $validRate = $false
    }
}
while (-not $validRate)

# Apply speech rate
$speaker.Rate = [int]$rate
Write-Host "Speech rate: $rate"

# Create $stream variable, creates file stream so I could use SetOutputToWaveStream write and output speech to .wav
$stream = New-Object System.IO.FileStream($wavPath, [System.IO.FileMode]::Create)

try {
    # Set output to the WAV file
    $speaker.SetOutputToWaveStream($stream)
    $speaker.Speak($textToSpeak)
    $stream.Close()

    # Speak TTS
    $speaker.SetOutputToDefaultAudioDevice()
    $speaker.Speak($textToSpeak)
    Write-Host "`nWAV file saved to"
    Write-Host $wavPath -ForegroundColor Green
}
catch {
    Write-Host $_ -ForegroundColor Red
    Write-Host "An error occured: $_"

    # Wait 5 seconds before closing
    Start-Sleep -Seconds 5
    exit
}
finally {
    if ($stream) { $stream.Close() }
    $speaker.Dispose()

    # Wait 5 seconds before closing
    Start-Sleep -Seconds 5
    exit
}