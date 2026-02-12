# Configuration Examples for LocalWriter

## Ollama (Local)

Copy one of the examples below to your LibreOffice config folder as `localwriter.json` (see [Configuration File Location](#configuration-file-location) for the path).

```json
{
    "endpoint": "http://localhost:11434",
    "model": "llama2",
    "api_key": "",
    "api_type": "chat",
    "is_openwebui": false,
    "openai_compatibility": false,
    "extend_selection_max_tokens": 70,
    "extend_selection_system_prompt": "",
    "edit_selection_max_new_tokens": 0,
    "edit_selection_system_prompt": ""
}
```

## OpenWebUI (Local)

```json
{
    "endpoint": "http://localhost:3000",
    "model": "llama2",
    "api_key": "",
    "api_type": "chat",
    "is_openwebui": true,
    "openai_compatibility": false,
    "extend_selection_max_tokens": 70,
    "extend_selection_system_prompt": "",
    "edit_selection_max_new_tokens": 0,
    "edit_selection_system_prompt": ""
}
```

## OpenAI

```json
{
    "endpoint": "https://api.openai.com",
    "model": "gpt-3.5-turbo",
    "api_key": "YOUR_API_KEY_HERE",
    "api_type": "chat",
    "is_openwebui": false,
    "openai_compatibility": true,
    "extend_selection_max_tokens": 100,
    "extend_selection_system_prompt": "",
    "edit_selection_max_new_tokens": 50,
    "edit_selection_system_prompt": ""
}
```

**IMPORTANT:** Never commit your actual API keys to git!

## LM Studio (Local)

```json
{
    "endpoint": "http://localhost:1234",
    "model": "local-model",
    "api_key": "",
    "api_type": "chat",
    "is_openwebui": false,
    "openai_compatibility": false,
    "extend_selection_max_tokens": 70,
    "extend_selection_system_prompt": "",
    "edit_selection_max_new_tokens": 0,
    "edit_selection_system_prompt": ""
}
```

## Configuration File Location

Copy one of the examples above to your LibreOffice user config directory as `localwriter.json`:
- macOS: `~/Library/Application Support/LibreOffice/4/user/localwriter.json`
- Linux: `~/.config/libreoffice/4/user/localwriter.json`
- Windows: `%APPDATA%\LibreOffice\4\user\localwriter.json`

Replace `YOUR_API_KEY_HERE` with your actual API key if needed.
