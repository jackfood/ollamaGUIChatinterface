# Ollama Desktop Chat Interface

A professional-grade, local-first Python GUI designed for seamless interaction with **Ollama**. This interface is specifically optimized for Windows users and features specialized support for **Intel IPEX-LLM** hardware acceleration.

![Python](https://img.shields.io/badge/Python-3.9+-blue?logo=python)
![Ollama](https://img.shields.io/badge/Ollama-Local-orange?logo=ollama)
![Windows](https://img.shields.io/badge/Platform-Windows-0078D4?logo=windows)
![Intel](https://img.shields.io/badge/Optimized-Intel_Arc_/_iGPU-0071C5?logo=intel)

---

## ✨ Key Features

*   **Rich Markdown Rendering:** High-fidelity rendering of headers, nested lists, bold/italic formatting, and strikethroughs.
*   **Office-Ready Clipboard:** Custom implementation using Windows API (`ctypes`) to copy content as **HTML**. Paste directly into **Microsoft Word, Outlook, or PowerPoint** with all formatting (tables, bold, code blocks) intact.
*   **Syntax Highlighted Code Blocks:** Beautifully styled code blocks with a dedicated "Copy Code" button.
*   **Persistent Sessions:** Automatic saving of chat history. Create, delete, and switch between multiple conversations.
*   **Intel IPEX-LLM Integration:** Built-in server manager to start and stop Intel-optimized Ollama instances.
*   **Dynamic Model Settings:** 
    *   Switch models on the fly.
    *   Adjust Temperature and System Prompts.
    *   **Context Length Auto-Detection:** Automatically detects model context limits (supporting up to 128k/131,072 tokens).
*   **Modern UI:** A clean, responsive interface built with `CustomTkinter` supporting both Dark and Light modes.

---

## 🖥️ System Requirements

*   **OS:** Windows 10 or 11 (64-bit required for clipboard features).
*   **Python:** 3.9 or higher.
*   **Hardware:** 
    *   Standard version: Any CPU/GPU supported by Ollama.
    *   Intel version: Intel Core Ultra, Intel Arc GPU, or Intel Iris Xe Graphics.

---

## 🚀 Installation

### 1. Clone the Repository
```bash
git clone https://github.com/your-username/ollama-chat-interface.git
cd ollama-chat-interface
```

### 2. Install Dependencies
```bash
pip install customtkinter requests openai
```

---

## 🧠 Setting Up Ollama (Intel Optimized)

For users with Intel hardware, using the **IPEX-LLM** version of Ollama provides significant performance gains by utilizing the iGPU or Arc GPU.

1.  **Download:** Navigate to the [Intel Analytics IPEX-LLM Releases Page](https://github.com/intel-analytics/ipex-llm/releases).
2.  **Select Package:** Download the latest `ollama-ipex-llm-windows-v...zip`.
3.  **Extract:** Unzip the folder to a preferred location (e.g., `C:\AI\Ollama-Intel`).
4.  **Identify Bootstrapper:** Inside the folder, look for `start-ollama.bat`. This is the file you will link in the Chat Interface settings.

---

## 🛠️ Usage Guide

### Initial Setup
1.  Copy the py into your python directory. Run the application:
    ```bash
    python OllamaChatInterface.py
    ```
2.  **Link Ollama:** The app will prompt you to find your Ollama executable. One-time set up. Browse and select either `ollama.exe` (Standard) or `start-ollama.bat` (Intel).
3.  **Start Server:** Click the **Start Server** button in the sidebar or directly from the pop up box. The status indicator will turn green once ready.

### Chatting
*   **New Chat:** Click the `+` button in the sidebar to start a fresh session.
*   **Send Message:** Type your prompt and press `Ctrl + Enter`.
*   **Formatting:** Use standard Markdown. If the AI provides a table or code, use the copy buttons to transfer them to other apps easily.

### Settings Panel
*   **Model Selection:** Choose from your locally pulled models. Use the 🔄 button to refresh the list.
*   **System Prompt:** Set the persona of the AI.
*   **Context Length:** Adjust how much memory the AI utilizes. The slider automatically adjusts its maximum based on the selected model's capabilities.

---

## 📋 Technical Implementation Details

### The "Copy as HTML" Feature
Standard Python clipboard libraries often strip formatting. This script uses `ctypes` to interface directly with `user32.dll` and `kernel32.dll`. By registering the `HTML Format` clipboard type, the script wraps Markdown-converted HTML into a specific Microsoft-compliant header. This allows for:
- Tables to remain tables in Excel/Word.
- Syntax highlights to remain colored in Outlook.
- Headers to remain as "Heading" styles in Word.

### Smart Streaming
The interface uses a token-counting threshold to switch between plain text and formatted rendering during streaming. This ensures the UI remains responsive and does not "flicker" while the AI is generating long responses.

---

## 📂 File Structure
- `settings.json`: Stores your selected model, theme, paths, and AI parameters.
- `sessions.json`: Stores all chat history and metadata.
- `main.py`: The primary script containing the UI and logic.

---

## 🤝 Troubleshooting

*   **Server Offline:** Ensure no other Ollama instances are running in the system tray before clicking "Start Server".
*   **High RAM Usage:** If the app becomes sluggish with long histories, try reducing the **Context Length** in the settings panel.
*   **Model Not Listed:** Use your terminal to run `ollama pull <model-name>` first, then click the refresh button in the app.

---

## 🔗 Links
- [Ollama Official Site](https://ollama.com/)
- [Intel IPEX-LLM GitHub](https://github.com/intel-analytics/ipex-llm)
- [CustomTkinter Documentation](https://customtkinter.tomschimansky.com/)
