# 🤖 AI_Desktop_Assistant

## 📘 Project Overview
The **AI Desktop Assistant** is an intelligent voice-controlled virtual assistant built using **Python**.  
It can perform daily computer tasks, answer user queries, open applications, search the web, play music, send emails, and more — all through **natural language commands**.

This project combines **Speech Recognition**, **Text-to-Speech (TTS)**, and **Natural Language Processing (NLP)** techniques to create a hands-free interactive desktop experience.

---

## 🎯 Objective
- Create a **voice-enabled desktop assistant** capable of understanding and executing commands.  
- Integrate **AI and NLP** for smarter response handling.  
- Automate daily system tasks such as opening apps, searching Google, sending messages, and more.  
- Provide both **voice and text-based interactions**.

---

## 🧠 Features
- 🎙️ **Speech Recognition:** Listens and converts your voice to text.  
- 🗣️ **Text-to-Speech Conversion:** Responds in a natural human-like voice.  
- 🌐 **Web Search:** Searches Google, Wikipedia, or YouTube.  
- 📧 **Email & Messaging:** Sends emails or messages using your voice.  
- 🖥️ **System Control:** Opens applications, folders, and files.  
- 🕒 **Date & Time Updates:** Tells current time, date, and weather info.  
- 🎵 **Entertainment:** Plays songs or opens YouTube.  
- 💬 **Conversational Mode:** Responds to greetings and small talk.  

---

## ⚙️ Technologies Used
- **Programming Language:** Python  
- **Libraries:**  
  - `speech_recognition` – for voice input  
  - `pyttsx3` – for text-to-speech output  
  - `wikipedia` – for fetching knowledge-based responses  
  - `webbrowser` – to open websites  
  - `os` – to control system functions  
  - `datetime` – to get date and time  
  - `smtplib` – for sending emails  
  - `openai` / `gemini` – (optional) for advanced conversational capabilities  

---

## 🧩 How It Works
1. The assistant continuously listens for voice commands using `speech_recognition`.  
2. The captured speech is converted into text and processed through logic or AI model.  
3. Depending on the command:
   - Opens applications or websites.  
   - Retrieves data from the internet (e.g., Wikipedia).  
   - Replies via `pyttsx3` speech output.  

---

## 🖥️ Example Commands
| Command | Action |
|----------|---------|
| "Open YouTube" | Opens YouTube in default browser |
| "What’s the time?" | Speaks current system time |
| "Search Wikipedia for Python" | Reads summary about Python |
| "Play music" | Opens music folder |
| "Send an email" | Prompts for recipient and message |
| "Who are you?" | Responds as your AI assistant |

---
