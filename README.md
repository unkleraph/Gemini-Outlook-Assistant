# Gemini-Outlook Assistant Add-in

## ✨ Overview
This add-in integrates [Google Gemini](https://aistudio.google.com/) with Microsoft Outlook to enhance productivity through AI-powered summarization, drafting, translation, and content extraction — right inside your inbox.

---

## 🚀 Features
- 🔍 **Filter** email threads by keyword
- 🧠 **Extract** semantic content (e.g., key decisions, tasks)
- ✍️ **Draft** emails with tone control
- 📚 **Summarize** long threads
- 💬 **Rephrase** or polish professional language
- 🌐 **Translate** between languages
- 🧾 **Generate templates** like out-of-office replies or meeting confirmations

---

## 📁 Project Structure
```
Gemini-Outlook-Assistant/
├── index.html        # Taskpane interface
├── taskpane.js       # Logic: filtering + API call
├── manifest.xml      # Outlook manifest for sideloading
├── README.md         # This file
```

---

## 🧰 Prerequisites
- Personal Google Account with access to [Google AI Studio](https://aistudio.google.com/)
- Outlook Web or Desktop (Windows/Mac)
- Microsoft 365 account (with sideload permission)

---

## 🔑 Get a Gemini API Key
1. Go to [https://aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)
2. Select or create a project
3. Click **Get API Key**
4. Copy the key (starts with `AIza...`)

---

## 🌐 Deploy to GitHub Pages
1. Create a new public repo on GitHub
2. Upload all files from this project
3. Go to **Settings > Pages**
4. Under "Source", choose `main` and root (`/`) folder
5. GitHub will give you a link like:
   ```
   https://yourusername.github.io/gemini-outlook-assistant/
   ```

---

## 📨 Sideload into Outlook (Web/Desktop)
### Option A: Outlook Web
1. Visit [https://outlook.office.com](https://outlook.office.com)
2. Open an email → click **More Apps** (⋯)
3. Click **Add Apps** → **Upload custom add-in** → **From URL**
4. Paste the raw GitHub Pages URL to your `manifest.xml` file

### Option B: Outlook Desktop
1. File → Manage Add-ins (opens browser)
2. Upload from URL or local file

---

## 🧪 Testing the Add-in
1. Open any email
2. Open **Gemini Assistant** from the ribbon or More Apps menu
3. Paste your API key
4. Enter a keyword to filter the email body (e.g., "invoice")
5. Give an instruction (e.g., "Summarize the action items")
6. View Gemini's AI response

---

## 🧱 Built With
- HTML + JavaScript (no frameworks)
- Gemini API (PaLM/Gemini models)
- Outlook JavaScript API (planned for future versions)

---

## 📄 License
MIT — free for personal and educational use.

---

## 🙋 Support or Questions
Open an issue or contact the author for help or custom integrations.
