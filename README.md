# 🤖 Gemini Outlook Assistant

An AI-powered Microsoft Outlook add-in that integrates Google Gemini API to enhance email productivity through intelligent drafting, summarizing, and content enhancement.

## ✨ Features

### Core Capabilities
- **✍️ Draft Email** - Generate professional email responses based on context
- **📚 Summarize Email Thread** - Extract key points and decisions from long conversations
- **💬 Rephrase/Polish Text** - Improve tone, grammar, and professionalism
- **🧠 Brainstorming Replies** - Get multiple response options for different tones
- **🌐 Translate Text** - Convert email content between languages
- **🧾 Generate Formal Templates** - Create professional replies, OOO notices, etc.
- **🔍 Keyword Filtering** - Filter email content by specific terms before processing
- **🎯 Semantic Extraction** - Extract specific information like dates, decisions, action items

### Technical Features
- 🔐 Secure API key storage (browser localStorage)
- 📋 One-click copy to clipboard
- 🎨 Modern, responsive UI
- 📱 Works in Outlook Web and Desktop
- ⚡ Real-time processing with loading states

## 🚀 Quick Start

### 1. Get Your Gemini API Key
1. Visit [Google AI Studio](https://aistudio.google.com/app/apikey)
2. Sign in with your personal Google account
3. Click "Create API Key"
4. Copy the generated key (starts with `AIza...`)

### 2. Deploy to GitHub Pages
1. **Create a new GitHub repository:**
   - Go to [github.com/new](https://github.com/new)
   - Name it `gemini-outlook-assistant` (or your preferred name)
   - Set it to **Public**
   - Initialize with README

2. **Upload the project files:**
   - Download all files from this project
   - Upload to your repository:
     - `index.html`
     - `taskpane.js`
     - `manifest.xml`
     - `README.md`

3. **Enable GitHub Pages:**
   - Go to Settings → Pages in your repository
   - Set Source to "Deploy from a branch"
   - Select `main` branch and `/ (root)` folder
   - Click Save

4. **Update manifest.xml:**
   - Replace `yourusername` with your GitHub username in all URLs
   - Replace `gemini-outlook-assistant` with your repository name if different

Your add-in will be available at: `https://yourusername.github.io/repository-name/`

### 3. Sideload into Outlook

#### Outlook Web (Browser)
1. Open [Outlook Web](https://outlook.office.com)
2. Open any email
3. Click the **"..."** (More actions) button in the email toolbar
4. Select **"Get Add-ins"**
5. Go to **"My Add-ins"** tab
6. Click **"Add a custom add-in"** → **"Add from URL"**
7. Enter your manifest URL: `https://yourusername.github.io/repository-name/manifest.xml`
8. Click **"OK"** and confirm installation

#### Outlook Desktop (Windows)
1. Open Outlook Desktop
2. Go to **File** → **Manage Add-ins**
3. Click **"Add a custom add-in"** → **"Add from URL"**
4. Enter your manifest URL
5. Follow the installation prompts

## 📖 How to Use

### Basic Workflow
1. **Open the Add-in:**
   - Click the "Gemini AI" button in your Outlook ribbon
   - Or find it in the "More Apps" menu when viewing an email

2. **Enter Your API Key:**
   - Paste your Gemini API key in the top field
   - It will be saved for future sessions

3. **Choose Your Task:**
   - Click a quick action button (Draft, Summarize, etc.)
   - Or write a custom instruction

4. **Add Context (Optional):**
   - Paste email content in the "Email Content" field
   - Add a keyword filter to focus on specific content

5. **Generate and Use:**
   - Click "Generate Response"
   - Copy the result to your clipboard
   - Paste into your email draft

### Quick Action Examples

| Button | What It Does | Example Use |
|--------|-------------|-------------|
| ✍️ Draft Email | Creates professional responses | "Draft a polite decline for this meeting request" |
| 📚 Summarize | Extracts key points | "Summarize the decisions made in this email thread" |
| 💬 Rephrase | Improves tone and clarity | "Make this message more professional and concise" |
| 🧠 Brainstorm | Suggests multiple response options | "Give me 3 different ways to respond to this complaint" |
| 🌐 Translate | Converts between languages | "Translate this email to Spanish" |
| 🧾 Formal Reply | Creates business templates | "Generate a formal project update template" |

### Advanced Features

#### Keyword Filtering
Use the filter field to focus Gemini on specific content:
- Filter: "budget" → Only processes sentences mentioning budget
- Filter: "deadline" → Focuses on time-sensitive information

#### Custom Instructions
Write specific prompts for unique tasks:
- "Extract all mentioned dates and create a timeline"
- "Identify any questions that need follow-up"
- "Suggest improvements to make this email more persuasive"

## 🔧 Technical Details

### Architecture
- **Frontend**: Pure HTML/CSS/JavaScript (no build tools required)
- **API**: Google Gemini Pro model via REST API
- **Storage**: Browser localStorage for API key persistence
- **Platform**: Microsoft Office Add-ins framework

### Files Structure
```
gemini-outlook-assistant/
├── index.html        # Main UI and styling
├── taskpane.js       # Core logic and API calls
├── manifest.xml      # Outlook add-in configuration
└── README.md         # This documentation
```

### Security Notes
- API keys are stored locally in your browser only
- All API calls go directly from your browser to Google's servers
- No data is stored on external servers
- Code is open source and fully auditable

## 🛠️ Customization

### Adding New Templates
Edit the `templates` object in `taskpane.js`:

```javascript
const templates = {
    yourTemplate: {
        instruction: "Your custom instruction for Gemini",
        placeholder: "Description for the user"
    }
};
```

### Modifying the UI
- Edit `index.html` for layout changes
- Modify the `<style>` section for appearance
- Add new buttons to the template grid

### Changing API Models
To use different Gemini models, update the API endpoint in `taskpane.js`:

```javascript
// Current: gemini-pro
const response = await fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=' + apiKey

// For other models (when available):
// gemini-pro-vision (for image analysis)
// text-bison-001 (legacy)
```

## 🐛 Troubleshooting

### Common Issues

#### "Invalid API Key" Error
- **Cause**: API key from wrong source or incorrectly formatted
- **Solution**: Ensure you got the key from [AI Studio](https://aistudio.google.com/app/apikey), not Google Cloud Console
- **Check**: Key should start with `AIza` and be about 39 characters long

#### Add-in Not Loading
- **Cause**: Manifest URL issues or GitHub Pages not enabled
- **Solution**: 
  1. Verify GitHub Pages is enabled and working
  2. Check that manifest.xml URLs point to your actual GitHub Pages URL
  3. Ensure repository is public

#### CORS Errors
- **Cause**: Browser blocking cross-origin requests
- **Solution**: 
  1. Ensure you're using HTTPS (GitHub Pages automatically provides this)
  2. Check that all URLs in manifest.xml use HTTPS
  3. Try in a different browser or incognito mode

#### "No response generated" Error
- **Cause**: Gemini API returned empty or filtered response
- **Solution**:
  1. Try a simpler, more direct prompt
  2. Check if your content triggered Gemini's safety filters
  3. Ensure your Google account has Gemini API access

### Debug Mode
To enable console logging for debugging:

1. Open browser Developer Tools (F12)
2. Go to Console tab
3. Look for error messages when using the add-in
4. Check Network tab to see API request/response details

## 🔄 Updates and Maintenance

### Updating the Add-in
1. Make changes to your GitHub repository files
2. GitHub Pages will automatically deploy updates
3. Users may need to refresh Outlook or restart the application
4. For manifest changes, users may need to reinstall the add-in

### Version Control
- Use git tags for releases: `git tag v1.0.0`
- Update version in manifest.xml for major changes
- Document changes in README or CHANGELOG

## 🤝 Contributing

### Development Setup
1. Fork this repository
2. Clone to your local machine
3. Make changes and test with local web server
4. Submit pull request with description of changes

### Testing
- Test in both Outlook Web and Desktop
- Verify with different email types (plain text, HTML, threads)
- Check with various prompt types and lengths
- Ensure error handling works correctly

## 📄 License

This project is open source and available under the MIT License.

## 🆘 Support

### Getting Help
- **Issues**: Create an issue on GitHub for bugs or feature requests
- **Documentation**: Check this README for common solutions
- **API**: Refer to [Google Gemini API docs](https://ai.google.dev/docs) for API-related questions

### Useful Links
- [Google AI Studio](https://aistudio.google.com/) - Get your API key
- [Outlook Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/) - Microsoft's official docs
- [GitHub Pages Guide](https://pages.github.com/) - Hosting documentation

## 🎯 Roadmap

### Planned Features
- 📎 **Attachment Analysis** - Analyze documents and images with Gemini Pro Vision
- 🔄 **Auto-refresh** - Automatically reload current email content
- 💾 **Response History** - Save and reuse previous generations
- 👥 **Team Templates** - Share custom prompts across organizations
- 📊 **Usage Analytics** - Track most-used features and prompts
- 🎨 **Theme Support** - Dark mode and Outlook theme integration

### Version History
- **v1.0.0** - Initial release with core features
- Future versions will be tracked in GitHub releases

---

## 🚀 Ready to Deploy?

1. **Get your Gemini API key** from [AI Studio](https://aistudio.google.com/app/apikey)
2. **Create your GitHub repository** and upload these files
3. **Enable GitHub Pages** in repository settings
4. **Update manifest.xml** with your URLs
5. **Sideload into Outlook** using the manifest URL
6. **Start using Gemini** to supercharge your email workflow!

---

*Made with ❤️ for productivity enthusiasts. Powered by Google Gemini AI.*
