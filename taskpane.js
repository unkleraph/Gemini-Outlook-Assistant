// Global variables
let lastResponse = '';

// Template configurations
const templates = {
    draft: {
        instruction: "Draft a professional email response based on the context provided. Keep it concise and appropriate for business communication.",
        placeholder: "Draft Email: Create a professional response..."
    },
    summarize: {
        instruction: "Summarize the key points, decisions, and action items from this email content. Use bullet points for clarity.",
        placeholder: "Summarize: Extract key points and decisions..."
    },
    rephrase: {
        instruction: "Rephrase the following text to be more professional, clear, and polished while maintaining the original meaning.",
        placeholder: "Rephrase: Make this text more professional..."
    },
    brainstorm: {
        instruction: "Suggest 3-4 different ways to respond to this email, covering different tones (professional, friendly, direct).",
        placeholder: "Brainstorm: Suggest response options..."
    },
    translate: {
        instruction: "Translate the following text to English (or specify target language). Maintain professional tone.",
        placeholder: "Translate: Convert to target language..."
    },
    formal: {
        instruction: "Generate a formal business template response. Include appropriate greetings, body, and professional closing.",
        placeholder: "Formal Template: Create professional reply..."
    }
};

// Initialize Office Add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('Gemini Outlook Assistant loaded successfully');
        
        // Try to load saved API key
        const savedKey = localStorage.getItem('gemini_api_key');
        if (savedKey) {
            document.getElementById('apiKey').value = savedKey;
        }

        // Try to get current email content
        loadCurrentEmail();
    }
});

// Load current email content if available
function loadCurrentEmail() {
    try {
        if (Office.context.mailbox && Office.context.mailbox.item) {
            const item = Office.context.mailbox.item;
            
            // Get email body
            if (item.body) {
                item.body.getAsync(Office.CoercionType.Text, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        document.getElementById('emailContent').value = result.value;
                    }
                });
            }
        }
    } catch (error) {
        console.log('Could not access current email:', error.message);
    }
}

// Set template based on button clicked
function setTemplate(templateKey) {
    const template = templates[templateKey];
    if (template) {
        document.getElementById('instruction').value = template.instruction;
        document.getElementById('instruction').placeholder = template.placeholder;
    }
}

// Filter content by search term
function filterContent(content, searchTerm) {
    if (!searchTerm || !content) return content;
    
    const lines = content.split('\n');
    const filteredLines = lines.filter(line => 
        line.toLowerCase().includes(searchTerm.toLowerCase())
    );
    
    return filteredLines.length > 0 ? filteredLines.join('\n') : content;
}

// Main function to generate response
async function generateResponse() {
    const apiKey = document.getElementById('apiKey').value.trim();
    const emailContent = document.getElementById('emailContent').value.trim();
    const searchFilter = document.getElementById('searchFilter').value.trim();
    const instruction = document.getElementById('instruction').value.trim();
    
    // Validation
    if (!apiKey) {
        showError('Please enter your Gemini API key');
        return;
    }
    
    if (!instruction) {
        showError('Please enter an instruction for Gemini');
        return;
    }
    
    // Save API key for future use
    localStorage.setItem('gemini_api_key', apiKey);
    
    // Show loading state
    const generateBtn = document.getElementById('generateBtn');
    const resultsDiv = document.getElementById('results');
    
    generateBtn.disabled = true;
    generateBtn.textContent = 'Generating...';
    
    try {
        // Filter content if search term provided
        let processedContent = emailContent;
        if (searchFilter && emailContent) {
            processedContent = filterContent(emailContent, searchFilter);
        }
        
        // Construct prompt
        let prompt = instruction;
        if (processedContent) {
            prompt += `\n\nEmail content:\n${processedContent}`;
        }
        
        // Call Gemini API with updated endpoint
        const response = await callGeminiAPI(apiKey, prompt);
        
        // Display results
        showResults(response);
        
    } catch (error) {
        showError(`Error: ${error.message}`);
    } finally {
        generateBtn.disabled = false;
        generateBtn.textContent = 'Generate Response';
    }
}

// Call Gemini API with updated endpoint
async function callGeminiAPI(apiKey, prompt) {
    // Try multiple model endpoints in order of preference
    const modelEndpoints = [
        'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent',
        'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent',
        'https://generativelanguage.googleapis.com/v1/models/gemini-pro:generateContent',
        'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent'
    ];
    
    let lastError = null;
    
    for (const endpoint of modelEndpoints) {
        try {
            console.log(`Trying endpoint: ${endpoint}`);
            
            const response = await fetch(`${endpoint}?key=${apiKey}`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    contents: [{
                        parts: [{
                            text: prompt
                        }]
                    }],
                    generationConfig: {
                        temperature: 0.7,
                        topK: 40,
                        topP: 0.95,
                        maxOutputTokens: 1024,
                    }
                })
            });
            
            if (!response.ok) {
                const errorData = await response.json();
                console.log(`Endpoint ${endpoint} failed:`, errorData);
                lastError = new Error(errorData.error?.message || `HTTP ${response.status}: ${response.statusText}`);
                continue; // Try next endpoint
            }
            
            const data = await response.json();
            
            if (!data.candidates || !data.candidates[0] || !data.candidates[0].content) {
                console.log(`Endpoint ${endpoint} returned no content:`, data);
                lastError = new Error('No response generated by Gemini');
                continue; // Try next endpoint
            }
            
            console.log(`Success with endpoint: ${endpoint}`);
            return data.candidates[0].content.parts[0].text;
            
        } catch (error) {
            console.log(`Endpoint ${endpoint} error:`, error);
            lastError = error;
            continue; // Try next endpoint
        }
    }
    
    // If all endpoints failed, throw the last error
    throw lastError || new Error('All Gemini API endpoints failed');
}

// Alternative API call function for fallback
async function callGeminiAPIFallback(apiKey, prompt) {
    // Fallback to older API structure if needed
    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta2/models/text-bison-001:generateText?key=${apiKey}`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            prompt: {
                text: prompt
            },
            temperature: 0.7,
            candidateCount: 1,
            maxOutputTokens: 1024,
        })
    });
    
    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error?.message || `HTTP ${response.status}: ${response.statusText}`);
    }
    
    const data = await response.json();
    
    if (!data.candidates || !data.candidates[0] || !data.candidates[0].output) {
        throw new Error('No response generated by Gemini');
    }
    
    return data.candidates[0].output;
}

// Show results
function showResults(text) {
    const resultsDiv = document.getElementById('results');
    const copyBtn = document.getElementById('copyBtn');
    
    lastResponse = text;
    resultsDiv.textContent = text;
    resultsDiv.style.display = 'block';
    copyBtn.style.display = 'inline-block';
    
    // Scroll to results
    resultsDiv.scrollIntoView({ behavior: 'smooth' });
}

// Show error message
function showError(message) {
    const resultsDiv = document.getElementById('results');
    resultsDiv.innerHTML = `<div class="error">${message}</div>`;
    resultsDiv.style.display = 'block';
    
    document.getElementById('copyBtn').style.display = 'none';
}

// Copy to clipboard
function copyToClipboard() {
    if (lastResponse) {
        navigator.clipboard.writeText(lastResponse).then(() => {
            const copyBtn = document.getElementById('copyBtn');
            const originalText = copyBtn.textContent;
            copyBtn.textContent = '✅ Copied!';
            setTimeout(() => {
                copyBtn.textContent = originalText;
            }, 2000);
        }).catch(err => {
            console.error('Failed to copy:', err);
            // Fallback for older browsers
            const textArea = document.createElement('textarea');
            textArea.value = lastResponse;
            document.body.appendChild(textArea);
            textArea.select();
            document.execCommand('copy');
            document.body.removeChild(textArea);
        });
    }
}

// Helper function to get current email content (alternative method)
function getCurrentEmailContent() {
    return new Promise((resolve, reject) => {
        if (Office.context.mailbox && Office.context.mailbox.item) {
            const item = Office.context.mailbox.item;
            
            if (item.body) {
                item.body.getAsync(Office.CoercionType.Text, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value);
                    } else {
                        reject(new Error('Could not read email content'));
                    }
                });
            } else {
                resolve('');
            }
        } else {
            resolve('');
        }
    });
}