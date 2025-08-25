# PowerPoint Presentation Generator

A complete, self-contained Python script that automatically generates professional PowerPoint presentations using AI content generation and real-time web search results.

## Features

- 🤖 **AI-Powered Content**: Uses Google Gemini AI to generate professional presentation content
- 🔍 **Real-time Web Search**: Integrates recent web search results for up-to-date information
- 📊 **Professional Layout**: Creates 7-slide presentations with consistent formatting
- 🎨 **Beautiful Design**: Professional color scheme and typography
- 📁 **Dynamic Naming**: Automatically generates filenames based on topic and timestamp
- 🛡️ **Error Handling**: Graceful fallback if AI generation fails

## Slide Structure

1. **Title Slide**: Compelling title + user-specified topic
2. **Overview**: 3-4 key bullet points summarizing the presentation
3. **Key Points Slide 1**: 4-5 detailed bullet points on first aspect
4. **Key Points Slide 2**: 4-5 detailed bullet points on second aspect
5. **Key Points Slide 3**: 4-5 detailed bullet points on third aspect
6. **Key Points Slide 4**: 4-5 detailed bullet points on fourth aspect
7. **Conclusion**: 3-4 main takeaways and concluding statement

## Installation

### 1. Clone or Download
Download the script files to your local machine.

### 2. Install Required Packages
```bash
pip install -r requirements.txt
```

Or install manually:
```bash
pip install python-pptx google-generativeai google-search-results python-dotenv
```

### 3. Set Up API Keys

You need two API keys:

#### **Gemini API Key** (Required)
- Visit [Google AI Studio](https://makersuite.google.com/app/apikey)
- Sign in with your Google account
- Click "Create API Key"
- Copy the generated key

#### **SerpAPI Key** (Required for web search)
- Visit [SerpAPI](https://serpapi.com/)
- Sign up for a free account
- Get your API key from the dashboard
- Free tier includes 100 searches per month

### 4. Configure API Keys

**Option A: Create a .env file (Recommended)**
Create a file named `.env` in the same directory as the script:

```env
GEMINI_API_KEY=your_actual_gemini_api_key_here
SERPAPI_API_KEY=your_actual_serpapi_key_here
```

**Option B: Set Environment Variables**
```bash
export GEMINI_API_KEY="your_actual_gemini_api_key_here"
export SERPAPI_API_KEY="your_actual_serpapi_key_here"
```

## Usage

### Basic Usage
```bash
python ppt_generator.py
```

### Interactive Process
1. Run the script
2. Enter your presentation topic when prompted
3. Wait for web search and AI content generation
4. The script will create and save your PowerPoint file
5. Open the generated `.pptx` file in PowerPoint or any compatible application

### Example Topics
- "The future of renewable energy"
- "Advancements in quantum computing"
- "The impact of AI on creative industries"
- "Climate change mitigation strategies"
- "Blockchain technology in healthcare"
- "Space exploration in the 21st century"

## File Output

The script generates files with the naming convention:
```
presentation_[topic]_[timestamp].pptx
```

Example: `presentation_renewable_energy_20241201_143022.pptx`

## Troubleshooting

### Common Issues

**"GEMINI_API_KEY environment variable not set"**
- Ensure you've created the `.env` file or set environment variables
- Check that the API key is correct and not expired

**"SERPAPI_API_KEY environment variable not set"**
- Same as above - verify your SerpAPI key is properly configured

**"AI content generation failed"**
- The script will automatically fall back to template content
- Check your internet connection and API key validity

**Import errors**
- Ensure all required packages are installed: `pip install -r requirements.txt`

### Getting Help

1. Verify your API keys are correct
2. Check your internet connection
3. Ensure all packages are installed
4. Try running with a simple topic first

## Security Notes

- **Never commit API keys to version control**
- **Use environment variables or .env files for local development**
- **Rotate API keys regularly**
- **Monitor API usage to avoid unexpected charges**

## Dependencies

- `python-pptx`: PowerPoint file creation
- `google-generativeai`: Google Gemini AI integration
- `google-search-results`: Web search functionality
- `python-dotenv`: Environment variable management

## License

This script is provided as-is for educational and personal use.

## Support

For issues related to:
- **Google Gemini API**: [Google AI Studio Support](https://ai.google.dev/support)
- **SerpAPI**: [SerpAPI Documentation](https://serpapi.com/docs)
- **python-pptx**: [python-pptx Documentation](https://python-pptx.readthedocs.io/)

---

**Happy Presentation Creating! 🎉**
