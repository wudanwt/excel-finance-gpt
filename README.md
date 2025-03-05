# Excel Finance GPT

An Excel add-in that leverages OpenAI-compatible APIs to analyze financial data directly within Excel spreadsheets.

## Features

- Financial data analysis using GPT models
- Trend analysis for time-series data
- Risk assessment for financial metrics
- Real-time AI-powered insights
- Native Excel integration
- Support for custom OpenAI-compatible API endpoints

## Prerequisites

- Node.js (v14 or higher)
- npm (v6 or higher)
- Microsoft Excel (2016 or higher)
- Access to an OpenAI-compatible API endpoint

## Quick Start

1. Clone the repository:
```bash
git clone https://github.com/yourusername/excel-finance-gpt.git
cd excel-finance-gpt
```

2. Install dependencies:
```bash
npm install
```

3. Set up environment variables:
```bash
cp .env.example .env
```

4. Configure your API endpoint:
Edit `.env` and set the following variables:
```
OPENAI_API_URL=your_api_endpoint_url
OPENAI_API_KEY=your_api_key
OPENAI_API_MODEL=your_model_name
```

5. Start the development server:
```bash
npm run dev
```

## API Configuration Options

The add-in supports any API endpoint that follows the OpenAI API specification:

1. Default OpenAI API:
```
OPENAI_API_URL=https://api.openai.com/v1
```

2. Custom API endpoint:
```
OPENAI_API_URL=https://your-custom-endpoint.com/v1
```

Make sure your API endpoint implements the OpenAI API specification for the completions endpoint.

## Development Commands

- `npm run build` - Build for production
- `npm run build:dev` - Build for development
- `npm start` - Start development server
- `npm run dev` - Start development server with hot reload
- `npm run lint` - Run ESLint
- `npm run validate` - Validate the manifest
- `npm run clean` - Clean build directories

## Project Structure

```
excel-finance-gpt/
├── src/
│   ├── taskpane/          # Main application components
│   ├── services/          # Excel and API services
│   ├── utils/             # Utility functions
│   ├── types/             # TypeScript type definitions
│   └── config/           # Application configuration
├── assets/               # Static assets
├── scripts/              # Build and utility scripts
└── manifest.xml         # Add-in manifest
```

## Using the Add-in

1. Start Excel
2. Go to Insert > Office Add-ins
3. Browse to the manifest file in your project
4. Select data range in Excel
5. Choose analysis type
6. Click "Analyze" to get AI-powered insights

## Configuration Options

Edit `.env` file to configure:

- `OPENAI_API_URL` - Your API endpoint URL
- `OPENAI_API_KEY` - Your API key
- `OPENAI_API_MODEL` - Model name to use
- `DEBUG` - Enable debug logging
- `API_TIMEOUT` - API request timeout
- `MAX_TOKENS` - Maximum tokens for API requests

## Troubleshooting

Common issues:

1. **Add-in not loading**
   - Check if development server is running
   - Verify manifest path is correct
   - Check console for errors

2. **Analysis fails**
   - Verify API endpoint is correctly configured
   - Check API key is valid
   - Ensure selected range format is valid
   - Check internet connectivity

3. **API Issues**
   - Verify API endpoint follows OpenAI specification
   - Check API response format matches OpenAI format
   - Validate API key permissions

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

MIT

## Support

Create an issue in the GitHub repository for:
- Bug reports
- Feature requests
- General questions
