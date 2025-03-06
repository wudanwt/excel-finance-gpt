# Excel Finance GPT

An Excel add-in that leverages GPT models to analyze financial data directly within Excel spreadsheets.

## Features

- Direct integration with Excel
- Financial data analysis using GPT models
- Support for custom OpenAI-compatible API endpoints
- Real-time data analysis
- Interactive user interface

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd excel-finance-gpt
```

2. Install dependencies:
```bash
npm install
```

3. Configure the environment:
```bash
cp .env.example .env
```

Edit `.env` and set your OpenAI API configuration:
```
OPENAI_API_URL=your_api_endpoint
OPENAI_API_KEY=your_api_key
```

## Development

1. Install development certificates:
```bash
npm run install-dev-certs
```

2. Start the development server:
```bash
npm run start:web
```

This will:
- Start the development server
- Generate and install necessary certificates
- Launch Excel with the add-in loaded

## Using the Add-in

1. Select a data range in Excel
2. Click the "Get Current Range" button in the add-in taskpane
3. Enter your analysis prompt
4. Click "Analyze" to get AI-powered insights

### Example Prompts

- "Analyze the trend in these financial numbers"
- "Calculate key financial ratios from this data"
- "Identify potential anomalies in this dataset"
- "Provide a summary of the financial performance"

## For Mac Users

The add-in development process on Mac requires specific steps:

1. Install certificates (one-time setup):
```bash
npm run install-dev-certs
```

2. Start the development server:
```bash
npm run start:web
```

3. The add-in should automatically load in Excel. If not:
   - Check that the development server is running at https://localhost:3000
   - Verify that the certificate is trusted in your system

## Troubleshooting

### Common Issues

1. Certificate Issues:
   - Run `npm run install-dev-certs` to reinstall certificates
   - Make sure to trust the certificate in your system

2. Add-in Not Loading:
   - Check the development server is running
   - Verify the manifest.xml is properly configured
   - Check Excel console for errors

3. API Issues:
   - Verify your API key in .env
   - Check API endpoint configuration
   - Ensure data format is correct

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

This project is licensed under the MIT License.
