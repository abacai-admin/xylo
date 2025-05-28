# Presentation Builder

A Streamlit-based web application for creating professional presentations with data from various sources, including S&P Global Market Intelligence (CIQ) API.

## Features

- Create and customize presentation slides
- Add text, bullet points, charts, and tables
- Connect to S&P Global Market Intelligence (CIQ) API
- Preview presentations before downloading
- Export to PowerPoint (.pptx) format
- Responsive design that works on desktop and tablet

## Prerequisites

- Python 3.8 or higher
- pip (Python package manager)
- S&P Global Market Intelligence (CIQ) API credentials (optional)

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd presentation_builder
   ```

2. Create and activate a virtual environment:
   ```bash
   # On Windows
   python -m venv venv
   .\venv\Scripts\activate
   
   # On macOS/Linux
   python3 -m venv venv
   source venv/bin/activate
   ```

3. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

4. Create a `.env` file in the project root with your API credentials (optional):
   ```
   # API Configuration
   CIQ_USER=your_ciq_username
   CIQ_PASS=your_ciq_password
   API_KEY=your_api_key
   API_URL=https://api.marketdata.example.com/v1
   ```

## Usage

1. Run the Streamlit app:
   ```bash
   streamlit run app.py
   ```

2. Open your web browser and navigate to `http://localhost:8501`

3. Use the navigation menu to:
   - **Slide Builder**: Create and edit presentation slides
   - **API Configuration**: Set up API credentials
   - **Preview & Download**: Preview and download your presentation

## Project Structure

```
presentation_builder/
│
├── app.py                    # Main Streamlit app entry
│
├── pages/
│   ├── 1_Slide_Builder.py    # Slide builder UI
│   ├── 2_Config.py           # API config form
│   └── 3_Preview.py          # Preview generated slides
│
├── logic/
│   ├── slide_generator.py    # Core logic to build slides
│   ├── api_handler.py        # API fetching logic
│   └── utils.py              # Reusable helpers
│
├── assets/
│   └── template.pptx         # PowerPoint template file
│
├── .env                      # Stores user-supplied config
├── requirements.txt           # Dependency list
└── README.md
```

## Configuration

### Environment Variables

- `CIQ_USER`: S&P Global Market Intelligence username
- `CIQ_PASS`: S&P Global Market Intelligence password
- `API_KEY`: General API key for data fetching
- `API_URL`: Base URL for the API endpoint

## Features in Detail

### Slide Builder
- Create multiple slides with custom titles and content
- Add bullet points and format text
- Insert charts and tables from data
- Reorder and remove slides

### API Integration
- Connect to S&P Global Market Intelligence (CIQ) API
- Fetch financial and market data
- Configure custom API endpoints

### Preview & Export
- Real-time preview of presentation
- Download as PowerPoint (.pptx) file
- Responsive design for different screen sizes

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Streamlit](https://streamlit.io/) for the amazing web app framework
- [python-pptx](https://python-pptx.readthedocs.io/) for PowerPoint file generation
- [S&P Global Market Intelligence](https://www.spglobal.com/marketintelligence/en/) for financial data APIs

## Support

For support, please open an issue in the GitHub repository.
