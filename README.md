# Features

- **Account Summary:** Displays cash info, open positions, historical transactions, and pie details.
- **Advanced Account Info:** Provides detailed transaction history, hold time statistics, fee breakdown, win/loss statistics, and visual graphs for capital gains and dividends.
- **AI Portfolio Analysis:** Offers insights and recommendations based on your portfolio data using OpenAI's GPT model.

## Setup

1.  **Create & activate virtual environment:**
    ```bash
    cd path/to/your/project
    python3 -m venv venv
    source venv/bin/activate
    ```
1.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
2.  **Run the script:**
    ```bash
    python code/main.py
    ```
    The script will prompt you for your Trading212 API key, whether you are using a demo account, and your OpenAI API key (optional).

## Output

-   `AccountAnalysis.xlsx`: An Excel file containing your Trading212 data and analysis.

## Dependencies

-   python-dotenv
-   requests
-   openai
-   openpyxl
-   Pillow
-   yfinance
-   matplotlib

## Disclaimer

This project uses the Trading212 API, which is still in beta. Account figures might occasionally be slightly off due to the nature of T212's internal logic APIs. Use this tool at your own risk.
