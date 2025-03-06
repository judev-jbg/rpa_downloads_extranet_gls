# RPA Data extraction

## Description

This project automates the downloading and processing of shipment reports from the GLS Spain platform using Selenium WebDriver. The implementation interacts directly with the HTML elements of the web page, providing greater robustness and reliability compared to solutions based on GUI automation.

## Features

- Direct interaction with HTML elements via selectors
- Automatic login to the GLS platform
- Search for shipments by current date
- Export of results to Excel/HTML
- Intelligent detection of the actual format of downloaded files
- Automatic conversion to CSV using multiple strategies
- Detailed process logging
- Externalized configuration in .env file
- Compatible with headless mode (no graphical interface)

## Advantages of using Selenium

- **Increased robustness**: Direct interaction with HTML elements rather than simulating clicks or keystrokes
- **Screen position independence**: No dependency on screen coordinates or relative positions
- **Intelligent waiting**: Can wait for elements to become available before interacting with them
- **Works in the background**: Can run with the window minimized or in headless mode
- **Increased compatibility**: Works on different resolutions and screen configurations

## Requirements

- Python 3.7+
- Google Chrome installed
- ChromeDriver compatible with your version of Chrome

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/judev-jbg/rpa_downloads_extranet_gls.git
cd rpa_downloads_extranet_gls
```

### 2. Create virtual environment (optional but recommended)

```bash
python -m venv venv
source venv/bin/activate  # In Windows: venv\Scripts\activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure the .env file

Create a `.env` file in the root of the project with the following content:

```env
# System URLs
URL_LOGIN=url_login
URL_SHIPMENTS=url_shipments

# Credentials (complete with the correct values)
USERNAME_GLS=your_user
PASSWORD_GLS=your_pass

# Folder paths
PATH_DOWNLOAD_FOLDER=path_download
PATH_FINAL_FOLDER=path_final
```

### Prepare ChromeDriver

Make sure you have the ChromeDriver in the `drivers/` folder of the project:

```bash
mkdir -p drivers
# Download and place chromedriver.exe in this folder
```

## Usage

### Basic execution

```bash
python main.py
```

### Running in headless mode (without GUI)

To run in headless mode, modify the configuration in the code (variable `headless` in the `load_config()` function).

## Project structure

```
rpa_downloads_extranet_gls/
│
├── main.py            # Main entry point
├── rpa.py             # Module with RPA functionalities based on Selenium
├── .env               # Configuration file with environment variables
├── requirements.txt   # Project dependencies
├── drivers/           # Folder for ChromeDriver
│   └── chromedriver.exe
├── LICENSE            # Project license
└── README.md          # Documentation
```

## Technical description

The RPA works in several key stages:

1. **Authentication**: Logs into the GLS platform using the configured credentials.
2. **Navigation**: Accesses the shipment query page.
3. **Search**: Performs a search by the current date.
4. **Export**: Download the results in XLS/HTML format.
5. **Processing**: Detects the actual file format, extracts the data and converts it to XLSX.

### Handling HTML files with XLS extension

The RPA includes specialized logic to handle a particular case of GLS: HTML files that have XLS extension. The system:

1. automatically detects whether the content is HTML despite the extension.
2. Uses multiple strategies to extract the tables:
   - pandas.read_html()
   - BeautifulSoup for manual extraction
3. Ensure that the final result is a clean and well-formatted XLSX.

## Troubleshooting

### Common problems

1. **ChromeDriver error**: Make sure you have the correct ChromeDriver for your version of Chrome in the `drivers/` folder.
2. **Elements not found**: Selectors (IDs, names) may change if GLS updates their platform.
3. **Timeouts timed out**: Adjust the timeout values in the settings if the page takes a long time to load.
4. **Error when converting to XLSX**: Verify that all dependencies are installed correctly.

### Logs

Detailed logs are stored in:

- `rpa_shipments.log` for RPA module

## Updates and maintenance

If the structure of the GLS web page changes, you may need to update the element selectors in the `rpa.py` module. Common changes include:

- Modified element IDs
- New fields or forms
- Changes in the navigation flow

## License

This project is licensed under the terms of the MIT license. See the [LICENSE](LICENSE) file for details.

## Contributions

Contributions are welcome. Please send a Pull Request or open an Issue to discuss proposed changes.
