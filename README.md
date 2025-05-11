# Automated Excel Dashboard Generator

## Overview
This project automates the creation of sales dashboards from Excel data, featuring data analysis, visualization, and automated email distribution. It's designed to streamline the process of generating and sharing sales insights.

## Features
- **Data Processing**: Automatically processes sales data from Excel files
- **Dynamic Dashboard Creation**: Generates interactive Excel dashboards with:
  - Monthly sales trends
  - Profit analysis
  - Interactive charts and visualizations
- **Automated Email Distribution**: Sends the generated dashboard via email
- **Data Cleaning**: Handles missing values and data formatting

## Technical Stack
- Python 3.9+
- Pandas: Data manipulation and analysis
- Openpyxl: Excel file handling and chart creation
- SMTP: Automated email functionality

## Prerequisites
```bash
pip install pandas openpyxl xlwings smtplib
```

## Usage
1. Place your sales data Excel file in the project directory
2. Update the email configuration in the script
3. Run the script:
```bash
python Superstore_Sales_Automated_Excel.py
```

## Project Structure
```
├── Superstore_Sales_Automated_Excel.py  # Main script
├── superstore_sales.xlsx               # Input data file
├── sales_dashboard.xlsx                # Generated dashboard
└── README.md                          # Project documentation
```

## Features in Detail
- **Data Analysis**:
  - Monthly sales aggregation
  - Profit tracking
  - Trend analysis
- **Visualization**:
  - Line charts for sales trends
  - Profit analysis charts
- **Automation**:
  - Automated dashboard generation
  - Scheduled email distribution

## Contributing
Feel free to submit issues and enhancement requests!

## License
This project is licensed under the MIT License - see the LICENSE file for details.

## Author
Aditya Mogadpally 