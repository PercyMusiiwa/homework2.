# Stock Analysis VBA Script

## Overview
This VBA script is designed to analyze stock data in an Excel workbook. It calculates the yearly change, percentage change, and total stock volume for each stock, as well as identifying the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. The script is intended to be used with Excel workbooks where each worksheet represents data for a different year.

## Features
- Calculates yearly change from opening price to closing price.
- Calculates percentage change from opening price to closing price.
- Calculates total stock volume.
- Highlights positive changes in green and negative changes in red.
- Identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

## How to Use
1. Open your Excel workbook containing the stock data.

2. Enable the Developer tab in Excel (if not already enabled). You can do this by going to `File > Options > Customize Ribbon` and checking the "Developer" option.

3. Press `Alt + F11` to open the VBA editor.

4. Insert a new module by right-clicking on "VBAProject (YourWorkbookName)" and selecting "Insert" > "Module."

5. Copy and paste the VBA code from this repository into the module.

6. Save your workbook as a macro-enabled file with the ".xlsm" extension.

7. Close the VBA editor.

8. Press `Alt + F8`, select "StockAnalysis," and click "Run" to execute the script on all worksheets.

9. The script will calculate the summary information and highlight positive and negative changes.

10. The stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume will be displayed in the worksheet.

## Notes
- Ensure your data is structured with each year's data in separate worksheets.
- Make sure your worksheet columns match the format: <ticker> <date> <open> <high> <low> <close> <vol>.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments
- This script was developed as a learning exercise and can be freely used and modified.

Feel free to customize this README file further to include any additional information or specific instructions for your users. Be sure to replace placeholders like "YourWorkbookName" and provide proper licensing information as needed for your project.
