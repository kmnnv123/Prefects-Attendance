# 🔥 Fingerprint Data Analyzer

A modern web application for analyzing prefect attendance data from Excel files with Firebase cloud storage integration.

## 🌟 Features

- **Excel File Processing**: Upload and analyze fingerprint attendance data from .xlsx/.xls files
- **Firebase Integration**: Cloud storage with real-time synchronization
- **Responsive Design**: Works perfectly on desktop and mobile devices
- **Data Analytics**: View attendance patterns, on-time rates, and performance metrics
- **Offline Support**: Local storage fallback when offline
- **Modern UI**: Dark theme with smooth animations and intuitive interface
- **Responsive Table**: View Excel data in a clean, responsive table format
- **CSV Export**: Download any worksheet as a CSV file
- **Data Summary**: Shows row and column counts for each worksheet

## How to Use

1. **Open the Application**: Open `index.html` in your web browser
2. **Upload Excel File**: 
   - Drag and drop an Excel file onto the upload area, OR
   - Click the upload area and browse for your file
3. **View Data**: 
   - The first worksheet will be displayed automatically
   - Click on different worksheet tabs to switch between sheets
   - Scroll through the data using the table scroll bars
4. **Export Data**: Click "Download as CSV" to export the current sheet
5. **Clear Data**: Click "Clear" to remove all data and upload a new file

## Supported File Types

- Excel 2007+ files (.xlsx)
- Excel 97-2003 files (.xls)

## Technical Details

- **Frontend**: HTML5, CSS3, JavaScript (ES6+)
- **Styling**: Bootstrap 5.3.0
- **Excel Processing**: SheetJS (xlsx library)
- **No Backend Required**: Runs entirely in the browser

## Browser Compatibility

Works in all modern browsers including:
- Chrome 60+
- Firefox 60+
- Safari 12+
- Edge 79+

## Running the Application

1. Simply open `index.html` in any web browser
2. No installation or server setup required
3. All processing happens client-side for privacy and speed

## File Size Limitations

- The application can handle moderately large Excel files
- Performance depends on your browser and device capabilities
- For very large files (>50MB), loading may take some time

## Privacy

- All file processing happens locally in your browser
- No data is sent to any server
- Your files remain completely private

## Development

To customize or extend the application:

1. Modify `index.html` for layout changes
2. Edit `script.js` for functionality updates
3. Add custom CSS styles to the `<style>` section in `index.html`

The application uses the SheetJS library for Excel file processing, which provides comprehensive support for various spreadsheet formats.
