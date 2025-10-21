# Excel Data Dashboard

A powerful Flask-based web application for visualizing and analyzing Excel data with interactive charts and filtering capabilities.

## Quick Start

### 1. Create Virtual Environment
```
python -m venv venv
```

### 2. Activate Virtual Environment
#### Windows
```
venv\Scripts\activate
```
#### Mac/Linux:
```
source venv/bin/activate
```

### 3. Install Dependencies
```
pip install flask plotly pandas openpyxl
```

### 4. Run the Application
#### Windows/Linux
```
python app.py
```
Then visit http://localhost:5000

#### Mac
You need to modify the app.py file first. Change line 775 from:

``` app.run(debug=True, host='localhost', port=5000) ```

to

``` app.run(debug=True, host='localhost', port=3000) ```

#### Then run:
``` python app.py ```

And visit: http://localhost:3000 or http://localhost:4000

## Features
- 📊 Interactive Charts: Bar charts, line charts, pie charts, scatter plots, and heatmaps

- 🔍 Excel-style Filtering: Filter data by multiple columns with various operators

- 📈 Real-time Updates: Charts update instantly when filters change

- 📥 Data Export: Export filtered data to Excel format

- 📱 Responsive Design: Works on desktop and mobile devices

## Project Structure

```
dashboard/
├── app.py                 # Main Flask application
├── templates/
│   └── index.html        # Dashboard HTML template
├── static/
│   └── style.css         # CSS styles
├── data.xlsx             # Generated Excel data (auto-created)
└── venv/                 # Virtual environment
```

## Troubleshooting
- Port already in use: Mac users should use port 3000 or 4000 instead of 5000
- Missing dependencies: Ensure all packages are installed in the virtual environment
- Data not loading: Check that the virtual environment is activated
- Chart not displaying: Check browser console for JavaScript errors

### Stopping the Application
Press Ctrl+C in the terminal to stop the Flask development server.

### Deactivating Virtual Environment
When done, deactivate the virtual environment:

``` deactivate```

## Dependencies
- Flask: Web framework
- Plotly: Interactive charting library
- Pandas: Data manipulation and analysis
- OpenPyXL: Excel file handling
> Note for Mac Users: macOS reserves port 5000 for AirPlay, so you must use port 3000 or 4000 by modifying the app.py file as shown above.

