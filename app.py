from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from plotly.subplots import make_subplots
import os
import numpy as np
from datetime import datetime, timedelta
import random
from io import BytesIO
import csv

app = Flask(__name__)

def load_data():
    """Load data from Excel file"""
    # Force regeneration by deleting existing file (optional)
    if os.path.exists('data.xlsx'):
        try:
            os.remove('data.xlsx')
            print("Deleted existing data file to regenerate with 50 columns")
        except:
            pass
    
    try:
        # Specify the engine explicitly
        df = pd.read_excel('data.xlsx', engine='openpyxl')
        print(f"Successfully loaded data from data.xlsx with {len(df)} rows and {len(df.columns)} columns")
        return df
    except FileNotFoundError:
        # Create sample data if file doesn't exist
        print("Excel file not found. Creating comprehensive sample data with 50 columns...")
        return create_sample_data()
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        print("Creating sample data instead...")
        return create_sample_data()

def create_sample_data():
    """Create comprehensive sample data with 1000 rows and exactly 50 columns"""
    np.random.seed(42)
    random.seed(42)
    
    n_rows = 1000
    data = {}

    # 1. IDENTIFIERS (5 columns)
    data['Transaction_ID'] = [f'TXN{str(i).zfill(6)}' for i in range(1, n_rows + 1)]
    data['Customer_ID'] = [f'CUST{str(random.randint(1, 500)).zfill(4)}' for _ in range(n_rows)]
    data['Product_ID'] = [f'PROD{str(random.randint(1, 200)).zfill(4)}' for _ in range(n_rows)]
    data['Order_ID'] = [f'ORD{str(random.randint(1, 800)).zfill(5)}' for _ in range(n_rows)]
    data['Invoice_Number'] = [f'INV{str(random.randint(1000, 9999))}' for _ in range(n_rows)]

    # 2. PRODUCT INFORMATION (10 columns)
    categories = ['Electronics', 'Clothing', 'Home & Garden', 'Sports', 'Books', 'Automotive', 'Beauty', 'Toys']
    subcategories = ['Smartphones', 'Laptops', 'Men', 'Women', 'Furniture', 'Kitchen', 'Fitness', 'Outdoor', 
                    'Fiction', 'Non-Fiction', 'Parts', 'Accessories', 'Skincare', 'Makeup', 'Educational', 'Games']
    
    data['Category'] = [random.choice(categories) for _ in range(n_rows)]
    data['Subcategory'] = [random.choice(subcategories) for _ in range(n_rows)]
    data['Product_Name'] = [f"Product_{i}" for i in range(1, n_rows + 1)]
    data['Brand'] = [random.choice(['Brand_A', 'Brand_B', 'Brand_C', 'Brand_D', 'Brand_E', 'Brand_F']) for _ in range(n_rows)]
    data['Supplier'] = [f"Supplier_{random.randint(1, 30)}" for _ in range(n_rows)]
    data['Manufacturer'] = [random.choice(['Manufacturer_A', 'Manufacturer_B', 'Manufacturer_C']) for _ in range(n_rows)]
    data['Product_Line'] = [random.choice(['Premium', 'Standard', 'Economy', 'Professional']) for _ in range(n_rows)]
    data['Product_Type'] = [random.choice(['Physical', 'Digital', 'Service', 'Subscription']) for _ in range(n_rows)]
    data['Product_Status'] = [random.choice(['Active', 'Discontinued', 'New', 'Promotional']) for _ in range(n_rows)]
    data['Warranty_Period'] = np.random.randint(0, 36, n_rows)

    # 3. GEOGRAPHIC DATA (8 columns)
    data['Region'] = [random.choice(['North America', 'Europe', 'Asia Pacific', 'Latin America', 'Middle East', 'Africa']) for _ in range(n_rows)]
    data['Country'] = [random.choice(['USA', 'Canada', 'UK', 'Germany', 'France', 'Japan', 'Australia', 'Brazil', 'India', 'China']) for _ in range(n_rows)]
    data['State'] = [random.choice(['California', 'Texas', 'New York', 'Florida', 'Illinois', 'Ontario', 'London', 'Bavaria', 'Tokyo', 'NSW']) for _ in range(n_rows)]
    data['City'] = [f"City_{random.randint(1, 50)}" for _ in range(n_rows)]
    data['Postal_Code'] = [f"{random.randint(10000, 99999)}" for _ in range(n_rows)]
    data['Store_ID'] = [f"Store_{random.randint(1, 25)}" for _ in range(n_rows)]
    data['Store_Type'] = [random.choice(['Retail', 'Online', 'Wholesale', 'Outlet']) for _ in range(n_rows)]
    data['Territory'] = [random.choice(['East', 'West', 'North', 'South', 'Central', 'International']) for _ in range(n_rows)]

    # 4. TIME DATA (6 columns)
    start_date = datetime(2022, 1, 1)
    end_date = datetime(2023, 12, 31)
    date_range = [start_date + timedelta(days=x) for x in range(0, (end_date - start_date).days + 1)]
    
    data['Order_Date'] = [random.choice(date_range) for _ in range(n_rows)]
    data['Shipping_Date'] = [date + timedelta(days=random.randint(1, 7)) for date in data['Order_Date']]
    data['Delivery_Date'] = [date + timedelta(days=random.randint(3, 14)) for date in data['Order_Date']]
    data['Quarter'] = [f"Q{(date.month-1)//3 + 1}-{date.year}" for date in data['Order_Date']]
    data['Month'] = [date.strftime("%Y-%m") for date in data['Order_Date']]
    data['Year'] = [date.year for date in data['Order_Date']]

    # 5. FINANCIAL METRICS (12 columns)
    data['Sales'] = np.random.uniform(10, 5000, n_rows).round(2)
    data['Cost'] = (np.array(data['Sales']) * np.random.uniform(0.3, 0.7, n_rows)).round(2)
    data['Profit'] = (np.array(data['Sales']) - np.array(data['Cost'])).round(2)
    data['Profit_Margin'] = (np.array(data['Profit']) / np.array(data['Sales']) * 100).round(2)
    data['Quantity'] = np.random.randint(1, 100, n_rows)
    data['Unit_Price'] = (np.array(data['Sales']) / np.array(data['Quantity'])).round(2)
    data['Unit_Cost'] = (np.array(data['Cost']) / np.array(data['Quantity'])).round(2)
    data['Discount_Rate'] = np.random.uniform(0, 0.4, n_rows).round(3)
    data['Discount_Amount'] = (np.array(data['Sales']) * np.array(data['Discount_Rate'])).round(2)
    data['Tax_Amount'] = (np.array(data['Sales']) * np.random.uniform(0.05, 0.15, n_rows)).round(2)
    data['Shipping_Cost'] = np.random.uniform(0, 100, n_rows).round(2)
    data['Total_Amount'] = (np.array(data['Sales']) + np.array(data['Tax_Amount']) + np.array(data['Shipping_Cost']) - np.array(data['Discount_Amount'])).round(2)

    # 6. CUSTOMER METRICS (9 columns)
    data['Customer_Age'] = np.random.randint(18, 80, n_rows)
    data['Customer_Segment'] = [random.choice(['Premium', 'Standard', 'Budget', 'New', 'VIP']) for _ in range(n_rows)]
    data['Customer_Rating'] = np.random.uniform(1, 5, n_rows).round(1)
    data['Customer_Tenure_Days'] = np.random.randint(0, 3650, n_rows)
    data['Loyalty_Level'] = [random.choice(['Bronze', 'Silver', 'Gold', 'Platinum']) for _ in range(n_rows)]
    data['Credit_Score'] = np.random.randint(300, 850, n_rows)
    data['Annual_Income'] = np.random.uniform(20000, 200000, n_rows).round(2)
    data['Number_of_Orders'] = np.random.randint(1, 50, n_rows)
    data['Days_Since_Last_Order'] = np.random.randint(0, 365, n_rows)

    # Create DataFrame first to check column count
    df = pd.DataFrame(data)
    
    print(f"Initial columns: {len(df.columns)}")
    
    # If we don't have 50 columns yet, add more
    if len(df.columns) < 50:
        # 7. ADDITIONAL COLUMNS to reach 50
        additional_needed = 50 - len(df.columns)
        print(f"Adding {additional_needed} more columns...")
        
        # Operational metrics
        if additional_needed > 0:
            data['Return_Rate'] = np.random.uniform(0, 0.2, n_rows).round(3)
            additional_needed -= 1
        
        if additional_needed > 0:
            data['Customer_Acquisition_Cost'] = np.random.uniform(10, 500, n_rows).round(2)
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Customer_Lifetime_Value'] = np.random.uniform(100, 10000, n_rows).round(2)
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Inventory_Level'] = np.random.randint(0, 1000, n_rows)
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Reorder_Level'] = np.random.randint(10, 200, n_rows)
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Stock_Out_Rate'] = np.random.uniform(0, 0.1, n_rows).round(3)
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Product_Rating'] = np.random.uniform(1, 5, n_rows).round(1)
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Product_Age_Days'] = np.random.randint(0, 1825, n_rows)
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Order_Status'] = [random.choice(['Completed', 'Pending', 'Cancelled', 'Shipped', 'Processing']) for _ in range(n_rows)]
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Payment_Method'] = [random.choice(['Credit Card', 'Debit Card', 'PayPal', 'Bank Transfer', 'Cash']) for _ in range(n_rows)]
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Shipping_Method'] = [random.choice(['Standard', 'Express', 'Overnight', 'International']) for _ in range(n_rows)]
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Shipping_Carrier'] = [random.choice(['UPS', 'FedEx', 'DHL', 'USPS', 'Amazon Logistics']) for _ in range(n_rows)]
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Return_Reason'] = [random.choice(['None', 'Defective', 'Wrong Item', 'Size Issue', 'Changed Mind']) for _ in range(n_rows)]
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Department'] = [random.choice(['Sales', 'Marketing', 'Operations', 'Finance', 'HR', 'IT']) for _ in range(n_rows)]
            additional_needed -= 1
            
        if additional_needed > 0:
            data['Employee_ID'] = [random.choice([f'Emp{str(i).zfill(3)}' for i in range(1, 51)]) for _ in range(n_rows)]
            additional_needed -= 1
            
        # Add any remaining columns as generic metrics
        for i in range(additional_needed):
            data[f'Metric_{i+1}'] = np.random.uniform(0, 100, n_rows).round(2)
    
    # Create final DataFrame
    df = pd.DataFrame(data)
    
    print(f"Final dataset: {len(df)} rows × {len(df.columns)} columns")
    print("Column names:", list(df.columns))
    
    # Save to Excel
    try:
        df.to_excel('data.xlsx', index=False, engine='openpyxl')
        print(f"Sample data saved to data.xlsx with {len(df)} rows and {len(df.columns)} columns")
    except Exception as e:
        print(f"Warning: Could not save sample data: {e}")
    
    return df

def create_grouped_stacked_bar_chart(df, x_axis, y_axis, stack_by, group_by=None):
    """Create grouped stacked bar chart based on selected columns"""
    try:
        if group_by and group_by != 'None':
            # Aggregate data with grouping
            pivot_df = df.groupby([x_axis, group_by, stack_by])[y_axis].sum().reset_index()
            fig = px.bar(
                pivot_df,
                x=x_axis,
                y=y_axis,
                color=stack_by,
                facet_col=group_by,
                barmode='stack',
                title=f'Grouped Stacked Bar Chart - {y_axis} by {x_axis}, {group_by} and {stack_by}',
            )
        else:
            # Aggregate data without grouping
            pivot_df = df.groupby([x_axis, stack_by])[y_axis].sum().reset_index()
            fig = px.bar(
                pivot_df,
                x=x_axis,
                y=y_axis,
                color=stack_by,
                barmode='stack',
                title=f'Stacked Bar Chart - {y_axis} by {x_axis} and {stack_by}',
            )
        
        fig.update_layout(
            xaxis_title=x_axis,
            yaxis_title=y_axis,
            legend_title=stack_by,
            height=500,
            autosize=True,
            margin=dict(l=50, r=50, t=80, b=50),
            showlegend=True
        )
        
        return pio.to_html(fig, full_html=False, config={
            'responsive': True,
            'displayModeBar': True,
            'displaylogo': False
        })
    except Exception as e:
        return f"<p>Error creating chart: {str(e)}</p>"

def create_line_chart(df, x_axis, y_axis, color_by=None):
    """Create line chart based on selected columns"""
    try:
        if color_by and color_by != 'None':
            trend_df = df.groupby([x_axis, color_by])[y_axis].sum().reset_index()
            fig = px.line(
                trend_df,
                x=x_axis,
                y=y_axis,
                color=color_by,
                title=f'{y_axis} Trend by {x_axis}',
                markers=True
            )
        else:
            trend_df = df.groupby(x_axis)[y_axis].sum().reset_index()
            fig = px.line(
                trend_df,
                x=x_axis,
                y=y_axis,
                title=f'{y_axis} Trend by {x_axis}',
                markers=True
            )
        
        fig.update_traces(line=dict(width=3))
        fig.update_layout(
            xaxis_title=x_axis,
            yaxis_title=y_axis,
            height=500,
            autosize=True,
            margin=dict(l=50, r=50, t=80, b=50),
            showlegend=True
        )
        
        return pio.to_html(fig, full_html=False, config={
            'responsive': True,
            'displayModeBar': True,
            'displaylogo': False
        })
    except Exception as e:
        return f"<p>Error creating chart: {str(e)}</p>"

def create_pie_chart(df, category_col, value_col):
    """Create pie chart based on selected columns"""
    try:
        print(f"Creating pie chart: {value_col} by {category_col}")
        
        # Check if columns exist
        if category_col not in df.columns:
            return f"<p>Error: Category column '{category_col}' not found in data</p>"
        if value_col not in df.columns:
            return f"<p>Error: Value column '{value_col}' not found in data</p>"
        
        # Aggregate data for pie chart
        pie_data = df.groupby(category_col, as_index=False)[value_col].sum()
        
        print(f"Aggregated data shape: {pie_data.shape}")
        print(f"Sample data:")
        print(pie_data.head())
        
        # Check if we have data
        if pie_data.empty:
            return f"<p>No data available for {value_col} by {category_col}</p>"
        
        fig = px.pie(
            pie_data,
            values=value_col,
            names=category_col,
            title=f'{value_col} Distribution by {category_col}',
            hover_data=[value_col]
        )
        
        fig.update_layout(
            height=500,
            margin=dict(l=50, r=50, t=80, b=50),
            showlegend=True
        )
        
        fig.update_traces(
            textposition='inside',
            textinfo='percent+label',
            hovertemplate=f"<b>%{{label}}</b><br>{value_col}: %{{value:,.2f}}<br>Percentage: %{{percent}}<extra></extra>"
        )
        
        chart_html = pio.to_html(fig, full_html=False, config={
            'responsive': True,
            'displayModeBar': True,
            'displaylogo': False
        })
        
        print(f"Generated pie chart HTML length: {len(chart_html)}")
        return chart_html
        
    except Exception as e:
        error_msg = f"<p>Error creating pie chart: {str(e)}</p>"
        print(f"Pie chart error: {str(e)}")
        return error_msg

def create_scatter_plot(df, x_axis, y_axis, color_by=None, size_by=None):
    """Create scatter plot based on selected columns"""
    try:
        fig = px.scatter(
            df,
            x=x_axis,
            y=y_axis,
            color=color_by if color_by and color_by != 'None' else None,
            size=size_by if size_by and size_by != 'None' else None,
            hover_data=df.columns.tolist(),
            title=f'{y_axis} vs {x_axis}'
        )
        
        fig.update_layout(
            xaxis_title=x_axis,
            yaxis_title=y_axis,
            height=500,
            autosize=True,
            margin=dict(l=50, r=50, t=80, b=50),
            showlegend=True
        )
        
        return pio.to_html(fig, full_html=False, config={
            'responsive': True,
            'displayModeBar': True,
            'displaylogo': False
        })
    except Exception as e:
        return f"<p>Error creating chart: {str(e)}</p>"

def create_heatmap(df, x_axis, y_axis, value_col):
    """Create heatmap based on selected columns"""
    try:
        heatmap_data = df.pivot_table(
            values=value_col, 
            index=y_axis, 
            columns=x_axis, 
            aggfunc='sum'
        ).fillna(0)
        
        fig = px.imshow(
            heatmap_data,
            title=f'{value_col} Heatmap - {y_axis} vs {x_axis}',
            aspect='auto',
            color_continuous_scale='Blues'
        )
        
        fig.update_layout(
            xaxis_title=x_axis,
            yaxis_title=y_axis,
            height=500,
            autosize=True,
            margin=dict(l=50, r=50, t=80, b=50)
        )
        
        return pio.to_html(fig, full_html=False, config={
            'responsive': True,
            'displayModeBar': True,
            'displaylogo': False
        })
    except Exception as e:
        return f"<p>Error creating chart: {str(e)}</p>"

@app.route('/')
def index():
    """Main dashboard route"""
    df = load_data()
    
    # Get all column names for filters
    numeric_columns = df.select_dtypes(include=['number']).columns.tolist()
    categorical_columns = df.select_dtypes(include=['object', 'category']).columns.tolist()
    all_columns = df.columns.tolist()
    
    # Get unique values for categorical columns (for dropdowns)
    unique_values = {}
    for column in categorical_columns:
        unique_values[column] = df[column].dropna().unique().tolist()[:100]  # Limit to 100 values
    
    # Calculate total pages for pagination
    total_pages = (len(df) + 19) // 20  # 20 rows per page
    
    # Create visualizations with default parameters
    grouped_bar_chart = create_grouped_stacked_bar_chart(
        df, 
        x_axis='Quarter', 
        y_axis='Sales', 
        stack_by='Subcategory',
        group_by='Category'
    )
    
    line_chart = create_line_chart(
        df, 
        x_axis='Quarter', 
        y_axis='Sales', 
        color_by='Category'
    )
    
    pie_chart = create_pie_chart(
        df, 
        category_col='Category', 
        value_col='Sales'
    )
    
    scatter_plot = create_scatter_plot(
        df, 
        x_axis='Sales', 
        y_axis='Profit', 
        color_by='Category'
    )
    
    heatmap = create_heatmap(
        df, 
        x_axis='Region', 
        y_axis='Category', 
        value_col='Sales'
    )
    
    # Get basic statistics
    total_sales = df['Sales'].sum() if 'Sales' in df.columns else 0
    total_profit = df['Profit'].sum() if 'Profit' in df.columns else 0
    avg_sales = df['Sales'].mean() if 'Sales' in df.columns else 0
    
    # Prepare data preview - show first 20 rows and all columns
    data_preview = df.head(20).to_dict('records')
    column_names = all_columns
    
    return render_template(
        'index.html',
        grouped_bar_chart=grouped_bar_chart,
        line_chart=line_chart,
        pie_chart=pie_chart,
        scatter_plot=scatter_plot,
        heatmap=heatmap,
        total_sales=total_sales,
        total_profit=total_profit,
        avg_sales=avg_sales,
        data_preview=data_preview,
        column_names=column_names,
        numeric_columns=numeric_columns,
        categorical_columns=categorical_columns,
        all_columns=all_columns,
        unique_values=unique_values,
        total_rows=len(df),
        total_columns=len(df.columns),
        total_pages=total_pages
    )

@app.route('/update_chart', methods=['POST'])
def update_chart():
    """Update chart based on user selections"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'error': 'No JSON data received'})
        
        chart_type = data.get('chart_type')
        filters = data.get('filters', {})
        
        print(f"=== CHART UPDATE REQUEST ===")
        print(f"Chart type: {chart_type}")
        print(f"Filters: {filters}")
        
        df = load_data()
        
        if chart_type == 'grouped_bar':
            chart_html = create_grouped_stacked_bar_chart(
                df,
                x_axis=filters.get('x_axis', 'Quarter'),
                y_axis=filters.get('y_axis', 'Sales'),
                stack_by=filters.get('stack_by', 'Subcategory'),
                group_by=filters.get('group_by', 'Category')
            )
        elif chart_type == 'line':
            chart_html = create_line_chart(
                df,
                x_axis=filters.get('x_axis', 'Quarter'),
                y_axis=filters.get('y_axis', 'Sales'),
                color_by=filters.get('color_by', 'Category')
            )
        elif chart_type == 'pie':
            # FIXED: Use 'category' and 'value' instead of 'category_col' and 'value_col'
            chart_html = create_pie_chart(
                df,
                category_col=filters.get('category', 'Category'),  # CHANGED
                value_col=filters.get('value', 'Sales')  # CHANGED
            )
        elif chart_type == 'scatter':
            chart_html = create_scatter_plot(
                df,
                x_axis=filters.get('x_axis', 'Sales'),
                y_axis=filters.get('y_axis', 'Profit'),
                color_by=filters.get('color_by', 'Category'),
                size_by=filters.get('size_by', 'None')
            )
        elif chart_type == 'heatmap':
            chart_html = create_heatmap(
                df,
                x_axis=filters.get('x_axis', 'Region'),
                y_axis=filters.get('y_axis', 'Category'),
                value_col=filters.get('value', 'Sales')  # CHANGED
            )
        else:
            return jsonify({'success': False, 'error': f'Invalid chart type: {chart_type}'})
        
        print(f"Chart HTML generated successfully, length: {len(chart_html)}")
        return jsonify({'success': True, 'chart_html': chart_html})
    
    except Exception as e:
        print(f"Error in update_chart: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/filter_data', methods=['POST'])
def filter_data():
    """Filter data based on user criteria"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'error': 'No JSON data received'})
        
        filters = data.get('filters', {})
        page = data.get('page', 1)
        page_size = data.get('page_size', 20)
        
        print(f"=== FILTER REQUEST ===")
        print(f"Filters: {filters}")
        print(f"Page: {page}, Page size: {page_size}")
        
        df = load_data()
        filtered_df = df.copy()
        
        # Apply filters
        for column, filter_config in filters.items():
            if column not in df.columns:
                continue
                
            if filter_config['type'] == 'categorical':
                # Categorical filter (multiple selection)
                values = filter_config['values']
                if values:
                    filtered_df = filtered_df[filtered_df[column].isin(values)]
                    
            elif filter_config['type'] == 'numeric':
                # Numeric filter with operators
                operator = filter_config['operator']
                value = filter_config['value']
                
                if operator == '=':
                    filtered_df = filtered_df[filtered_df[column] == value]
                elif operator == '>':
                    filtered_df = filtered_df[filtered_df[column] > value]
                elif operator == '<':
                    filtered_df = filtered_df[filtered_df[column] < value]
                elif operator == '>=':
                    filtered_df = filtered_df[filtered_df[column] >= value]
                elif operator == '<=':
                    filtered_df = filtered_df[filtered_df[column] <= value]
                elif operator == '!=':
                    filtered_df = filtered_df[filtered_df[column] != value]
        
        # Calculate pagination
        total_filtered_rows = len(filtered_df)
        start_idx = (page - 1) * page_size
        end_idx = start_idx + page_size
        paginated_df = filtered_df.iloc[start_idx:end_idx]
        
        # Convert to dictionary for JSON response
        filtered_data = paginated_df.to_dict('records')
        
        print(f"Filtered data: {total_filtered_rows} rows found")
        
        return jsonify({
            'success': True,
            'filtered_data': filtered_data,
            'total_filtered_rows': total_filtered_rows,
            'current_page': page,
            'total_pages': (total_filtered_rows + page_size - 1) // page_size
        })
        
    except Exception as e:
        print(f"Error in filter_data: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/export_data', methods=['POST'])
def export_data():
    """Export filtered data to Excel or CSV file"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'error': 'No JSON data received'})
        
        filters = data.get('filters', {})
        export_format = data.get('format', 'excel')
        
        print(f"=== EXPORT REQUEST ===")
        print(f"Filters: {filters}")
        print(f"Format: {export_format}")
        
        df = load_data()
        filtered_df = df.copy()
        
        # Apply the same filters as in filter_data
        for column, filter_config in filters.items():
            if column not in df.columns:
                continue
                
            if filter_config['type'] == 'categorical':
                # Categorical filter (multiple selection)
                values = filter_config['values']
                if values:
                    filtered_df = filtered_df[filtered_df[column].isin(values)]
                    
            elif filter_config['type'] == 'numeric':
                # Numeric filter with operators
                operator = filter_config['operator']
                value = filter_config['value']
                
                if operator == '=':
                    filtered_df = filtered_df[filtered_df[column] == value]
                elif operator == '>':
                    filtered_df = filtered_df[filtered_df[column] > value]
                elif operator == '<':
                    filtered_df = filtered_df[filtered_df[column] < value]
                elif operator == '>=':
                    filtered_df = filtered_df[filtered_df[column] >= value]
                elif operator == '<=':
                    filtered_df = filtered_df[filtered_df[column] <= value]
                elif operator == '!=':
                    filtered_df = filtered_df[filtered_df[column] != value]
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if export_format == 'csv':
            # Create CSV file in memory
            output = BytesIO()
            
            # Write CSV data
            filtered_df.to_csv(output, index=False, encoding='utf-8')
            output.seek(0)
            
            filename = f"filtered_data_export_{timestamp}.csv"
            
            return send_file(
                output,
                as_attachment=True,
                download_name=filename,
                mimetype='text/csv'
            )
            
        else:  # Excel format
            # Create Excel file in memory
            output = BytesIO()
            
            # Create Excel writer
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Write filtered data to first sheet
                filtered_df.to_excel(writer, sheet_name='Filtered Data', index=False)
                
                # Create a summary sheet with filter information
                summary_data = {
                    'Export Information': [
                        f'Export Date: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}',
                        f'Total Rows: {len(filtered_df)}',
                        f'Original Rows: {len(df)}',
                        f'Filtered Percentage: {len(filtered_df)/len(df)*100:.1f}%',
                        '',
                        'Applied Filters:'
                    ]
                }
                
                # Add filter details
                if filters:
                    for column, filter_config in filters.items():
                        if filter_config['type'] == 'categorical':
                            summary_data['Export Information'].append(
                                f"{column}: {', '.join(filter_config['values'])}"
                            )
                        elif filter_config['type'] == 'numeric':
                            summary_data['Export Information'].append(
                                f"{column} {filter_config['operator']} {filter_config['value']}"
                            )
                else:
                    summary_data['Export Information'].append('No filters applied')
                
                # Create summary DataFrame
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Export Summary', index=False)
                
                # Auto-adjust column widths
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Prepare file for download
            output.seek(0)
            filename = f"filtered_data_export_{timestamp}.xlsx"
            
            return send_file(
                output,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        
    except Exception as e:
        print(f"Error in export_data: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})

if __name__ == '__main__':
    app.run(debug=True, host='localhost', port=5000)