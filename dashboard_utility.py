import streamlit as st
import pandas as pd
import plotly.express as px
import datetime

# --- Configuration and Setup (Step 1 & 6 part) ---
st.set_page_config(layout="wide")
st.title("📊 Sales MIS Reporting Dashboard Utility")
st.markdown("Upload your Excel sales data to generate an interactive dashboard.")

uploaded_file = st.file_uploader("Choose an Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    # ------------------ Step 2: Read Excel File ------------------
    try:
        # Read the Excel file into a Pandas DataFrame
        df = pd.read_excel(uploaded_file)
        st.success("File uploaded and read successfully!")
        
        # Display the first few rows for confirmation
        st.subheader("Raw Data Preview")
        st.dataframe(df.head())

        # ------------------ Step 4: Handle DATEVALUE Conversion ------------------
        # Assuming the date column is named 'Date' or similar. Adjust as needed.
        # This function converts Excel serial date numbers (like 41760) to datetime objects.
        
        def excel_serial_to_date(serial_number):
            """Converts Excel serial date (e.g., 41760) to datetime object."""
            # Excel base date is Dec 30, 1899 (for Windows)
            if pd.isna(serial_number):
                return None
            try:
                # Check if the number is within a reasonable range for serial dates
                if serial_number > 1000 and serial_number < 90000:
                    return pd.to_datetime('1899-12-30') + pd.to_timedelta(serial_number, unit='D')
                else:
                    return serial_number # Return original if it doesn't look like a serial date
            except Exception:
                return serial_number # Return original on error

        # Example: Apply the conversion to a likely date column (You MUST adjust the column name)
        # **Replace 'Order_Date' with the actual name of your date column in Excel.**
        DATE_COLUMN = 'Order_Date' 
        
        if DATE_COLUMN in df.columns:
            st.info(f"Checking and converting Excel DATEVALUE in the column: **{DATE_COLUMN}**")
            
            # Use a helper column to try and apply the conversion
            df['Corrected_Date'] = df[DATE_COLUMN].apply(excel_serial_to_date)
            
            # Drop the original date column and rename the corrected one
            df.drop(columns=[DATE_COLUMN], inplace=True)
            df.rename(columns={'Corrected_Date': DATE_COLUMN}, inplace=True)
            
            # Ensure the column is properly cast to datetime for plotting
            df[DATE_COLUMN] = pd.to_datetime(df[DATE_COLUMN], errors='coerce')
            df.dropna(subset=[DATE_COLUMN], inplace=True) # Remove rows where date conversion failed

        # ------------------ Step 3: Analyze Data and Create Visuals ------------------
        st.header("📈 Key Performance Indicators (KPIs) & Trends")
        
        # **Data Cleaning & Analysis** (Adjust 'Sales_Amount' and 'Product_Category' names)
        # You MUST replace 'Sales_Amount' and 'Product_Category' with your actual column names.
        SALES_COLUMN = 'Sales_Amount'
        CATEGORY_COLUMN = 'Product_Category'
        
        if SALES_COLUMN in df.columns:
            total_sales = df[SALES_COLUMN].sum()
            avg_sales = df[SALES_COLUMN].mean()
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Sales", f"${total_sales:,.2f}")
            with col2:
                st.metric("Average Sale Value", f"${avg_sales:,.2f}")
            with col3:
                st.metric("Total Transactions", f"{len(df):,}")

            # --- Visual 1: Sales Trend Over Time (Line Chart) ---
            if DATE_COLUMN in df.columns:
                sales_trend = df.set_index(DATE_COLUMN)[SALES_COLUMN].resample('M').sum().reset_index()
                fig_line = px.line(
                    sales_trend,
                    x=DATE_COLUMN,
                    y=SALES_COLUMN,
                    title="Monthly Sales Trend",
                    labels={SALES_COLUMN: "Total Sales ($)", DATE_COLUMN: "Date"}
                )
                st.plotly_chart(fig_line, use_container_width=True)
        
            # --- Visual 2: Sales by Category (Bar Chart) ---
            if CATEGORY_COLUMN in df.columns:
                category_sales = df.groupby(CATEGORY_COLUMN)[SALES_COLUMN].sum().reset_index().sort_values(by=SALES_COLUMN, ascending=False)
                fig_bar = px.bar(
                    category_sales,
                    x=CATEGORY_COLUMN,
                    y=SALES_COLUMN,
                    title="Sales by Product Category",
                    color=SALES_COLUMN,
                    labels={SALES_COLUMN: "Total Sales ($)", CATEGORY_COLUMN: "Category"}
                )
                st.plotly_chart(fig_bar, use_container_width=True)

            # --- Visual 3: Sales Distribution (Pie Chart) ---
            if CATEGORY_COLUMN in df.columns:
                fig_pie = px.pie(
                    category_sales,
                    values=SALES_COLUMN,
                    names=CATEGORY_COLUMN,
                    title="Sales Share by Category",
                    hole=.3 # Creates a donut chart
                )
                st.plotly_chart(fig_pie, use_container_width=True)

        else:
             st.error(f"Required column '{SALES_COLUMN}' not found. Please check column names.")

    except Exception as e:
        st.error(f"An error occurred during file processing: {e}")

# ------------------ Step 5 & 6: Running the Utility ------------------
# To run this utility, save the code as `dashboard_utility.py` and execute the following command in your terminal:
# streamlit run dashboard_utility.py