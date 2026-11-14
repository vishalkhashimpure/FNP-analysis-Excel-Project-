# Ferns-and-Petals-Excel-Project-
This project focuses on analyzing Ferns N Petals (FNP) sales data using Microsoft Excel to understand sales performance, customer purchasing behaviour, and order trends. The analysis uses Excel functions, Pivot Tables, charts, and an interactive dashboard to identify key business insights and support better decision-making.


Problem Statement: Ferns and Petals Sales Analysis

You have been given a dataset from FNP (Ferns and Petals) that specializes in sending gifts for various occasions like Diwali, Raksha Bandhan, Holi, Valentine's Day, 	Birthdays, and Anniversaries. The dataset contains details about the products, orders, customers, and relevant dates. Your task is to analyze this dataset to uncover key insights 	related to sales trends, customer behavior, and product performance.

Create a dashboard and answer the following key business questions to help the company
improve its sales strategy and optimize customer satisfaction:
1. Total Revenue: Identify the overall revenue.
2. Average Order and Delivery Time: Evaluate the time taken for orders to be delivered.
3. Monthly Sales Performance: Examine how sales fluctuate across the months of 2023.
4. Top Products by Revenue: Determine which products are the top revenue generators.
5. Customer Spending Analysis: Understand how much customers are spending on average.
6. Sales Performance by Top 5 Product: Track the sales performance of top 5 products.
7. Top 10 Cities by Number of Orders: Find out which cities are placing the highest number of orders.
8. Order Quantity vs. Delivery Time: Analyze if higher order quantities impact delivery times.
9. Revenue Comparison Between Occasions: Compare revenue generated across different occasions.
10. Product Popularity by Occasion: Identify which products are most popular during specific occasions.

Project Workflow – FNP Sales Analysis Using Microsoft Excel

Step 1: Data Extraction Using Power Query
The first step is to import the datasets into Microsoft Excel using Power Query.
1. Open Excel and navigate to the Data tab.
2. Select Get Data → From File → From Folder.
3. Browse and select the folder containing the project datasets.
4. Click Combine & Transform Data to load the data into Power Query.
5. The project contains three datasets: Customers, Orders, and Products.
6. Review the datasets to understand their structure, columns, and business information.
7. Understand the business problem statement and the questions that need to be answered through the analysis.
Step 2: Data Cleaning Using Power Query
The next step is to examine the datasets and identify data quality issues.
1. Check all three datasets for missing or null values.
2. Verify that the required columns contain complete information.
3. Check for duplicate records and ensure that unique identifiers, such as Customer ID and Product ID, are properly maintained.
4. Check the data types of each column, including dates, numbers, and text fields.
5. In this project, no null values were identified during the initial data quality check.
Step 3: Data Transformation Using Power Query
Data transformation involves modifying the existing datasets and creating additional columns required for analysis.

1. Customer Phone Number

The Customers table contains a phone number column with some invalid or incorrectly formatted numbers. These numbers will be validated and corrected during the later data processing stage.

2. Extract Month Names

In the Orders table, extract the month names from the Order Date and Delivery Date columns.

Steps:
- Select the Order Date column.
- Navigate to the Add Column tab.
- Select Date → Month → Name of Month.
- Repeat the same process for the Delivery Date column.
This creates separate columns containing the order month and delivery month.

3. Extract Order and Delivery Hours

Extract the hour from the Order Time and Delivery Time columns.

Steps:
- Select the Order Time column.
- Navigate to Add Column → Time → Hour.
- Repeat the same process for the Delivery Time column.
This creates two additional columns that can be used to analyze order and delivery timings.

4. Calculate Delivery Duration

Calculate the number of days between the order date and delivery date.

Steps:
- Navigate to the Add Column tab.
- Select Custom Column.
- Enter the expression:
[Delivery_Date] - [Order_Date]
- The result will be displayed as a duration.
- Use the Duration options to extract the number of days.
This column helps analyze the average time taken to deliver customer orders.

5. Extract Unit Price Using Merge Queries

To calculate revenue, the Unit Price column is required in the Orders table.

Steps:
- Navigate to the Home tab in Power Query.
- Select Merge Queries.
- Select the Orders table and the relevant table containing Unit Price.
- Choose the common column required to match the records.
- Select Left Outer Join.
- Expand the merged table and select only the Unit Price column.
- Change the Unit Price data type to Currency or Decimal Number.
6. Load the Transformed Data

After completing the transformations:
- Navigate to the Home tab.
- Select Close & Load To.
- Load the datasets into Excel as tables.
- Make the datasets available for further data modelling.
Step 4: Phone Number Validation Using Excel Formulas
After loading the datasets into Excel, validate the customer phone numbers.

1. Convert Phone Numbers into Text

Use the TEXT function to convert phone numbers into a consistent text format.

Example:

=TEXT(A2,"0")

2. Validate Phone Numbers

Use the IF and AND functions to check whether the phone numbers meet the required format.

For Indian mobile numbers, a valid number should contain exactly 10 digits and start with 6, 7, 8, or 9.

The validation formula can be:

=IF(AND(LEN(B2)=10,ISNUMBER(--B2),LEFT(B2,1)>="6",LEFT(B2,1)<="9"),"Valid","Invalid")

This helps identify valid and invalid phone numbers.
Step 5: Data Modelling Using Power Pivot
The next step is to add all three datasets to the Excel Data Model.

Steps:
1. Select the Customers table.
2. Navigate to the Power Pivot tab.
3. Select Add to Data Model.
4. Repeat the same process for the Orders and Products tables.
5. Verify that all three tables are available in the Data Model.
This allows us to establish relationships between the datasets and perform analysis using multiple tables.
Step 6: Establish Relationships Between Tables
Before creating Pivot Tables, establish relationships between the three datasets.

1. Enable Power Pivot

If the Power Pivot tab is not visible:
- Go to File → Options.
- Select Add-ins.
- At the bottom, select COM Add-ins and click Go.
- Enable Microsoft Power Pivot for Excel.
- Click OK.
The Developer tab is not required for this process.

2. Create Relationships
1. Navigate to Data → Manage Data Model.
2. Open the Power Pivot window.
3. Switch to Diagram View.
4. Establish relationships between the Customers, Orders, and Products tables using their common identifiers.
5. Organize the tables using a Star Schema approach.
The Orders table acts as the central fact table, while Customers and Products act as dimension tables.

Important: During the initial Pivot Table analysis, the total values for both genders appeared identical. This indicated that the data model or table relationships needed to be checked. Verify the relationships and aggregation methods before proceeding with the analysis.
Step 7: Create Calculated Columns and Measures Using DAX
After establishing relationships, create calculated columns and measures in the Data Model.

1. Calculate Revenue

Create a calculated column named Revenue using the Quantity and Unit Price columns.

DAX formula:

Revenue = Orders[Quantity] * Orders[Unit Price]

This calculates the revenue generated by each order item.

2. Extract Order Day

Create a calculated column to identify the day of the week for each order.

DAX formula:

Order Day = FORMAT(Orders[Order_Date], "dddd")

This returns the weekday name, such as Monday, Tuesday, or Wednesday.

3. Create Measures

Create the required measures for calculating business metrics, such as total revenue and order count.

These calculated values will be used in Pivot Tables and the final dashboard.
Step 8: Data Analysis Using Pivot Tables
After preparing the Data Model, use Pivot Tables to answer the business questions provided in the project.

1. Total Revenue

Identify the overall revenue generated by the business.
- Drag the Revenue column into the Values area.
- Use the SUM aggregation to calculate total revenue.
2. Average Order and Delivery Time

Evaluate the average time taken to deliver customer orders.
- Drag the delivery duration column into Values.
- Open Value Field Settings.
- Change the aggregation from Sum to Average.
3. Monthly Sales Performance

Analyze how sales fluctuate across different months of 2023.
- Drag Month Name into Rows.
- Drag Revenue into Values.
- Arrange the months in chronological order using the month number.
4. Top Products by Revenue

Identify the products generating the highest revenue.
- Drag Product Name into Rows.
- Drag Revenue into Values.
- Apply Value Filters → Top 10.
- Sort the results by revenue.
5. Customer Spending Analysis

Understand the average amount spent per order.
- Drag Revenue into Values.
- Change the aggregation from Sum to Average.
- Rename the result as Average Revenue per Order.
6. Sales Performance by Product Category

Compare revenue generated by different product categories.
- Drag Category into Rows.
- Drag Revenue into Values.
- Sort the categories based on revenue.
7. Top 10 Cities by Number of Orders

Identify the cities generating the highest number of orders.
- Drag City into Rows.
- Drag Order ID into Values.
- Change the aggregation to Count.
- Apply Value Filters → Top 10.
8. Order Quantity vs. Delivery Time

Analyze whether larger order quantities are associated with longer delivery times.
- Use the Quantity and Delivery Duration columns.
- Apply the CORREL function to calculate the correlation.
Example:

=CORREL(Quantity_Range,Delivery_Duration_Range)

9. Revenue Comparison Between Occasions

Compare revenue generated across different occasions.
- Drag Occasion into Rows.
- Drag Revenue into Values.
- Analyze the revenue contribution of each occasion.
10. Product Popularity by Occasion

Identify which products are popular during specific occasions.
- Create a Pivot Table with Product Name and Revenue.
- Insert a slicer using the Occasion field.
- Use the slicer to filter product performance by occasion.
Step 9: Interactive Dashboard Development
The final step is to create an interactive executive dashboard that presents the important business metrics and analysis visually.

1. Prepare Additional Measures and Analysis

Before creating the dashboard:
- Create an additional analysis for Revenue by Order Hour.
- Create a measure for the total number of orders using Order ID.
2. Create Charts

Use Pivot Charts to visualize the analysis.

The dashboard contains six main visualizations:
1. Revenue by Occasion: Column Chart to compare revenue across different occasions.
2. Revenue by Category: Bar Chart to identify the highest revenue-generating product categories.
3. Revenue by Product: Column Chart to compare revenue generated by individual products.
4. Monthly Revenue: Line Chart to visualize monthly sales trends.
5. Orders by Top 10 Cities: Column Chart to compare the number of orders across cities.
6. Revenue by Order Hour: Line Chart to understand revenue distribution throughout the day.
3. Add Interactive Filters
- Insert slicers to filter the dashboard based on Occasion and other relevant fields.
- Insert Timelines for Order Date and Delivery Date.
- Connect the slicers and timelines to all relevant Pivot Tables and Pivot Charts using Report Connections.
4. Create KPI Cards

Add important business metrics at the top of the dashboard.

Examples:
- Total Orders
- Total Revenue
- Average Order and Delivery Time
- Average Revenue per Order
Steps:
1. Go to the Insert tab and select Shapes.
2. Choose suitable shapes for the KPI cards.
3. Link the KPI values to the corresponding cells in the summary sheet using cell references.
Example:

=Sheet1!B2

5. Dashboard Formatting

Complete the dashboard design by:
- Adding a suitable dashboard title.
- Aligning and resizing charts and KPI cards.
- Hiding unnecessary Pivot Chart field buttons.
- Removing unnecessary legends.
- Adding clear chart titles.
- Applying a consistent dashboard colour theme.
- Adjusting the layout to make the dashboard easy to understand.
The final dashboard provides an interactive overview of FNP's sales performance and allows users to explore the data using slicers and timelines.
Project Outcome
This project demonstrates an end-to-end sales analysis workflow using Microsoft Excel and its advanced data analysis features.

It covers data extraction, data cleaning, data transformation, data modelling, DAX calculations, Pivot Table analysis, and interactive dashboard development.

The final dashboard helps understand sales performance, customer spending, product popularity, occasion-wise revenue, monthly trends, and order delivery patterns. These insights can support better business decisions and improve sales planning.


Key Insights


1. Overall Sales Performance: The business recorded 1,000 orders and generated total revenue of ₹35,20,984.
2. Average Customer Spending: The average revenue per order is approximately ₹3,520.98.
3. Order Delivery: The average time taken between order placement and delivery is 5.53 days.
4. Revenue by Occasion: Anniversary contributed the highest revenue, followed by Raksha Bandhan, All Occasions, and Holi. Diwali and Valentine's Day generated comparatively lower revenue.
5. Revenue by Category: Colors was the highest revenue-generating category, followed by Soft Toys and Sweets. Mugs and Plants contributed comparatively less.
6. Monthly Revenue Trends: August recorded the highest revenue, followed by February. January, April, and September showed comparatively lower revenue.
7. City-wise Order Analysis: Imphal, Dhanbad, and Kavali were among the leading cities based on the number of orders.
8. Revenue by Order Time: Revenue was relatively higher during the evening hours, particularly around 7 PM and 8 PM.

Conclusion

The FNP Sales Analysis project provides an overview of sales performance, revenue distribution, customer spending, and order trends. The analysis highlights the occasions, product categories, months, cities, and order timings that contribute to overall revenue.

The insights can help the business identify high-performing categories, understand seasonal sales patterns, and plan targeted marketing campaigns. Focusing on popular occasions, high-revenue products, and peak purchasing hours could help improve sales performance and business decisions.
Project Outcome

Through this project, I gained practical experience in Microsoft Excel, including data cleaning, data processing, Pivot Tables, data analysis, chart creation, and interactive dashboard development. It also helped me understand how to convert raw sales data into meaningful business insights.
