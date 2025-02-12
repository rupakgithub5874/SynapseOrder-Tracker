SynapseOrder-Tracker  
Overview  
SynapseOrder-Tracker is an interactive Excel-based dashboard designed to provide a comprehensive view of customer transactions. This project enables users to select a customer name and instantly retrieve detailed order-related data, including contact details, transaction history, and key metrics. The dashboard is equipped with PivotTables, charts, and slicers to enhance data visualization and filtering.

Features and Functionalities
1. Customer Selection & Automatic Data Retrieval
Users can select a customer name from a dropdown list (enabled via a slicer).  
Upon selection, the dashboard automatically fetches and displays key customer details:  
  a.Contact Information: Phone number, fax number, and address.  
  b.Geographical Details: City, postal code, and country.  
Order Summary:  
  a.Total Order Amount: The sum of all orders placed by the customer.  
  b.Order Count: The total number of orders made.  
  c.Average Freight Cost: The mean freight cost for the customer's shipments.  


2. Embedded Slicer for Easy Filtering  
The slicer allows users to quickly filter data based on customer selection.  
Filtering updates all related PivotTables and charts dynamically, ensuring real-time insights.  


3. Customer Order History Table  
Displays a detailed order history for the selected customer, including:  
   a.Customer ID: Unique identifier for the customer.  
   b.Order ID: Unique order identifier.  
   c.Order Date: Date when the order was placed.  
   d.Required Date: Date by which the order is required.  
   e.Shipment Date: The actual shipment date.  
   f.Shipper Name: The logistics provider handling the delivery.  
   g.Order Amount: The total value of each order.  
   h.Customer Name: The name of the selected customer.  
   i.Month: The month in which the order was placed (useful for trend analysis).  


4. Interactive Charts for Data Visualization  
 a. A visually appealing chart is embedded in the dashboard to represent order trends.  
 b. The chart updates dynamically based on customer selection, providing insights into order volume and financial metrics over time.  


5. Automated Data Filtering Using VBA  
a.The dashboard uses VBA automation to apply filters and update the displayed data.  
b.The script ensures that:  
      The selected customer’s details are applied to the relevant PivotTable fields.  
      The order history and charts update automatically to reflect the selection.  
      If a customer does not exist, a message box alerts the user and hides irrelevant elements.  


6.How It Works  
1.Open the Excel file containing the SynapseOrder-Tracker dashboard.  
2.Use the slicer to select a customer name.  
3.The dashboard will automatically:  
    a.Display the customer’s contact details and location.  
    b.Show total orders, order count, and average freight cost.  
    c.Update the order history table with transaction details.  
    d.Refresh the chart to visualize the customer’s order trends.  
4.If an invalid customer name is selected, a warning message appears, and the chart and slicer are hidden.  


7.Technologies & Tools Used  

  a.Microsoft Excel – Data storage and visualization.  
  b.PivotTables & Slicers – Dynamic data filtering.  
  c.Charts – Graphical representation of customer order trends.  
  d.VBA (Visual Basic for Applications) – Automating data filtering and error handling.  
