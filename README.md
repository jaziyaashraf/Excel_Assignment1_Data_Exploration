# Excel_Assignment1_Data_Exploration
Perform the following in the dataset from the 'Dataset' sheet.		
		
	1) Sum, Count, Average:	
		• What is the total price of all products in the dataset?
		• How many products are there in the dataset?
		• Calculate the average price of the products.
		
	2) Min and Max:	
		• Determine the minimum price among all products.
		• Find the maximum price among all products.
		
	3) IF Function:	
		• Using an IF function, create a new column named Price Range to categorize products with a price greater than or equal to $500 as 'High Price' and others as 'Standard Price'.
		
	4) SUMIF and COUNTIF:	
		• Calculate the total price for products in the 'Electronics' category using the SUMIF function.
		• Determine the count of products with a price less than $100 using the COUNTIF function.
		
	5) Text Formatting - LEFT, RIGHT, MID:	
		• Create a new column named Day with the first 2 characters of each 'Product ID' using the LEFT function.
		• Create a new column named Country Code by extracting the last 2 characters from the 'Product ID' column using the RIGHT function.
		• Create a new column named Month by extracting 4th to 6th characters from the 'Product ID' column using the MID function.


This assignment is about analyzing product data in Excel. Use different formulas to find the total, count, average, minimum, and maximum prices, categorize products based on price, and extract information from Product IDs.

SUM, COUNT, AVERAGE
SUM → =SUM(range)
COUNT → =COUNT(range)
AVERAGE → =AVERAGE(range)


MIN, MAX
MIN → =MIN(range)
MAX → =MAX(range)

IF
=IF(price>=500,"High Price","Standard Price")

SUMIF, COUNTIF
SUMIF → =SUMIF(category_range,"Electronics",price_range)
COUNTIF → =COUNTIF(price_range,"<100")

LEFT, RIGHT, MID
LEFT → =LEFT(Product_ID,2)
RIGHT → =RIGHT(Product_ID,2)
MID → =MID(Product_ID,4,3)
