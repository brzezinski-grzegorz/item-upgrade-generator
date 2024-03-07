
# Upgrade Item Script Creator   

This is a simple tool for Knight Online v1298 server admins, that allows quickly and without a lot of work create Item Upgrade query for database. 

## Screenshots

![App Screenshot](https://ko4life.net/uploads/monthly_2020_04/1865344740_Item0.png.7f9676eeb374e8322f6098877ca9794e.png)
![App Screenshot](https://ko4life.net/uploads/monthly_2020_04/521803348_Item1.thumb.png.76816bf1196005b0e9467b54e675a42b.png)
![App Screenshot](https://ko4life.net/uploads/monthly_2020_04/2126444447_Item2.png.807d2411baa5d84ed60c15d6b5895be3.png)

## Deployment

So first of all you will have to input your connection data, it's pretty much self explanatory, just the server name can be confusing for some so let me get it out of the way and say that it's actually a Server Instance Name normally when you install MSSQL you either have YOURCOMPUTERNAME\SQLEXPRESS or YOURCOMPUTERNAME\KO its the same you have put in the ODBC settings.

Now that you are connected to the database you see the main form. In the Item Name window you can input name of item and hit enter, and it will search your database for that item if it's in the database then it will show the name and corresponding item number. From that item number you can easily fill the nOriginItem column.

In Item Group there are pre-defined by me profiles that will put some of the values for you like required item, required money & upgrade rate. Also the Index is taken from your ITEM_UPGRADE it's set by taking the last number and adding 100 to it so the numbers won't collide.

And last but not least there are two magic buttons - also pretty self explanatory. The first one save the query to ITEM_UPGRADE table and closes the application, another creates *.sql script in the application directory.

