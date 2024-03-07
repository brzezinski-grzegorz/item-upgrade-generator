
# Upgrade Item Script Creator   

This is a simple tool for Knight Online v1298 server admins, that allows quickly and without a lot of work create Item Upgrade query for database. 

## Screenshots

![1](https://github.com/brzezinski-grzegorz/item-upgrade-generator/assets/63126985/f949dd66-fc60-4493-adfb-9f700dd7a8be)
![2](https://github.com/brzezinski-grzegorz/item-upgrade-generator/assets/63126985/3c3712cd-bbaa-4dae-97c0-59d81092fcf7)
![3](https://github.com/brzezinski-grzegorz/item-upgrade-generator/assets/63126985/d039b1c2-a794-4320-bd39-8a2ca7ef4c41)

## Deployment

So first of all you will have to input your connection data, it's pretty much self explanatory, just the server name can be confusing for some so let me get it out of the way and say that it's actually a Server Instance Name normally when you install MSSQL you either have YOURCOMPUTERNAME\SQLEXPRESS or YOURCOMPUTERNAME\KO its the same you have put in the ODBC settings.

Now that you are connected to the database you see the main form. In the Item Name window you can input name of item and hit enter, and it will search your database for that item if it's in the database then it will show the name and corresponding item number. From that item number you can easily fill the nOriginItem column.

In Item Group there are pre-defined by me profiles that will put some of the values for you like required item, required money & upgrade rate. Also the Index is taken from your ITEM_UPGRADE it's set by taking the last number and adding 100 to it so the numbers won't collide.

And last but not least there are two magic buttons - also pretty self explanatory. The first one save the query to ITEM_UPGRADE table and closes the application, another creates *.sql script in the application directory.

