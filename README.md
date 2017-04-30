# smartsheet-api-dev
Custom scripts pertaining to the Smartsheet API using requests, smartsheet-python-sdk, pyodbc, et al.

The scripts in this repo are primarily freelance work that I have performed for clients. It includes:

ss_to_mssql.py
- This is a dynamic script developed for a pc user who requested python2.
- The script allows the user to input by name which sheet or report to be uploaded into a MSSQL database.
- Your SQL schema and insert statements are generated dynamically based on your selected sheet/report.
- The trickiest part of the pyodbc module module is getting the connection string right. It is platform dependent.
    If you are a mac/linux user, check out FreeTDS as a driver.
    See here: https://github.com/mkleehammer/pyodbc/wiki/Connecting-to-SQL-Server-from-Windows
    Also: https://tryolabs.com/blog/2012/06/25/connecting-sql-server-database-python-under-ubuntu/
    
csv_ss_update.py
- Another dynamic script for a client who wanted to automate updates from a 56,000 row csv file to different sheets.
- Script finds matches based on uniquely-identifiable value (SKU), updates two values (QOH and Priority), and finally
    stamps the row with the date at which the row was updated.
    
ss_custom_func.py
- Some basic interactions with the Smartsheet API using the smartsheet-python-sdk.
- This is a nice file to look over if you are new the the API. 

archiving.py
- In development; a script that will automate the archival process for a client based on cell values. 
