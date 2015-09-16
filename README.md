# DumpSqlServerDbSchema

A script that, given a connection string to a SQL Server database, 
dumps out its tables and views as a HTML document. There are 
many like it but this one is mine. It understands SQL Server-isms 
like the "MS_Description" extended property (you *do* fill those in, 
right?) An unholy mix of VBScript, SQL and JavaScript/ jQuery.

I usually end up re-creating this script, or something like it, on 
every job to help get my head around the database schema. This is 
an amalgamation of several variations.


## Usage

    DumpSqlServerDbSchema /c connection-string [/t] [/i title] [/rp regexp | /rf file-name ]
    
    connection-string: ADO connection string to connect to database (required)
    /t: include table of contents (optional)
    /t: set the title of the HTML page (optional)
    regexp: regular expression that matches the names of tables containing reference data (optional)
    file-name: file name containing a list of names of tables containing reference data (optional)


## Bugs/ Missing Features

# Extend to include other database objects (stored procedures, etc.)
