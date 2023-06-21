# XLAPI

One of the common problems that many people face is how to efficiently store and manage data collected from online forms. For example, suppose you have a Google Forms or any other type of form from which the answers will be received in an Excel file. Then, you require a system to generate a relational database to keep an institutional or personal record by capturing the records from that Excel file. The system would have to take the Excel file and create an entity for each record, which can then be inserted into a table in the database.

However, this process can be tedious and error-prone if done manually. You would have to open the Excel file, read each row and column, and write SQL queries to insert the data into the database. Moreover, you would have to deal with different formats and types of data, such as dates, numbers, text, etc.

With an API for Excel files, you can programmatically read and write data from Excel files. You can also perform calculations, generate reports, analyze results, and manipulate the content of Excel files in various ways. This way, you can automate the process of creating a relational database from Excel files and save time and effort.

XLAPI is a Java library that provides a simple and efficient way to read and write xlsx files. It allows you to annotate your POJO classes with annotations to map them to the corresponding cells and rows in the spreadsheet. XLAPI supports multiple serialization modes for the same POJO, so you can easily switch between different serialization results for the same spreadsheet. XLAPI uses w3c.org.dom API internally to manipulate the xml files within the xlsx archive. XLAPI is still in development and may have some bugs or limitations. Please feel free to report any issues or suggestions on GitHub.
