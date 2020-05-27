# xlsx to sql

Simple


#### Usage

##### xlsx table requirements
Data in xlsx file should be stored in specific way

- The script iterates through all sheet 

- Data in sheets must be stored as follows
<table>
 <tr><td><i>MySQL_datatype_of_column_1</i> </td> 
    <td><i>MySQL_datatype_of_column_2</i> 
    <td><i>MySQL_datatype_of_column_3</i></td></tr>
 <tr><td><i>column_name_1</i> </td> 
    <td><i>column_name_2</i> 
    <td><i>column_name_3</i></td></tr>
 <tr><td><i>column_data_1</i> </td> 
    <td><i>column_data_2</i> 
    <td><i>column_data_3</i></td></tr>
</table>

<img src="https://i.ibb.co/0y6LbHY/exampl.png" alt="exampl img" ></a>

##### command
<pre>
python xlsx_to_sql -f <i>path_to_xlsx_file</i> -d <i>path_to_destination_sql_file</i>
</pre>

#### Requirements

- python 3.x
