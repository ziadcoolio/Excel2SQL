# Excel2SQL
This script converts tables created in excel into SQL queries for easoy conversion of excel to SQL. 

## Getting Started

You'll need Python3 and OpenPyxl

### Prerequisites

You can install OpenPyxl with pip like this

```
pip install openpyxl --user
```
The excel file has to be formatted correctly, please look at the included demo excel file for proper operation.

### Installing

To start run the following commands.

```
python3 Excel2SQL [Inputfile.xlsx] [Outputfile.sql]
```
Replace [Inputfile.xlsx] with the included "data.xlsx" file (Outputfile is the name of the output file make sure type is sql)
Like so
```
python3 Excel2SQL data.xlsx insertions.sql
```
If the code runs successfully you should have an insertions.sql in your directory


## Built With

* [Python3](https://www.python.org/) - The Language Used
* [OpenPyxl](https://openpyxl.readthedocs.io/en/stable/) - The Library Used


## Authors

* **Mohammad Taha Bin Firoz** - [Taha Firoz](https://github.com/Taha-Firoz)
* **Mohammad Ziad Siddiqui** - [Ziad Coolio](https://github.com/ziadcoolio)

See also the list of [contributors](https://github.com/your/project/contributors) who participated in this project.

## License

This project is licensed under the GPL V3 License - see the [LICENSE.md](LICENSE.md) file for details

