# First Aid
#### Technologies: ASP, VBScript, HTML, CSS, MS SQL Server
### [All Saints' Catholic Academy](http://www.allsaints.notts.sch.uk) - Built on 05/12/2012

## Index
* [Installation](#Install)
* [Usage](#Usage)
* [Screen Shots](#Shots)

## Challenege
A web-based First-Aid database where uses can keep track of what first-aid was administered and to whom.

## <a name="Install">Installation</a>
* To clone the repo
```shell
$ git clone https://github.com/adrianeyre/first-aid
$ cd first-aid
```

* Set up a web framework such as MS IIS

* Add an ODBC connection to your SQL Server

* Update the file `Connections/PCRoomConnection.asp' with your connection, username and password
```shell
MM_PCRoomConnection_STRING = "dsn=<ODBC Connection>;uid=<USERNAME>;pwd=<PASSWORD>;"
```

## <a name="Shots">Screen Shots</a>
### Default Screen
[![Screenshot](https://raw.githubusercontent.com/adrianeyre/first-aid/master/images/screenshot1.png)](https://raw.githubusercontent.com/adrianeyre/first-aid/master/images/screenshot1.png "Screen Shot 1")
