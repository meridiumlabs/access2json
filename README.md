# access2json

This is a simple console application that can dump data from a Microsoft Access database to [JSON](http://www.json.org/). It can also display some metadata about the tables and columns in the database. 

## Usage

The application supports two verbs, `info` and `json`. You can also use the verb `help` to list the supported options.

Two options are common for both the `info` and the `json` verbs

| option | description |
|--------|-------------|
| `--database` | the path to Access database file |
| `--password` | password of the database, if it is protected |

### info

The info verb displays metadata about the tables in the database. 

**Options**

| option | description |
|--------|-------------|
| `-t, --table` | display info about a specific table | 

**Examples**

Dump meta data about all tables in the database:

```batch
access2json info --database phonebook.mdb
```

Display metadata about the **Address** table:

```batch
access2json info --database phonebook.mdb --table Address
```

### json

Generates JSON from the data in the database. There are a few options you can use to customize the output

**Options**

| option | description |
|--------|-------------|
| `-t, --tables` | Dump data from these tables. Separate names with a comma. Default is to dump all tables. | 
| `-o, --out-file` | A file path. Write the resulting JSON to this file. The default is to write output to standard out. |
| `--pretty` | Indent the JSON output nice and pretty. The default is one line |
| `--normalize` | Try to make property names usable via dot notation by removing diacritics and replacing non-identifier characters with underscores. |
| `--force-numbers` | If a string value from the database looks like a number, cast it to a number in the output |
| `--keep-text`| Comma separated property names that will be kept as text no matter what |

**Examples**

Write data from the **Address** table to, nicely indented, with normalized property names and all numberlike strings cast to numbers to a file named address.json

```batch
access2json json --database phonebook.mdb -t Address --pretty --normalize --force-numbers > address.json
```

Dump all data of the protected database to the file **dump.js**

```batch
access2json json --database phonebook.mdb --password 123456 --out-file c:\users\john\dump.js  
```

## Prerequisites
You'll need the Microsoft Access Database Engine installed. This can be downloaded here:
[https://www.microsoft.com/en-us/download/details.aspx?id=13255](https://www.microsoft.com/en-us/download/details.aspx?id=13255)

If you get an error when running the application that says something like **The 'microsoft.ace.oledb.12.0' provider is not registered on the local machine**, you probably don't have an Access db engine installed. Follow [these instructions](https://social.msdn.microsoft.com/Forums/en-US/1d5c04c7-157f-4955-a14b-41d912d50a64/how-to-fix-error-the-microsoftaceoledb120-provider-is-not-registered-on-the-local-machine?forum=vstsdb) to get that stuff installed.
