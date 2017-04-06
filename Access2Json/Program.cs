using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using CommandLine;
using Newtonsoft.Json;

namespace Access2Json {
    class Program {
        static int Main(string[] args) {
            try {
                return Parser.Default.ParseArguments<InfoOptions, JsonOptions>(args)
                    .MapResult(
                        (InfoOptions options) => DisplayInfo(options),
                        (JsonOptions options) => GenerateJson(options),
                        error => 1);
            } catch(Exception e) {
                Console.Error.WriteLine("An exception went unhandled [{0}]: {1}", 
                    e.GetType().Name, e.Message);
                return 1;
            }
        }

        static int GenerateJson(JsonOptions options) {
            var tables = GetTables(options);

            using (var writer = GetTextWriter(options))
            using (var json = new JsonTextWriter(writer)) {
                if (options.PrettyPrint) {
                    json.Formatting = Formatting.Indented;
                    json.Indentation = 2;
                    json.IndentChar = ' ';
                }
                json.WriteStartObject();
                foreach (var table in tables) {
                    var normalizedTableName = options.Normalized(table.Name);
                    json.WritePropertyName(normalizedTableName);
                    json.WriteStartArray();
                    foreach (var row in table.Rows) {
                        json.WriteStartObject();
                        foreach (var field in row.Fields) {
                            var normalizedFieldName = options.Normalized(field.Name);
                            json.WritePropertyName(normalizedFieldName);
                            json.WriteValue(
                                options.KeepList.Any(x => x.Key.Equals(normalizedTableName)
                                                            && x.Value.Equals(normalizedFieldName))
                                    ? field.Value
                                    : options.CastNumber(field.Value)
                                );
                        }
                        json.WriteEndObject();
                    }
                    json.WriteEndArray();
                }
                json.WriteEndObject();
            }
            
            return 0;
        }

        static TextWriter GetTextWriter(JsonOptions options) {
            if (options.OutputFile == null) {
                return Console.Out;
            } else {
                return new StreamWriter(options.OutputFile, false, Encoding.UTF8); 
            }
        }

        static IEnumerable<Table> GetTables(JsonOptions options) {
            var connectionString =
                new ConnectionString(options.AccessFilePath, options.Password);

            using (var db = new DB(connectionString)) {
                var tableNames = options.Tables.Any()
                    ? options.Tables
                    : db.Tables.Select(t => t.Name);

                foreach (var tableName in tableNames) {
                    var rows = db.GetTableData(tableName).ToArray();
                    yield return new Table(tableName, rows);
                }
            }
        }

        static int DisplayInfo(InfoOptions options) {
            var connectionString =
                new ConnectionString(options.AccessFilePath, options.Password);

            using (var db = new DB(connectionString)) {
                var tablesToDisplay = db.Tables
                    .Where(t => options.Table == null || t.Name.Equals(options.Table));

                foreach (var item in tablesToDisplay) {
                    Console.WriteLine("-- {0} --", item.Name);
                    Console.WriteLine(item.Description);
                    foreach (var column in item.Fields) {
                        Console.WriteLine("       Name: [{0}]", column.Name);
                        Console.WriteLine("OLE DB type: [{0}]", column.Type);
                        Console.WriteLine("  .Net Type: [{0}]", TypeMap.FromOleDbType(column.Type));
                        if (column.Description != null) {
                            Console.WriteLine("Description: {0}", column.Description);
                        }
                        Console.WriteLine();
                    }
                    Console.WriteLine();
                }
            }

            return 0;
        }
    }


    [Verb("info", HelpText = "Displays information about the database and tables")]
    class InfoOptions : GeneralOptions {
        [Option('t', "table",   
                HelpText = "The table to display information about. " +
                           "If this parameter is not specified, all tables are included")]
        public string Table { get; set; }
    }

    [Verb("json", HelpText = "Generates json from the specified Access database")]
    class JsonOptions : GeneralOptions {
        [Option('t', "tables", 
                HelpText = "The tables to include in the result. If this parameter " + 
                           "is not specified, all tables are included.",
                Separator = ',')]
        public IEnumerable<string> Tables { get; set; }

        [Option('o', "out-file", 
                HelpText="The name of the JSON file to write output to. " + 
                         "If not specified, output i sent to standard out.")]
        public string OutputFile { get; set; }

        [Option("pretty", HelpText="Pretty print the resulting JSON data" )]
        public bool PrettyPrint { get; set; }

        [Option("normalize", 
                HelpText="Normalize JSON property names by removing " + 
                         "diacritics and replacing non-identifier characters with '_'." )]
        public bool Normalize { get; set; }

        [Option("force-numbers",
                HelpText="Cast all values that look like a number to double")]
        public bool ForceNumbers { get; set; }

        [Option("keep-text", HelpText = "Properties to leave unharmed. Commaseparated string in format [table].[property] (normalized names)")]
        public string KeepText { get; set; }

        public IEnumerable<KeyValuePair<string, string>> KeepList {
            get {
                if (string.IsNullOrEmpty(KeepText))
                    return new List<KeyValuePair<string, string>>();
                var tablesAndProperties = KeepText.Split(',');
                return tablesAndProperties.Select(tp => tp.Split('.'))
                    .Select(data => new KeyValuePair<string, string>(data[0], data[1]))
                    .ToList();
            }
        }

        public object CastNumber(object value) {
            if (!ForceNumbers) return value;

            var strValue = value as string;
            if (strValue == null) return value;

            if (Regex.IsMatch(strValue, @"^-?(:?0|[1-9]\d*)([.,]\d*)?$")){
                return double.Parse(strValue.Replace(',', '.'), CultureInfo.InvariantCulture);
            }

            return value;
        }

        public string Normalized(string text) {
            if (!Normalize || string.IsNullOrWhiteSpace(text))
                return text;

            // Remove diacritics
            text = text.Normalize(NormalizationForm.FormD);
            var chars = text
                .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                .ToArray();
            var noDiacritics = new string(chars).Normalize(NormalizationForm.FormC);

            // Replace all non-identifier chars with an underscore
            var onlyAscii = Regex.Replace(noDiacritics, @"[^_0-9a-z]+", "_", RegexOptions.IgnoreCase);

            // If text starts with a digit, prepend an underscore
            var noNumbersAtStart = Regex.Replace(onlyAscii, @"^(\d)(.*)$", "_$1$2", RegexOptions.IgnoreCase);

            return noNumbersAtStart;
        }

    }

    class GeneralOptions {
        [Option("database", Required = true, HelpText = "Path to the MS Access file")]
        public string AccessFilePath { get; set; }

        [Option("password", HelpText = "The password if the database is protected")]
        public string Password { get; set; }
    }

    class ConnectionString {
        public ConnectionString(string filepath, string password = null) {
            var template = password == null
                ? TemplateWithoutPassword
                : TemplateWithPassword;

            this.connectionString = string.Format(template, filepath, password);
        }

        public static implicit operator string(ConnectionString connectionString) {
            return connectionString.ToString();
        }

        public override string ToString() {
            return this.connectionString;
        }

        private readonly string connectionString;

        private const string TemplateWithoutPassword =
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}; Persist Security Info=False;{1}";
        private const string TemplateWithPassword =
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}; Jet OLEDB:Database Password={1};";
    }

    // GetSchema restrictions
    // https://msdn.microsoft.com/en-us/library/cc716722(v=vs.110).aspx
    static class SchemaRestriction {
        public static readonly string[] UserTables =
            new[] { null, null, null, "TABLE" };

        public static string[] ColumnsOfTable(string tableName) {
            return new[] { null, null, tableName, null };
        }
    }

    // Magic indexes for GetSchema result tables
    struct SchemaInfo {
        public struct Table {
            public const int Name = 2;
            public const int Description = 5;
        }

        public struct Column {
            public const int Name = 3;
            public const int Description = 27;
            public const int Type = 11;
        }
    }

    class DB : IDisposable {
        public DB(ConnectionString connectionString) {
            this.connectionString = connectionString;

            // Getting the 'provider not registered' error?
            // See this: https://social.msdn.microsoft.com/Forums/en-US/1d5c04c7-157f-4955-a14b-41d912d50a64/how-to-fix-error-the-microsoftaceoledb120-provider-is-not-registered-on-the-local-machine?forum=vstsdb
            this.db = new OleDbConnection(this.connectionString);
            this.db.Open();
        }

        public IEnumerable<TableInfo> Tables {
            get {
                var tables = this.db.GetSchema("Tables", SchemaRestriction.UserTables);

                foreach (DataRow table in tables.Rows) {
                    var tableName = table[SchemaInfo.Table.Name].ToString();
                    var description = table[SchemaInfo.Table.Description].ToString();
                    var columns = this.db
                        .GetSchema("Columns", SchemaRestriction.ColumnsOfTable(tableName))
                        .Rows
                        .Cast<DataRow>()
                        .Select( row => new ColumnInfo(
                            row.Field<string>(SchemaInfo.Column.Name), 
                            row.Field<string>(SchemaInfo.Column.Description),
                            row.Field<OleDbType>(SchemaInfo.Column.Type)))
                        .ToArray();

                    yield return new TableInfo(tableName, description, columns);
                }
            }
        }

        public IEnumerable<Row> GetTableData(string tablename) {
            var query = @"SELECT * FROM [" + tablename + "]";
            var command = new OleDbCommand(query, this.db);

            var reader = command.ExecuteReader();

            while (reader.Read()) {
                var fields = new Field[reader.FieldCount];
                for (var i = 0; i < reader.FieldCount; i++) {
                    fields[i] = new Field(
                        reader.GetName(i),
                        reader.GetValue(i),
                        reader.GetFieldType(i)
                    );
                }
                yield return new Row(fields);
            }
        }

        private readonly ConnectionString connectionString;
        private readonly OleDbConnection db;

        public void Dispose() {
            db.Close();
        }
    }

    struct Table {
        public Table(string name, Row[] rows) {
            Name = name;
            Rows = rows;
        }
        public readonly string Name;
        public readonly Row[] Rows;
    }

    struct TableInfo {
        public TableInfo(string name, string description, ColumnInfo[] fields) {
            Name = name;
            Description = description;
            Fields = fields;
        }
        public readonly string Name;
        public readonly string Description;
        public readonly ColumnInfo[] Fields;
    }

    struct ColumnInfo {
        public ColumnInfo(string name, string description, OleDbType type) {
            Name = name;
            Description = description;
            Type = type;
        }
        public readonly string Name;
        public readonly string Description;
        public readonly OleDbType Type ;
    }

    struct Row {
        public Row(params Field[] fields) {
            Fields = fields;
        }
        public readonly Field[] Fields;

        public override string ToString() {
            if (Fields == null || Fields.Length == 0) {
                return "<empty>";
            }
            var output = new StringBuilder();
            foreach (var field in Fields) {
                output.Append(field + " ");
            }
            return output.ToString();
        }
    }

    struct Field {
        public Field(string name, object value, Type type) {
            Name = name;
            Value = value;
            Type = type;
        }

        public readonly string Name;
        public readonly object Value;

        // Access/OLE DB data types
        // https://support.microsoft.com/en-us/kb/320435
        public readonly Type Type;

        public override string ToString() {
            return string.Format(
                @"[{0}] {1}: {2}",
                Type, Name, Value ?? "<null>");
        }
    }

    class TypeMap {
        public static Type FromOleDbType(OleDbType value) {
            return mappings[value];
        }

        private static readonly Dictionary<OleDbType, Type> mappings =
            new Dictionary<OleDbType, Type> {
                {OleDbType.WChar, typeof(System.String) },
                {OleDbType.VarWChar, typeof(System.String) },
                {OleDbType.LongVarWChar, typeof(System.String) },
                {OleDbType.UnsignedTinyInt, typeof(System.Byte) },
                {OleDbType.Boolean, typeof(System.Boolean) },
                {OleDbType.Date, typeof(System.DateTime) },
                {OleDbType.Numeric, typeof(System.Decimal) },
                {OleDbType.Double, typeof(System.Double) },
                {OleDbType.Guid, typeof(System.Guid) },
                {OleDbType.Integer, typeof(System.Int32) },
                {OleDbType.Single, typeof(System.Single) },
                {OleDbType.SmallInt, typeof(System.Int16) },
                {OleDbType.Binary, typeof(System.Byte[]) },
                {OleDbType.LongVarBinary, typeof(System.Byte[]) }
            };
    }
}
