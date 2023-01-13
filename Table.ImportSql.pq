let
    Main = (ConfigTableName as text) =>
        let
            Sql.WrapVarchar = (Text as text) => "'" & Text & "'",
            Sql.WrapNVarchar = (Text as text) => "N'" & Text & "'",
            Sql.ConvertByType = (_value as any) =>
                let
                    _type = Value.Type(_value),
                    _convert =
                        if _type = type logical then
                            Text.From(Number.From(_value))
                        else if _value = "" then
                            "NULL"
                        else if _value = null then
                            "NULL"
                        else if _type = type date then
                            Sql.WrapVarchar(Date.ToText(_value, [Format = "yyyyMMdd"]))
                        else if _type = type datetime then
                            Sql.WrapVarchar(DateTime.ToText(_value, [Format = "yyyyMMdd"]))
                        else if _type = type text then
                            Sql.WrapNVarchar(Text.Replace(_value, "'", "''"))
                        else
                            Sql.WrapVarchar(Text.From(_value))
                in
                    _convert,
            SqlQuery.Insert = (PowerQueryTable as table, NewColumnName as text, SqlTable as text) =>
                let
                    RowCount = Table.RowCount(PowerQueryTable),
                    #"Added Index" = Table.AddIndexColumn(PowerQueryTable, "Index", 1, 1, Int64.Type),
                    AddValuesRows = Table.AddColumn(
                        #"Added Index",
                        NewColumnName,
                        each
                            (if [Index] = 1 then " DECLARE @COUNT_INSERT INT = 0;#(lf)" else "")
                                & "INSERT INTO "
                                & SqlTable
                                & " VALUES ("
                                & Text.Combine(
                                    List.Transform(
                                        Record.ToList(Record.RemoveFields(_, {"Index"})), each Sql.ConvertByType(_)
                                    ),
                                    ","
                                )
                                & " ); SET @COUNT_INSERT = @COUNT_INSERT + @@ROWCOUNT ;"
                                & (
                                    if Number.Mod([Index], 10000) = 0 or [Index] = RowCount then
                                        "#(lf)UPDATE  ##INSERT SET INSERT_ROWS_COUNT = INSERT_ROWS_COUNT + @COUNT_INSERT;"
                                            & "#(lf)GO"
                                            & "#(lf)DECLARE @COUNT_INSERT INT = 0;"
                                    else
                                        ""
                                )
                    ),
                    RemoveCols = Table.SelectColumns(AddValuesRows, {NewColumnName})
                in
                    RemoveCols,
            Table.ToSql = SqlQuery.Insert,
            Config = Table.PromoteHeaders(
                Table.Transpose(Excel.CurrentWorkbook(){[
                    Name = ConfigTableName
                ]}[Content]),
                [
                    PromoteAllScalars = true
                ]
            ){
                0
            },
            Source = Table.Combine({Excel.CurrentWorkbook(){[
                Name = Config[Excel Table]
            ]}[Content]}),
            ConvertTblToSql = Table.ToSql(Source, "Column1", Config[Sql Table]),
            DDL_DML_Query = Table.FromRecords(
                {
                    [
                        Column1 = Config[DDL Query]
                    ],
                    [
                        Column1 = if Config[Open Transaction] = "YES" then "BEGIN TRANSACTION " else ""
                    ],
                    [
                        Column1 = "SET NOCOUNT ON;#(lf)" & Config[DML Query]
                    ],
                    [
                        Column1 = "DROP TABLE IF EXISTS ##INSERT ;"
                    ],
                    [
                        Column1 = "SELECT  0 AS DML_ROWS_COUNT, 0 AS INSERT_ROWS_COUNT INTO ##INSERT ;"
                    ],
                    [
                        Column1 = "UPDATE ##INSERT SET DML_ROWS_COUNT = @@ROWCOUNT"
                    ]
                }
            ),
            EndTransaction = Table.FromRecords(
                {
                    [
                        Column1 = if Config[Open Transaction] = "YES" then "COMMIT TRANSACTION " else ""
                    ],
                    [
                        Column1 = if Config[Open Transaction] = "YES" then "SELECT * FROM ##INSERT " else ""
                    ]
                }
            ),
            Combine = Table.Combine({DDL_DML_Query, ConvertTblToSql, EndTransaction}),
            #"Split Column by Delimiter" = Table.ExpandListColumn(
                Table.TransformColumns(
                    Combine,
                    {
                        {
                            "Column1",
                            Splitter.SplitTextByDelimiter("#(lf)", QuoteStyle.Csv),
                            let
                                itemType = (type nullable text) meta [
                                    Serialized.Text = true
                                ]
                            in
                                type {
                                    itemType
                                }
                        }
                    }
                ),
                "Column1"
            ),
            #"Renamed Columns" = Table.RenameColumns(#"Split Column by Delimiter", {{"Column1", "COPY_TO_SSMS"}})
        in
            #"Renamed Columns"
in
    Main
