let
    FUNC =
        (MODULES_TBL as text, FOLDER_PATH_RANGE as text) =>
            let
                Source = Excel.CurrentWorkbook(){[Name = MODULES_TBL]}[Content],
                IS_LOAD_FILTER =
                    Table.SelectRows(
                        Source,
                        each ([IS_LOAD] = "YES")
                    ),
                GROUP_DUP_FILES =
                    Table.Group(
                        IS_LOAD_FILTER,
                        {
                            "FILE_ID",
                            "FILE_ID_LEN",
                            "OBJECT_NAME",
                            "OBJECT_TYPE",
                            "FIXED_PATH"
                        },
                        {
                            {
                                "TBL",
                                each _
                            }
                        }
                    ),
                LOAD_XL_DATA =
                    Table.AddColumn(
                        GROUP_DUP_FILES,
                        "FILE_BINARY",
                        each
                            let
                                __FILE_ID_LEN = [FILE_ID_LEN],
                                __FILE_ID = [FILE_ID],
                                __OBJ_NAME = [OBJECT_NAME],
                                __OBJ_TYPE = [OBJECT_TYPE],
                                __Fixed_Path = [FIXED_PATH],
                                __Dynamic_Path = Excel.CurrentWorkbook(){[Name = FOLDER_PATH_RANGE]}[Content]{0}[Column1],
                                __Folder_Path =
                                    if __Fixed_Path = null or __Fixed_Path = "" then
                                        __Dynamic_Path
                                    else
                                        __Fixed_Path,
                                __LOCAL_FOLDER = Folder.Files(__Folder_Path),
                                __SHAREPOINT_FOLDER =
                                    Table.SelectRows(
                                        SharePoint.Files(
                                            Text.Start(
                                                __Folder_Path,
                                                Text.PositionOfAny(
                                                    __Folder_Path,
                                                    {"/"},
                                                    Occurrence.All
                                                ){4}
                                                + 1
                                            )
                                        ),
                                        each
                                            Text.Contains(
                                                [Folder Path],
                                                __Folder_Path
                                            )
                                    ),
                                __FILTER_FOLDER =
                                    Table.SelectRows(
                                        try __LOCAL_FOLDER otherwise __SHAREPOINT_FOLDER,
                                        each
                                            Text.End([Name], 5 + __FILE_ID_LEN)
                                            = (__FILE_ID & ".xlsx")
                                    ),
                                __OpenExcel = Excel.Workbook(__FILTER_FOLDER[Content]{0}),
                                __Get_Model_Input =
                                    __OpenExcel{[
                                        Item = __OBJ_NAME,
                                        Kind = __OBJ_TYPE
                                    ]}
                                        [Data]
                            in
                                try __Get_Model_Input otherwise null
                    ),
                CHOOSE_COLUMNS =
                    let
                        __LIST_COL_NAMES = List.Buffer(Table.ColumnNames(Table.Combine(LOAD_XL_DATA[FILE_BINARY]))),
                        __TRY_EXPAND =
                            Table.ExpandTableColumn(
                                LOAD_XL_DATA,
                                "FILE_BINARY",
                                __LIST_COL_NAMES,
                                __LIST_COL_NAMES
                            )
                    in
                        try __TRY_EXPAND otherwise LOAD_XL_DATA,
                EXPAND_TBL =
                    Table.ExpandTableColumn(
                        CHOOSE_COLUMNS,
                        "TBL",
                        {
                            "DESC",
                            "DATA_TYPE",
                            "VERSION"
                        },
                        {
                            "DESC",
                            "DATA_TYPE",
                            "VERSION"
                        }
                    )
            in
                EXPAND_TBL,
    META =
        type function (//
        MODULES_TBL as
            (
                type text
                meta
                [
                    Documentation.FieldCaption = "Table have list XLSX Files to be loaded.",
                    Documentation.SampleValues = {
                        "Table Name like: Table1, Table2, ..."
                    }
                ]
            ), FOLDER_PATH_RANGE as
            (
                type text
                meta
                [
                    Documentation.FieldCaption = "Range Name have Current Folder Path",
                    Documentation.SampleValues = {
                        "CURRENT_FOLDER_PATH"
                    }
                ]
            )) as table
        meta
        [
            Documentation.Name = "XLFilesLoad",
            Documentation.Description = "Automate Combine Tables/Sheet from Multiple XLSX Files",
            Documentation.Author = "Kit Cognac",
            Documentation.Examples = {
                [
                    Description = "Replace values on all columns",
                    Code =
                        "
Create a table in XLSX as below, name it Table1:

| FILE_ID | FILE_ID_LEN   | DESC   | DATA_TYPE | OBJECT_NAME | OBJECT_TYPE | IS_LOAD | VERSION | FIXED_PATH |
|---------|---------------|--------|-----------|-------------|-------------|---------|---------|------------|
| DT1     | LEN([FILE_ID] | Data 1 | Actual    | Table1      | Table       | YES     |         | D:\ABC\    |
| DT2     | LEN([FILE_ID] | Data 2 | Actual    | Table1      | Table       | YES     |         | D:\ABC\    |

Name a cell ""CURRENT_FOLDER_PATH"" with formula =LEFT(CELL(""filename"",$A$1),FIND(""["",CELL(""filename"",$A$1),1)-1)

Call function = XLFilesLoad(""Table1"",""CURRENT_FOLDER_PATH"")
                    ",
                    Result = "Table output will try to auto expand, if it fails you need to manually expand"
                ]
            }
        ]
in
    Value.ReplaceType(FUNC, META)