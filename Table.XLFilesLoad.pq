let
    FUNC =
        (MODULES_TBL as text, FOLDER_PATH_RANGE as text) =>
            let
                // Can be Local Path like "C:\GIC\" or Sharepoint like "https://xxxx.sharepoint.com/sites/xxxx/"
                CurrentFolderPath = Excel.CurrentWorkbook(){[Name = FOLDER_PATH_RANGE]}[Content]{0}[Column1],
                // Extract Root URL from current Folder Path
                UDF.SharepointRoot =
                    (__Folder_Path as text) =>
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
                // Get List File to be combined infomation
                FileList = Excel.CurrentWorkbook(){[Name = MODULES_TBL]}[Content],
                // Filter IS_LOAD = YES
                FileList.IsLoadOnly =
                    Table.SelectRows(
                        FileList,
                        each ([IS_LOAD] = "YES")
                    ),
                // Group to File scope. Only download/open files once.
                FileList.GroupFile =
                    Table.Group(
                        FileList.IsLoadOnly,
                        {
                            "FILE_ID",
                            "FILE_ID_LEN",
                            "FIXED_PATH"
                        },
                        {
                            {
                                "FILE_LIST_PROPERTIES",
                                each _
                            }
                        }
                    ),
                // Detect if a Fixed Path is configed or use a Dynamic Folder Path
                FileList.FolderFilePath =
                    Table.AddColumn(
                        FileList.GroupFile,
                        "FOLDER_FILE_PATH",
                        each
                            if [FIXED_PATH] = null or [FIXED_PATH] = "" then
                                CurrentFolderPath
                            else
                                [FIXED_PATH]
                    ),
                // Detect if path is Folder Path on Sharepoint or on Local Machine
                FileList.DetectPathType =
                    Table.AddColumn(
                        FileList.FolderFilePath,
                        "PATH_TYPE",
                        each
                            if
                                Text.Start([FOLDER_FILE_PATH], 4)
                                = "http"
                            then
                                if
                                    Text.Contains(
                                        [FOLDER_FILE_PATH],
                                        "sharepoint",
                                        Comparer.OrdinalIgnoreCase
                                    )
                                then
                                    "SHAREPOINT"
                                else
                                    "WEB"
                            else
                                "LOCAL"
                    ),
                // Open Path using corresponding Path Type
                FileList.OpenPath =
                    Table.AddColumn(
                        FileList.DetectPathType,
                        "OPEN_PATH",
                        each
                            if [PATH_TYPE] = "LOCAL" then
                                Folder.Files([FOLDER_FILE_PATH])
                            else if [PATH_TYPE] = "SHAREPOINT" then
                                UDF.SharepointRoot([FOLDER_FILE_PATH])
                            else
                                null
                    ),
                // Expand Binary Content and File Names
                FileList.ExpandFolder =
                    Table.ExpandTableColumn(
                        FileList.OpenPath,
                        "OPEN_PATH",
                        {
                            "Content",
                            "Name"
                        },
                        {
                            "Content",
                            "Name"
                        }
                    ),
                // Filter Files base on ID in the end of File Name
                FileList.FilterFiles =
                    Table.SelectRows(
                        FileList.ExpandFolder,
                        each
                            Text.Contains(
                                [Name],
                                [FILE_ID] & ".xlsx",
                                Comparer.OrdinalIgnoreCase
                            )
                    ),
                // Open Excel File
                FileList.OpenFile =
                    Table.AddColumn(
                        FileList.FilterFiles,
                        "OPEN_FILE",
                        each Excel.Workbook([Content])
                    ),
                // Expand All Table/Sheet Data in Files
                FileList.ExpandFileContent =
                    Table.ExpandTableColumn(
                        FileList.OpenFile,
                        "OPEN_FILE",
                        {
                            "Data",
                            "Item",
                            "Kind"
                        },
                        {
                            "Excel.Data",
                            "Excel.Item",
                            "Excel.Kind"
                        }
                    ),
                // Expand Origin FileList Properties
                FileList.ExpandTargetFilter =
                    Table.ExpandTableColumn(
                        FileList.ExpandFileContent,
                        "FILE_LIST_PROPERTIES",
                        {
                            "DESC",
                            "DATA_TYPE",
                            "OBJECT_NAME",
                            "OBJECT_TYPE",
                            "VERSION"
                        },
                        {
                            "DESC",
                            "DATA_TYPE",
                            "OBJECT_NAME",
                            "OBJECT_TYPE",
                            "VERSION"
                        }
                    ),
                // Filter Table/Sheet
                FileList.FilterTarget =
                    Table.SelectRows(
                        FileList.ExpandTargetFilter,
                        each
                            [OBJECT_NAME]
                            = [Excel.Item]
                            and [OBJECT_TYPE]
                            = [Excel.Kind]
                    ),
                // Filter Columns needed
                FileList.FilterColumns =
                    Table.SelectColumns(
                        FileList.FilterTarget,
                        {
                            "FILE_ID",
                            "DESC",
                            "DATA_TYPE",
                            "VERSION",
                            "Excel.Data"
                        }
                    ),
                // Add FILE ID, DESC, etc to final Table result
                FileList.AddFileProperties =
                    Table.AddColumn(
                        FileList.FilterColumns,
                        "Excel.DataModified",
                        each
                            let
                                _prop =
                                    Record.SelectFields(
                                        _,
                                        {
                                            "FILE_ID",
                                            "DESC",
                                            "DATA_TYPE",
                                            "VERSION"
                                        }
                                    )
                            in
                                Table.AddColumn(
                                    [Excel.Data],
                                    "FILE_PROPERTIES",
                                    each _prop
                                )
                    ),
                // Combine all Table from Excel Files
                FileList.CombineAll = Table.Combine(FileList.AddFileProperties[Excel.DataModified]),
                // Expand File ID from FileList
                FileList.ExpandFileProperties =
                    Table.ExpandRecordColumn(
                        FileList.CombineAll,
                        "FILE_PROPERTIES",
                        {
                            "FILE_ID",
                            "DESC",
                            "DATA_TYPE",
                            "VERSION"
                        },
                        {
                            "FILE_ID",
                            "DESC",
                            "DATA_TYPE",
                            "VERSION"
                        }
                    )
            in
                FileList.ExpandFileProperties,
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