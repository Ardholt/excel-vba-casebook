let
    // Returns a record with Combined, Manifest and Rejected tables.
    SourceFolder = Text.From(Excel.CurrentWorkbook(){[Name="SourceFolder"]}[Content]{0}[Column1]),
    Required = {"record_id", "date", "category", "amount"},
    Origin = {"source_file", "source_row"},
    OutputNames = {"consolidated.xlsx", "vba-result.xlsx", "combine.xlsm"},
    Empty = #table({}, {}),
    ReadFile = (content as binary, filename as text) as record =>
        let
            Open = try Excel.Workbook(content, null, true),
            Sheet = if Open[HasError] then null else
                Table.SelectRows(Open[Value], each [Item] = "Data" and [Kind] = "Sheet"),
            Raw = if Sheet = null then null else if Table.RowCount(Sheet) = 0 then null else Sheet{0}[Data],
            Headers = if Raw = null then {} else if Table.RowCount(Raw) = 0 then {} else
                List.Transform(Record.FieldValues(Raw{0}), each if _ = null then "" else Text.Trim(Text.From(_))),
            Reason = if Open[HasError] then "Unreadable workbook"
                else if Raw = null then "Missing Data sheet"
                else if List.Count(Headers) = 0 or List.Contains(Headers, "") then "Blank header"
                else if List.Count(List.Distinct(Headers)) <> List.Count(Headers) then "Duplicate header"
                else if List.Count(List.Intersect({Headers, Origin})) > 0 then "Reserved provenance header"
                else if not List.ContainsAll(Headers, Required) then "Missing required columns"
                else "",
            Parsed = if Reason <> "" then Empty else
                let
                    Body = Table.Skip(Raw, 1),
                    Named = Table.RenameColumns(Body, List.Zip({Table.ColumnNames(Body), Headers})),
                    Indexed = Table.AddIndexColumn(Named, "source_row", 2, 1, Int64.Type),
                    Nonblank = Table.SelectRows(Indexed, (row) =>
                        List.AnyTrue(List.Transform(Headers, (h) =>
                            let v = Record.Field(row, h) in v <> null and v <> ""))),
                    Traced = Table.AddColumn(Nonblank, "source_file", each filename, type text)
                in Traced,
            Count = Table.RowCount(Parsed)
        in [file=filename, status=if Reason <> "" then "rejected" else if Count=0 then "empty" else "accepted",
            rows=Count, reason=Reason, data=Parsed],
    Entries = Table.SelectRows(Folder.Contents(SourceFolder), each [Attributes]?[Directory]? <> true),
    Sorted = Table.Sort(Entries, {{"Name", Order.Ascending}}),
    Decisions = List.Transform(Table.ToRecords(Sorted), (f) =>
        let
            Name = f[Name],
            Ignore = Text.StartsWith(Name, "~$") or List.Contains(OutputNames, Text.Lower(Name)),
            IsWorkbook = Text.EndsWith(Text.Lower(Name), ".xlsx"),
            Attempt = if Ignore or not IsWorkbook then null else try ReadFile(f[Content], Name)
        in if Ignore or not IsWorkbook then
            [file=Name,status="ignored",rows=0,reason=if Ignore then "Temporary or output file" else "Not an .xlsx workbook",data=Empty]
        else if Attempt[HasError] then
            [file=Name,status="rejected",rows=0,reason="Unreadable workbook",data=Empty]
        else Attempt[Value]),
    Joined = Table.Combine(List.Transform(Decisions, each [data])),
    Columns = Required & Origin & List.Difference(Table.ColumnNames(Joined), Required & Origin),
    Combined = Table.SelectColumns(Joined, Columns, MissingField.UseNull),
    Manifest = Table.FromRecords(List.Transform(Decisions, each Record.RemoveFields(_, {"data"})),
        type table [file=text,status=text,rows=Int64.Type,reason=text]),
    Rejected = Table.SelectRows(Manifest, each [status]="rejected")
in
    [Combined=Combined, Manifest=Manifest, Rejected=Rejected]
