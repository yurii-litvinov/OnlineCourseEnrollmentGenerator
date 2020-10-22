open System
open System.IO
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open DocumentFormat.OpenXml

let fillEmails () =
    let mutable emailsMap = Map.empty

    use dataDocument = SpreadsheetDocument.Open("data.xlsx", false)
    let workbookPart = dataDocument.WorkbookPart
    let worksheetPart = workbookPart.WorksheetParts |> Seq.skip 1 |> Seq.head
    let sheetData = worksheetPart.Worksheet.Elements<Spreadsheet.SheetData>() |> Seq.head

    let sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>() |> Seq.head
    let sst = sstPart.SharedStringTable;

    let cellValue (row : Spreadsheet.Row)  i = 
        let cell = row.Elements<Spreadsheet.Cell>() |> Seq.skip i |> Seq.head
        if cell.DataType <> null && cell.DataType = EnumValue(Spreadsheet.CellValues.SharedString) then
            let ssid = cell.CellValue.Text |> int
            sst.ChildElements.[ssid].InnerText
        else
            cell.CellValue.Text
    
    for row in sheetData.Elements<Spreadsheet.Row>() do
        if row.Elements<Spreadsheet.Cell>() |> Seq.length > 12 then
            let surname = cellValue row 0
            let name = cellValue row 1
            let fatherName = cellValue row 2
            let faculty = cellValue row 6
            let level = cellValue row 7

            if faculty = "Математика и механика" && (level = "Бакалавр" || level = "Магистр") then
                let student = surname + " " + name + " " + fatherName

                let multipleMails = emailsMap.ContainsKey student

                if multipleMails then
                    printfn "Multiple mails for %s" student

                emailsMap <- emailsMap.Add (student, (cellValue row 18, multipleMails))

    emailsMap

let createParagraph text isBold =
    let paragraph = new Paragraph()
    let run = new Run()
    let text = new Text(text)
    
    let runProperties = run.AppendChild(new RunProperties())
    let fontSize = new FontSize(Val = StringValue("24"))
    runProperties.AppendChild(fontSize) |> ignore

    if isBold then
        let bold = new Bold()
        bold.Val <- OnOffValue.FromBoolean(true)
        runProperties.AppendChild(bold) |> ignore

    run.AppendChild(text) |> ignore
    paragraph.AppendChild(run) |> ignore

    paragraph

let createTable () =
    let table = Table()

    let tableBorders = 
        new TableBorders(
            TopBorder(Val = EnumValue<_>(BorderValues.BasicThinLines), Size = UInt32Value(1u)),
            BottomBorder(Val = EnumValue<_>(BorderValues.BasicThinLines), Size = UInt32Value(1u)),
            LeftBorder(Val = EnumValue<_>(BorderValues.BasicThinLines), Size = UInt32Value(1u)),
            RightBorder(Val = EnumValue<_>(BorderValues.BasicThinLines), Size = UInt32Value(1u)),
            InsideHorizontalBorder(Val = EnumValue<_>(BorderValues.BasicThinLines), Size = UInt32Value(1u)),
            InsideVerticalBorder(Val = EnumValue<_>(BorderValues.BasicThinLines), Size = UInt32Value(1u))
        )

    let tblProp = new TableProperties()
    tblProp.AppendChild(tableBorders) |> ignore
    table.AppendChild(tblProp) |> ignore

    table

let createTableRow (values: string seq) isBold =
    let tr = TableRow()

    for value in values do
        let tc = new TableCell()

        let tcWidth = TableCellWidth()
        tcWidth.Type <- EnumValue(TableWidthUnitValues.Dxa)
        tcWidth.Width <- StringValue("2500")
        let tcProps = TableCellProperties(tcWidth)

        tc.AppendChild(tcProps) |> ignore
        tc.AppendChild(createParagraph value isBold) |> ignore
        tr.AppendChild(tc) |> ignore

    tr

let addAdditionalData (map : Map<_, _>) =
    use input = new StreamReader(File.OpenRead("additionalData.txt"))
    let mutable map = map
    while not input.EndOfStream do
        let student = input.ReadLine()
        let tokens = student.Split(' ')
        map <- map.Add(tokens.[0] + " " + tokens.[1] + " " + tokens.[2], (tokens.[3], false))
    map

let cleanLine (str : string) =
    let tokenized = str.Split([|' '; '\t'|], StringSplitOptions.RemoveEmptyEntries)
    tokenized.[1] + " " + tokenized.[2] + " " + tokenized.[3]

[<EntryPoint>]
let main argv =
    use input = new StreamReader(File.OpenRead(argv.[0]))
    let disciplineName = input.ReadLine().Replace('\t', ' ')
    let courseName = input.ReadLine()
    let courseLink = input.ReadLine()
    let instructor = input.ReadLine()
    let instructorEmail = input.ReadLine()

    use wordDocument = WordprocessingDocument.Create(courseName + ".docx", WordprocessingDocumentType.Document, true)
    let mainPart = wordDocument.AddMainDocumentPart()
    
    mainPart.Document <- new Document()
    let body = new Body()
    let header = createParagraph ("Курс " + courseName + " на Coursera, " + courseLink) false
    let discipline = createParagraph ("Для поддержки дисциплины " + disciplineName) false

    body.AppendChild(header) |> ignore
    body.AppendChild(discipline) |> ignore

    let table = createTable()

    let header = createTableRow ["№"; "Фамилия, имя и отчество"; "Студент / Преподаватель"; "Адрес электронной почты"] true
    table.AppendChild(header) |> ignore

    input.ReadLine() |> ignore
    let mutable i = 0

    let emailsMap = fillEmails () |> addAdditionalData
    
    while not input.EndOfStream do
        i <- i + 1
        let student' = input.ReadLine()
        let student = cleanLine student'
        let email = 
            if emailsMap.ContainsKey student then 
                if emailsMap.[student] |> snd then 
                    "" 
                else 
                    emailsMap.[student] |> fst 
            else 
                ""

        let tr = createTableRow [i.ToString(); student; "Студент"; email] false
        table.AppendChild(tr) |> ignore
        ()

    i <- i + 1
    let tr = createTableRow [i.ToString(); instructor; "Преподаватель"; instructorEmail] false
    table.AppendChild(tr) |> ignore

    body.AppendChild(table) |> ignore

    mainPart.Document.AppendChild(body) |> ignore

    0
