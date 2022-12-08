
using ExcelReader;
using SpreadsheetLight;

ReadXLS();

Console.WriteLine("Done");
Console.Read();


void ReadXLS()
{
    List<FileModel> dataList = new();

    using (var doc = new SLDocument("C:\\Users\\bhara\\Desktop\\AA\\xls.xlsx"))
    {
        var stats = doc.GetWorksheetStatistics();

        List<FileModel> fileObj = new();

        

        for (int row = 2; row < stats.EndRowIndex +1; row++)
        {
            FileModel dataObj = new();

            int counter = 0;

            for (int column = 1; column < stats.EndColumnIndex; column++)
            {
                counter++;
                if(column == 26)
                {
                    continue;
                }
                else
                {
                    switch (column)
                    {
                        case (1):
                            dataObj.Superclass = doc.GetCellValueAsString(row, column);
                            break;
                        case (2):
                            dataObj.Label = doc.GetCellValueAsString(row, column);
                            break;
                        case (3):
                            dataObj.IRI = doc.GetCellValueAsString(row, column);
                            break;
                        case (4):
                            dataObj.Type = doc.GetCellValueAsString(row, column);
                            break;
                        case (5):
                            dataObj.Sali_appealsTo = doc.GetCellValueAsString(row, column);
                            break;
                        case (6):
                            dataObj.Sali_hasExpense = doc.GetCellValueAsString(row, column);
                            break;
                        case (7):
                            dataObj.Sali_filed = doc.GetCellValueAsString(row, column);
                            break;
                        case (8):
                            dataObj.Sali_seeksToAchieve = doc.GetCellValueAsString(row, column);
                            break;
                        case (9):
                            dataObj.Sali_workedFor = doc.GetCellValueAsString(row, column);
                            break;
                        case (10):
                            dataObj.Sali_cited = doc.GetCellValueAsString(row, column);
                            break;
                        case (11):
                            dataObj.Sali_participatedIn = doc.GetCellValueAsString(row, column);
                            break;
                        case (12):
                            dataObj.LocatedIn = doc.GetCellValueAsString(row, column);
                            break;
                        case (13):
                            dataObj.SeeAlso = doc.GetCellValueAsString(row, column);
                            break;
                        case (14):
                            dataObj.IsAuthor = doc.GetCellValueAsString(row, column);
                            break;
                        case (15):
                            dataObj.SameAs = doc.GetCellValueAsString(row, column);
                            break;
                        case (16):
                            dataObj.Before = doc.GetCellValueAsString(row, column);
                            break;
                        case (17):
                            dataObj.Sali_seealso = doc.GetCellValueAsString(row, column);
                            break;
                        case (18):
                            dataObj.legacyIdentifier = doc.GetCellValueAsString(row, column);
                            break;
                        case (19):
                            dataObj.Description = doc.GetCellValueAsString(row, column);
                            break;
                        case (20):
                            dataObj.Identifier = doc.GetCellValueAsString(row, column);
                            break;
                        case (21):
                            dataObj.LinkFirst = doc.GetCellValueAsString(row, column);
                            break;
                        case (22):
                            dataObj.LinkSecond = doc.GetCellValueAsString(row, column);
                            break;
                        case (23):
                            dataObj.HasRelatedSynonym = doc.GetCellValueAsString(row, column);
                            break;
                        case (24):
                            dataObj.Comment = doc.GetCellValueAsString(row, column);
                            break;
                        case (25):
                            dataObj.IsDefinedBy = doc.GetCellValueAsString(row, column);
                            break;
                        case (27):
                            dataObj.Deprecated = doc.GetCellValueAsString(row, column);
                            break;
                        case (28):
                            dataObj.InverseOf = doc.GetCellValueAsString(row, column);
                            break;
                        case (29):
                            dataObj.AltLabel = doc.GetCellValueAsString(row, column);
                            break;
                        case (30):
                            dataObj.Definition = doc.GetCellValueAsString(row, column);
                            break;
                        case (31):
                            dataObj.HiddenLabel = doc.GetCellValueAsString(row, column);
                            break;
                        case (32):
                            dataObj.PrefLabel = doc.GetCellValueAsString(row, column);
                            break;
                        
                            default:
                            Console.WriteLine("File reading Error");
                            break;
                    }
                }

            }
            dataList.Add(dataObj);

        }
        AddBulk(dataList);
    }
}

void AddBulk(List<FileModel> fileData)
{
    for (int i = 0; i<fileData.Count; i+=500)
    {
        var toAddList = fileData.Skip(i).Take(500).ToList();

        /*_context.Table.AddRange(toAddList);
        _context.SaveChanges();*/

        Console.WriteLine("Added {0} data", toAddList.Count);


        /*foreach (var item in toAddList)
        {
            Console.WriteLine(item.PrefLabel);
        }*/
    }
    Console.WriteLine("Added {0} data in total", fileData.Count);

}