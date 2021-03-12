using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MTG_Parse
{
    class Program
    {
        object missing = Type.Missing;
        public Microsoft.Office.Interop.Excel.Application oXL;
        public Microsoft.Office.Interop.Excel._Workbook oWB;
        public Microsoft.Office.Interop.Excel._Worksheet IzzetSheet;
        public Microsoft.Office.Interop.Excel.Range oRng;
        public Microsoft.Office.Interop.Excel.Worksheet BlackSheet;
        public Microsoft.Office.Interop.Excel.Worksheet RedSheet;
        public Microsoft.Office.Interop.Excel.Worksheet BlueSheet;
        public Microsoft.Office.Interop.Excel.Worksheet GreenSheet;
        public Microsoft.Office.Interop.Excel.Worksheet WhiteSheet;
        public Microsoft.Office.Interop.Excel.Worksheet DimirSheet;
        public Microsoft.Office.Interop.Excel.Worksheet SelesnyaSheet;
        public Microsoft.Office.Interop.Excel.Worksheet BorosSheet;
        public Microsoft.Office.Interop.Excel.Worksheet GolgariSheet;
        public Microsoft.Office.Interop.Excel.Worksheet ElseSheet;
        public RootItem array;
        static void Main(string[] args)
        {
            
            
            Program p = new Program();
            p.LoadJson();
            p.InitExcel();
            p.Print();
            while (true) { }
        }

        public void LoadJson()
        {
            string path = "X:\\Git\\MTG\\MTG_Parse\\MTG_Parse\\GRN.json";
            using (StreamReader r = new StreamReader(path))
            {
                string json = r.ReadToEnd();
                array = JsonConvert.DeserializeObject<RootItem>(json);
            }
        }
        public void Print()
        {
            foreach (var item in array.cards)
            {
                if (item.colors != null && item.colors.Count() > 1)
                {
                    if (item.colors[0].Equals("White", StringComparison.OrdinalIgnoreCase) && (item.colors[1].Equals("Red", StringComparison.OrdinalIgnoreCase)))
                    {
                        InputData(BorosSheet, item.name, item.cmc, item.types);
                    } else if (item.colors[0].Equals("White", StringComparison.OrdinalIgnoreCase) && (item.colors[1].Equals("Green", StringComparison.OrdinalIgnoreCase)))
                    {
                        InputData(SelesnyaSheet, item.name, item.cmc, item.types);
                    } else if (item.colors[0].Equals("Blue", StringComparison.OrdinalIgnoreCase) && (item.colors[1].Equals("Red", StringComparison.OrdinalIgnoreCase)))
                    {
                        InputData(IzzetSheet, item.name, item.cmc, item.types);
                    } else if (item.colors[0].Equals("Blue", StringComparison.OrdinalIgnoreCase) && (item.colors[1].Equals("Black", StringComparison.OrdinalIgnoreCase)))
                    {
                        InputData(DimirSheet, item.name, item.cmc, item.types);
                    } else if (item.colors[0].Equals("Black", StringComparison.OrdinalIgnoreCase) && (item.colors[1].Equals("Green", StringComparison.OrdinalIgnoreCase)))
                    {
                        InputData(GolgariSheet, item.name, item.cmc, item.types);
                    } else
                    {
                        InputData(ElseSheet, item.name, item.cmc, item.types);
                    }
                } else
                {
                    if(item.colors == null)
                    {
                        InputData(ElseSheet, item.name, item.cmc, item.types);
                    } else
                    {
                        switch (item.colors[0])
                        {
                            case "Blue":
                                InputData(BlueSheet, item.name, item.cmc, item.types);
                                break;
                            case "Green":
                                InputData(GreenSheet, item.name, item.cmc, item.types);
                                break;
                            case "Black":
                                InputData(BlackSheet, item.name, item.cmc, item.types);
                                break;
                            case "Red":
                                InputData(RedSheet, item.name, item.cmc, item.types);
                                break;
                            case "White":
                                InputData(WhiteSheet, item.name, item.cmc, item.types);
                                break;
                            default:
                                InputData(ElseSheet, item.name, item.cmc, item.types);
                                break;
                        }
                    }
                    
                }
            }
            Console.WriteLine("Done");
        }
        public void InputData(Microsoft.Office.Interop.Excel.Worksheet s, string cardname, string cmc, string[] type)
        {
            oRng = (Microsoft.Office.Interop.Excel.Range)s.Cells[s.Rows.Count, 1];
            long lastRow = oRng.get_End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row;
            long newRow = lastRow + 1;
            s.Cells[newRow, 1] = cardname;
            if(cmc == null)
            {
                s.Cells[newRow, 2] = 0;
            } else
            {
                s.Cells[newRow, 2] = cmc;
            }
            s.Cells[newRow, 3] = type;
        }
        public void InputData(Microsoft.Office.Interop.Excel._Worksheet s, string cardname, string cmc, string[] type)
        {
            oRng = (Microsoft.Office.Interop.Excel.Range)s.Cells[s.Rows.Count, 1];
            long lastRow = oRng.get_End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row;
            long newRow = lastRow + 1;
            s.Cells[newRow, 1] = cardname;
            if (cmc == null)
            {
                s.Cells[newRow, 2] = 0;
            } else
            {
                s.Cells[newRow, 2] = cmc;
            }
            s.Cells[newRow, 3] = type;
        }
        public void InitExcel()
        {
            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;

            //Get a new workbook.
            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            //Izzet Sheet
            IzzetSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            IzzetSheet.Name = "Red Blue";
            IzzetSheet.Cells[1, 1] = "Card Name";
            IzzetSheet.Cells[1, 2] = "Mana Cost";
            IzzetSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            IzzetSheet.Columns[1].ColumnWidth = 26;
            IzzetSheet.Columns[2].ColumnWidth = 10;
            IzzetSheet.Columns[3].ColumnWidth = 12;
            //Dimir Sheet
            DimirSheet = oWB.Sheets.Add(missing, missing, 1, missing)
                         as Microsoft.Office.Interop.Excel.Worksheet;
            DimirSheet.Name = "Blue Black";
            DimirSheet.Cells[1, 1] = "Card Name";
            DimirSheet.Cells[1, 2] = "Mana Cost";
            DimirSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            DimirSheet.Columns[1].ColumnWidth = 26;
            DimirSheet.Columns[2].ColumnWidth = 10;
            DimirSheet.Columns[3].ColumnWidth = 12;
            //Golgari Sheet
            GolgariSheet = oWB.Sheets.Add(missing, missing, 1, missing)
                         as Microsoft.Office.Interop.Excel.Worksheet;
            GolgariSheet.Name = "Black Green";
            GolgariSheet.Cells[1, 1] = "Card Name";
            GolgariSheet.Cells[1, 2] = "Mana Cost";
            GolgariSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            GolgariSheet.Columns[1].ColumnWidth = 26;
            GolgariSheet.Columns[2].ColumnWidth = 10;
            GolgariSheet.Columns[3].ColumnWidth = 12;
            //Selesnya Sheet
            SelesnyaSheet = oWB.Sheets.Add(missing, missing, 1, missing)
                         as Microsoft.Office.Interop.Excel.Worksheet;
            SelesnyaSheet.Name = "Green White";
            SelesnyaSheet.Cells[1, 1] = "Card Name";
            SelesnyaSheet.Cells[1, 2] = "Mana Cost";
            SelesnyaSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            SelesnyaSheet.Columns[1].ColumnWidth = 26;
            SelesnyaSheet.Columns[2].ColumnWidth = 10;
            SelesnyaSheet.Columns[3].ColumnWidth = 12;
            //Boros Sheet
            BorosSheet = oWB.Sheets.Add(missing, missing, 1, missing)
                         as Microsoft.Office.Interop.Excel.Worksheet;
            BorosSheet.Name = "White Red";
            BorosSheet.Cells[1, 1] = "Card Name";
            BorosSheet.Cells[1, 2] = "Mana Cost";
            BorosSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            BorosSheet.Columns[1].ColumnWidth = 26;
            BorosSheet.Columns[2].ColumnWidth = 10;
            BorosSheet.Columns[3].ColumnWidth = 12;
            //Red Sheet
            RedSheet = oWB.Sheets.Add(missing, missing, 1, missing)
                        as Microsoft.Office.Interop.Excel.Worksheet;
            RedSheet.Name = "Red";
            RedSheet.Cells[1, 1] = "Card Name";
            RedSheet.Cells[1, 2] = "Mana Cost";
            RedSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            RedSheet.Columns[1].ColumnWidth = 26;
            RedSheet.Columns[2].ColumnWidth = 10;
            RedSheet.Columns[3].ColumnWidth = 12;
            //Black Sheet
            BlackSheet = oWB.Sheets.Add(missing, missing, 1, missing)
                         as Microsoft.Office.Interop.Excel.Worksheet;
            BlackSheet.Name = "Black";
            BlackSheet.Cells[1, 1] = "Card Name";
            BlackSheet.Cells[1, 2] = "Mana Cost";
            BlackSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            BlackSheet.Columns[1].ColumnWidth = 26;
            BlackSheet.Columns[2].ColumnWidth = 10;
            BlackSheet.Columns[3].ColumnWidth = 12;
            //Green Sheet
            GreenSheet = oWB.Sheets.Add(missing, missing, 1, missing)
                         as Microsoft.Office.Interop.Excel.Worksheet;
            GreenSheet.Name = "Green";
            GreenSheet.Cells[1, 1] = "Card Name";
            GreenSheet.Cells[1, 2] = "Mana Cost";
            GreenSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            GreenSheet.Columns[1].ColumnWidth = 26;
            GreenSheet.Columns[2].ColumnWidth = 10;
            GreenSheet.Columns[3].ColumnWidth = 12;
            //White Sheet
            WhiteSheet = oWB.Sheets.Add(missing, missing, 1, missing)
                         as Microsoft.Office.Interop.Excel.Worksheet;
            WhiteSheet.Name = "White";
            WhiteSheet.Cells[1, 1] = "Card Name";
            WhiteSheet.Cells[1, 2] = "Mana Cost";
            WhiteSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            WhiteSheet.Columns[1].ColumnWidth = 26;
            WhiteSheet.Columns[2].ColumnWidth = 10;
            WhiteSheet.Columns[3].ColumnWidth = 12;
            //Blue Sheet
            BlueSheet = oWB.Sheets.Add(missing, missing, 1, missing)
                         as Microsoft.Office.Interop.Excel.Worksheet;
            BlueSheet.Name = "Blue";
            BlueSheet.Cells[1, 1] = "Card Name";
            BlueSheet.Cells[1, 2] = "Mana Cost";
            BlueSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            BlueSheet.Columns[1].ColumnWidth = 26;
            BlueSheet.Columns[2].ColumnWidth = 10;
            BlueSheet.Columns[3].ColumnWidth = 12;
            //Else Sheet
            ElseSheet = oWB.Sheets.Add(missing, missing, 1, missing)
                         as Microsoft.Office.Interop.Excel.Worksheet;
            ElseSheet.Name = "Else";
            ElseSheet.Cells[1, 1] = "Card Name";
            ElseSheet.Cells[1, 2] = "Mana Cost";
            ElseSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            ElseSheet.Columns[1].ColumnWidth = 26;
            ElseSheet.Columns[2].ColumnWidth = 10;
            ElseSheet.Columns[3].ColumnWidth = 12;

        }
    }
    public class RootItem
    {
        public string name { get; set; }
        public string code { get; set; }
        public string releaseDate { get; set; }
        public string border { get; set; }
        public string type { get; set; }
        public Dictionary<string, List<string>> Boosters = new Dictionary<string, List<string>>();
        public List<Cards> cards { get; set; }
    }

    

    public class Cards
    {
        public string artist { get; set; }
        public string cmc;
        public string[] colorIdentity { get; set; }
        public string[] colors { get; set; }
        //public Colors colors = new Colors();
        public string flavor { get; set; }
        public string id { get; set; }
        public string imageName { get; set; }
        public string layout { get; set; }
        public string manaCost { get; set; }
        public string multiverseid { get; set; }
        public string name { get; set; }
        public string number { get; set; }
        public string power { get; set; }
        public string rarity { get; set; }
        public string[] subtypes { get; set; }
        public string text { get; set; }
        public string toughness { get; set; }
        public string type { get; set; }
        public string[] types { get; set; }
        public string watermark { get; set; }
    }
}
