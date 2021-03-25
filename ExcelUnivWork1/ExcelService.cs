using System.IO;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Hosting;

namespace ExcelUnivWork1
{
    public class ExcelService
    {
        private readonly IWebHostEnvironment _hostingEnvironment;

        public ExcelService(IWebHostEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        public async Task<MemoryStream> Get(
            Request request)
        {
            var memoryStream = new MemoryStream();

            var workbook =
                new XLWorkbook(Path.Combine(_hostingEnvironment.ContentRootPath, "dataset-example.xlsx"));

            var originalWorksheet = workbook.Worksheet("original");

            var narmalize01Worksheet = workbook.Worksheet("norm (0;1)");
            var narmalizeMin11Worksheet = workbook.Worksheet("norm (-1;1)");

            var firstOriginalDataRow = GetFirstDataRow(originalWorksheet);
            var narmalize01DataRow = GetFirstDataRow(narmalize01Worksheet);
            var narmalizeMin11DataRow = GetFirstDataRow(narmalizeMin11Worksheet);


            var kEnd = 0.6;
            var kStart = 0.1;
            var v0End = 0.8;
            var v0Start = 0.1;

            var v0step = (v0End - v0Start) / 49.0;
            var kstep = (kEnd - kStart) / 49.0;

            var iForT0 = 0;
            var k = kStart;
            var v0 = v0Start;
            var dt = 0.013;

            for (int i = 1; i <= 125000; i++)
            {
                var currentRow = i + 1;
                firstOriginalDataRow.Cell(1).Value = iForT0;
                iForT0++;
                firstOriginalDataRow.Cell(1).SetDataType(XLDataType.Number);
                firstOriginalDataRow.Cell(2).Value = k;
                firstOriginalDataRow.Cell(2).SetDataType(XLDataType.Number);

                firstOriginalDataRow.Cell(3).FormulaA1 = $"=2*PI()*SQRT($B$8/G{currentRow})";
                firstOriginalDataRow.Cell(4).FormulaA1 = $"=$B$4+F{currentRow}*H{currentRow}/($D$4-1)";

                firstOriginalDataRow.Cell(5).Value = v0;
                firstOriginalDataRow.Cell(5).SetDataType(XLDataType.Number);

                firstOriginalDataRow.Cell(6).Value = dt;
                firstOriginalDataRow.Cell(6).SetDataType(XLDataType.Number);

                firstOriginalDataRow.Cell(8).FormulaA1 = $"=I{currentRow}";
                firstOriginalDataRow.Cell(9).FormulaA1 = $"=M{currentRow}+$K$2";
                firstOriginalDataRow.Cell(10).FormulaA1 = $"=N{currentRow}+$K$2";
                firstOriginalDataRow.Cell(11).FormulaA1 = $"=O{currentRow}+$K$2";
                firstOriginalDataRow.Cell(12).FormulaA1 = $"=P{currentRow}+$K$2";
                firstOriginalDataRow.Cell(13).FormulaA1 = $"=Q{currentRow}+$K$2";

                firstOriginalDataRow.Cell(15).FormulaA1 = $"=(SQRT($B$8)*$J{currentRow}*SIN((SQRT($G{currentRow})/SQRT($B$8))*$M{currentRow}))/SQRT($G{currentRow})";
                firstOriginalDataRow.Cell(16).FormulaA1 = $"=(SQRT($B$8)*$J{currentRow}*SIN((SQRT($G{currentRow})/SQRT($B$8))*$N{currentRow}))/SQRT($G{currentRow})";
                firstOriginalDataRow.Cell(17).FormulaA1 = $"=(SQRT($B$8)*$J{currentRow}*SIN((SQRT($G{currentRow})/SQRT($B$8))*$O{currentRow}))/SQRT($G{currentRow})";
                firstOriginalDataRow.Cell(18).FormulaA1 = $"=(SQRT($B$8)*$J{currentRow}*SIN((SQRT($G{currentRow})/SQRT($B$8))*$P{currentRow}))/SQRT($G{currentRow})";
                firstOriginalDataRow.Cell(19).FormulaA1 = $"=(SQRT($B$8)*$J{currentRow}*SIN((SQRT($G{currentRow})/SQRT($B$8))*$Q{currentRow}))/SQRT($G{currentRow})";
                firstOriginalDataRow.Cell(20).FormulaA1 = $"=(SQRT($B$8)*$J{currentRow}*SIN((SQRT($G{currentRow})/SQRT($B$8))*$R{currentRow}))/SQRT($G{currentRow})";
                firstOriginalDataRow.Cell(21).FormulaA1 = $"=$G{currentRow}";

                narmalize01DataRow.Cell(1).FormulaA1 =
                    $"=(original!T{currentRow}-original!$B$6)/(original!$C$6-original!$B$6)";
                narmalize01DataRow.Cell(2).FormulaA1 =
                    $"=(original!U{currentRow}-original!$B$6)/(original!$C$6-original!$B$6)";
                narmalize01DataRow.Cell(3).FormulaA1 =
                    $"=(original!V{currentRow}-original!$B$6)/(original!$C$6-original!$B$6)";
                narmalize01DataRow.Cell(4).FormulaA1 =
                    $"=(original!W{currentRow}-original!$B$6)/(original!$C$6-original!$B$6)";
                narmalize01DataRow.Cell(5).FormulaA1 =
                    $"=(original!X{currentRow}-original!$B$6)/(original!$C$6-original!$B$6)";
                narmalize01DataRow.Cell(6).FormulaA1 =
                    $"=(original!Y{currentRow}-original!$B$6)/(original!$C$6-original!$B$6)";
                narmalize01DataRow.Cell(7).FormulaA1 =
                    $"=(original!Z{currentRow}-original!$B$3)/(original!$C$3-original!$B$3)";


                narmalizeMin11DataRow.Cell(1).FormulaA1 = $"='norm (0;1)'!A{currentRow}*2-1";
                narmalizeMin11DataRow.Cell(2).FormulaA1 = $"='norm (0;1)'!B{currentRow}*2-1";
                narmalizeMin11DataRow.Cell(3).FormulaA1 = $"='norm (0;1)'!C{currentRow}*2-1";
                narmalizeMin11DataRow.Cell(4).FormulaA1 = $"='norm (0;1)'!D{currentRow}*2-1";
                narmalizeMin11DataRow.Cell(5).FormulaA1 = $"='norm (0;1)'!E{currentRow}*2-1";
                narmalizeMin11DataRow.Cell(6).FormulaA1 = $"='norm (0;1)'!F{currentRow}*2-1";
                narmalizeMin11DataRow.Cell(7).FormulaA1 = $"='norm (0;1)'!G{currentRow}*2-1";

                if (i % 50 == 0)
                {
                    iForT0 = 0;
                    k += kstep;
                }

                if (i % 2500 == 0)
                {
                    k = kStart;
                    v0 += v0step;
                }

                firstOriginalDataRow = firstOriginalDataRow.RowBelow();
                narmalize01DataRow = narmalize01DataRow.RowBelow();
                narmalizeMin11DataRow = narmalizeMin11DataRow.RowBelow();
            }

            workbook.SaveAs(memoryStream);
            memoryStream.Position = 0;

            return memoryStream;
        }

        private IXLRangeRow GetFirstDataRow(IXLWorksheet worksheet)
        {
            var firstStopRowUsed = worksheet.FirstRowUsed();
            var headerStopsRow = firstStopRowUsed.RowUsed();
            return headerStopsRow.RowBelow();
        }
    }
}