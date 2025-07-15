using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using VOR.Convert;
using VOR.Models;
using VOR.Models.LineProcess;
using VOR.View;

namespace VOR
{
    public class Main
    {
        /*public static readonly bool PNR_logic_var = true;*/
        public void CreateVOR(string filePath, List<SpecRow> specRows)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                Body body = mainPart.Document.Body;

                // Создаем таблицу
                Table table = new Table();

                // Задаем свойства таблицы
                TableProperties tableProperties = new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 6 },
                        new BottomBorder { Val = BorderValues.Single, Size = 6 },
                        new LeftBorder { Val = BorderValues.Single, Size = 6 },
                        new RightBorder { Val = BorderValues.Single, Size = 6 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 6 }
                        ),
                    new TableWidth { Width = ((1.8 + 12.5 + 1.75 + 2.0 + 4.0 + 4.4) * 567).ToString(), Type = TableWidthUnitValues.Dxa }, // Занять всю ширину страницы
                    new TableLayout { Type = TableLayoutValues.Fixed }
                    );
                table.Append(tableProperties);

                // Задаем ширину столбцов в twips, 1 см = 567 twips
                TableGrid tableGrid = new TableGrid(
                    new GridColumn() { Width = ((int)(1.8 * 567)).ToString() },
                    new GridColumn() { Width = ((int)(12.5 * 567)).ToString() },
                    new GridColumn() { Width = ((int)(1.75 * 567)).ToString() },
                    new GridColumn() { Width = ((int)(2.0 * 567)).ToString() },
                    new GridColumn() { Width = ((int)(4.0 * 567)).ToString() },
                    new GridColumn() { Width = ((int)(4.4 * 567)).ToString() }
                    );
                table.Append(tableGrid);

                // Добавляем нумерацию
                if (mainPart.NumberingDefinitionsPart == null)
                {
                    mainPart.AddNewPart<NumberingDefinitionsPart>();
                    mainPart.NumberingDefinitionsPart.Numbering = new Numbering();
                }

                AddMultilevelNumberingStyle(mainPart.NumberingDefinitionsPart);

                // Добавляем заголовки
                table.Append(CreateTableRowFromText(
                    "№ п/п", "Наименование работ", "Ед. изм.", "Количество",
                    "Ссылка на чертежи, спецификации", "Расчет объемов работ и расхода материалов", true));
                table.Append(CreateTableRowFromText(
                    "1", "2", "3", "4", "5", "6", true));
                table.Append(CreateTableRowFromText(
                    "", "Монтаж оборудования и материалов", "", "", "", "", true));

                var factory = new LineProcessorFactory();
                #region Регистрация обработчиков
                factory.RegisterProcessor(new AttributeProcessor());
                factory.RegisterProcessor(new HeaderProcessor());
                factory.RegisterProcessor(new InterBlockCablesProcessor());
                factory.RegisterProcessor(new CabelProcessor());
                factory.RegisterProcessor(new FLCabelProcessor());
                factory.RegisterProcessor(new KTPProcessor());
                factory.RegisterProcessor(new NKUProcessor());
                factory.RegisterProcessor(new PipeProcessor());
                factory.RegisterProcessor(new NakProcessor());
                factory.RegisterProcessor(new TubeProcessor());
                factory.RegisterProcessor(new GruntProcessor());
                factory.RegisterProcessor(new UgolokProcessor());
                factory.RegisterProcessor(new ListMKProcessor());
                factory.RegisterProcessor(new ShvellerProcessor());
                factory.RegisterProcessor(new ListPSH150Processor());
                factory.RegisterProcessor(new ListPSH250Processor());
                factory.RegisterProcessor(new SvetGolProcessor());
                factory.RegisterProcessor(new CommonProcessor());
                #endregion

                foreach (var row in specRows)
                {
                    var processor = factory.GetProcessor(row);
                    if (processor != null)
                    {
                        processor.Process(row, out List<VORRow> row_out);
                        foreach (var vrow in row_out)
                        {
                            table.Append(CreateTableRowFromVORRow(vrow));
                        }
                    }
                }

                if (MainWindow.pnr_list.Any())
                {
                    table.Append(CreateTableRowFromVORRow(new VORRow(
                    0, "Пуско-наладочные работы", "", "", "", "")));

                    PNRToVOR PNR = new PNRToVOR();
                    foreach (var pnr in MainWindow.pnr_list)
                    {
                        foreach (var vrow in PNR.Convert(pnr))
                        {
                            table.Append(CreateTableRowFromVORRow(vrow));
                        }
                    }
                }

                // Добавляем таблицу в конец документа
                body.Append(table);

                // Сохраняем изменения
                wordDoc.MainDocumentPart.Document.Save();

                if (CabelProcessor.NotFoundItems.Any() || InterBlockCablesProcessor.NotFoundItems.Any())
                {
                    var result = new ObservableCollection<InfoCableRow>(CabelProcessor.NotFoundItems.Concat(InterBlockCablesProcessor.NotFoundItems));
                    CablesWindow cablesWindow = new CablesWindow(result);
                    cablesWindow.ShowDialog();
                }
            }
        }

        private TableRow CreateTableRowFromVORRow(VORRow vORRow)
        {
            TableRow row = new TableRow();

            // Создание первой ячейки с автоматической нумерацией
            TableCell cell = new TableCell();
            cell.Append(CreateNumberedParagraph("", vORRow.NumberLevel, 1));
            row.Append(cell);

            // Создание второй ячейки с наименованием работы
            cell = new TableCell();

            foreach (var text in vORRow.Name.Split('\n'))
            {
                Run run = new Run(new Text(text));

                if (vORRow.NumberLevel == 0 || vORRow.NumberLevel == 1)
                {
                    RunProperties runProperties = new RunProperties();
                    runProperties.Append(new Bold());
                    run.RunProperties = runProperties;
                }

                Paragraph paragraph = new Paragraph(run);
                cell.Append(paragraph);
            }
            row.Append(cell);

            // Создание третьей ячейки с единицей измерения
            cell = new TableCell(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center }),
                    new Run(new Text(vORRow.Unit))
                    )
                );
            row.Append(cell);

            // Создание четвертой ячейки с количеством
            cell = new TableCell(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center }),
                    new Run(new Text(vORRow.Quantity))
                    )
                );
            row.Append(cell);

            // Создание пятой ячейки со ссылкой на РД
            cell = new TableCell(
                new Paragraph(
                    new Run(
                        new RunProperties(
                            new FontSize() { Val = "20" }),
                        new Text(vORRow.Link))
                    )
                );
            row.Append(cell);

            // Создание шестой ячейки с формулой для расчета
            cell = new TableCell(
                new Paragraph(
                    new Run(
                        new RunProperties(
                            new FontSize() { Val = "20" }),
                        new Text(vORRow.Calculation))
                    )
                );
            row.Append(cell);
            
            return row;
        }

        private TableRow CreateTableRowFromText(string number, string name, string unit, string quantity, string link, string calculation, bool bold)
        {
            TableRow row = new TableRow();
            List<string> param = new List<string>() { number, name, unit, quantity, link, calculation};

            foreach (var element in param)
            {
                TableCell cell = new TableCell();

                Run run = new Run(new Text(element));
                if (bold)
                {
                    RunProperties runProperties = new RunProperties();
                    runProperties.Append(new Bold());
                    run.RunProperties = runProperties;
                }

                Paragraph paragraph = new Paragraph(run);

                paragraph.ParagraphProperties = new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center });

                cell.Append(paragraph);
                row.Append(cell);
            }
            return row;
        }

        private void AddMultilevelNumberingStyle(NumberingDefinitionsPart numberingPart)
        {
            var element = new Numbering(
                new AbstractNum(
                    new Level(
                        new StartNumberingValue() { Val = 1 },
                        new NumberingFormat() { Val = NumberFormatValues.Decimal },
                        new LevelText() { Val = "%1" },
                        new LevelJustification() { Val = LevelJustificationValues.Left },
                        new ParagraphProperties(
                            new Indentation() { Left = "0", Hanging = "0" }
                            ),
                        new RunProperties(
                            new FontSize() { Val = "20" })
                        )
                    { LevelIndex = 0 },

                    new Level(
                        new StartNumberingValue() { Val = 1 },
                        new NumberingFormat() { Val = NumberFormatValues.Decimal },
                        new LevelText() { Val = "%1.%2" },
                        new LevelJustification() { Val = LevelJustificationValues.Left },
                        new ParagraphProperties(
                            new Indentation() { Left = "0", Hanging = "0" }
                            ),
                        new RunProperties(
                            new FontSize() { Val = "20" })
                        )
                    { LevelIndex = 1 },

                    new Level(
                        new StartNumberingValue() { Val = 1 },
                        new NumberingFormat() { Val = NumberFormatValues.Decimal },
                        new LevelText() { Val = "%1.%2.%3" },
                        new LevelJustification() { Val = LevelJustificationValues.Left },
                        new ParagraphProperties(
                            new Indentation() { Left = "0", Hanging = "0" }
                            ),
                        new RunProperties(
                            new FontSize() { Val = "20" })
                        )
                    { LevelIndex = 2 },

                    new Level(
                        new StartNumberingValue() { Val = 1 },
                        new NumberingFormat() { Val = NumberFormatValues.Decimal },
                        new LevelText() { Val = "%1.%2.%3.%4" },
                        new LevelJustification() { Val = LevelJustificationValues.Left },
                        new ParagraphProperties(
                            new Indentation() { Left = "0", Hanging = "0" }
                            ),
                        new RunProperties(
                            new FontSize() { Val = "20" })
                        )
                    { LevelIndex = 3 }
                    )
                { AbstractNumberId = 1 },
                
                new NumberingInstance(
                    new AbstractNumId() { Val = 1 }
                    )
                { NumberID = 1 }
                );

            numberingPart.Numbering.Append(element);
        }

        private Paragraph CreateNumberedParagraph(string text, int level, int numberingId)
        {
            return new Paragraph(
                new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference() { Val = level },
                        new NumberingId() { Val = numberingId }
                        )
                    ),
                new Run(
                    new Text(text))
                );
        }
    }
}
