import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class Main {

    private static final String[] titles = {
            "№", "Место отправки", "№ ф103", "Вид отправления", "Дата сдачи\nотправления",
            "ШПИ", "Стоимость за СМС-\nуведомление, (руб)", "Тариф за отправление,\n(руб)/Тариф и о/ц с НДС + тариф СМС", "Масса, (кг)"
    };

    private Main() {
    }


    public static void main(String[] args) throws FileNotFoundException, IOException, ParseException {

        String excelFilePath = "C:/Users/А/Desktop/Programming/ExcelFiles/Reports/MainReport.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);

        XSSFWorkbook workbookForRead = new XSSFWorkbook(inputStream);
        XSSFSheet sheetReadBook = workbookForRead.getSheetAt(0);

        int rows = sheetReadBook.getLastRowNum();

        ArrayList<DataRow> rowsData = new ArrayList<>();

        for (int i = 1; i <= rows; i++) {

            String agent = sheetReadBook.getRow(i).getCell(0).getStringCellValue();

            String signatory =  sheetReadBook.getRow(i).getCell(1).getStringCellValue();

            String numberContract = sheetReadBook.getRow(i).getCell(2). getStringCellValue();

            String sender = sheetReadBook.getRow(i).getCell(4).getStringCellValue();

            DataFormatter dataFormatter = new DataFormatter();
            String dateSource = dataFormatter.formatCellValue(sheetReadBook.getRow(i).getCell(7));
            String dateReplaceSlashToFullStop = dateSource.replace("/",".");
            SimpleDateFormat oldFormat = new SimpleDateFormat("MM.dd.yy");
            SimpleDateFormat newFormat = new SimpleDateFormat("dd.MM.yy");
            Date date = oldFormat.parse(dateReplaceSlashToFullStop);
            String dateNew = newFormat.format(date);

            String part = sheetReadBook.getRow(i).getCell(9).getStringCellValue();

            String mailType = sheetReadBook.getRow(i).getCell(10).getStringCellValue();

            String id = sheetReadBook.getRow(i).getCell(11).getStringCellValue();

            double priceNotification = sheetReadBook.getRow(i).getCell(13).getNumericCellValue();

            double rate = sheetReadBook.getRow(i).getCell(16).getNumericCellValue();

            double weight = sheetReadBook.getRow(i).getCell(17).getNumericCellValue();

            rowsData.add(new DataRow(agent, signatory, numberContract, sender, dateNew, part, mailType, id, priceNotification, rate, weight));

        }




        Map<String, ArrayList<Operation>> dataMap = new HashMap<>();
        for (DataRow dataRow : rowsData) {
            Operation operation = new Operation();

            operation.agent = dataRow.agent;
            operation.signatory = dataRow.signatory;
            operation.numberContract = dataRow.numberContract;
            operation.sender = dataRow.sender;
            operation.date = dataRow.date;
            operation.part = dataRow.part;
            operation.mailType = dataRow.mailType;
            operation.id = dataRow.id;
            operation.priceNotification = dataRow.priceNotification;
            operation.rate = dataRow.rate;
            operation.weight = dataRow.weight;

            if (!dataMap.containsKey(dataRow.agent)) {
                dataMap.put(dataRow.agent, new ArrayList<>());
            }

            dataMap.get(dataRow.agent).add(operation);
        }

        //////////////////////// all in map


        Workbook workbookForRecord;
        if (args.length > 0 && args[0].equals("-xls")) workbookForRecord = new HSSFWorkbook();
        else workbookForRecord = new XSSFWorkbook();

        Map<String, CellStyle> styles = createStyles(workbookForRecord);

        Set<String> keys = dataMap.keySet();
        String[] nameOfSheets = keys.toArray(new String[keys.size()]);


        for (int sender = 0; sender < dataMap.size(); sender++) {

            ArrayList<Operation> everyOperations = dataMap.get(nameOfSheets[sender]);

            String[][] table = new String[everyOperations.size()][titles.length];

            for (int row = 0, order = 1; row < everyOperations.size(); row++) {
                table[row][0] = Integer.toString(order++);
                table[row][1] = everyOperations.get(row).sender;
                table[row][2] = everyOperations.get(row).part;
                table[row][3] = everyOperations.get(row).mailType;
                table[row][4] = everyOperations.get(row).date;
                table[row][5] = everyOperations.get(row).id;
                table[row][6] = Double.toString(everyOperations.get(row).priceNotification);
                table[row][7] = Double.toString(everyOperations.get(row).rate);
                table[row][8] = Double.toString(everyOperations.get(row).weight);
            }

            DecimalFormat dcf = new DecimalFormat("###.##");
            int countSends = 0;
            /// Сумма за СМС
            double summaryNotification = 0;
            for (int i = 0; i < table.length; i++) {
                summaryNotification += Double.parseDouble(table[i][6]);
                countSends++;
            }

            /// Сумма за тариф + СМС
            double summaryRate = 0;
            for (int i = 0; i < table.length; i++) {
                summaryRate += Double.parseDouble(table[i][7]);
            }
            summaryRate = Math.round(summaryRate * 100.0) / 100.0;


            /// Сумма вес
            double summaryWeight = 0;
            for (int i = 0; i < table.length; i++) {
                summaryWeight += Double.parseDouble(table[i][8]);
            }

            /// Сумма без НДС
            double rateNoNDS = (summaryRate - summaryNotification) / 1.2;

            /// Агентское вознаграждение
            double totalAgentReward = rateNoNDS * 0.15;
            totalAgentReward = Math.round(totalAgentReward * 100.0) / 100.0;


            Sheet sheet = workbookForRecord.createSheet(nameOfSheets[sender]);
            // delete all lines
            sheet.setPrintGridlines(false);
            sheet.setDisplayGridlines(false);

            PrintSetup printSetup = sheet.getPrintSetup();
            printSetup.setLandscape(true);
            sheet.setFitToPage(true);
            sheet.setHorizontallyCenter(true);

            sheet.setZoom(85);

            // create rows
            Row row1 = sheet.createRow(0);
            row1.setHeightInPoints((float) 5.3);

            Row row2 = sheet.createRow(1);
            row2.setHeightInPoints(15);
            String[] row2Words = {"Отчет агента ", "к Договору №" + everyOperations.get(0).numberContract, everyOperations.get(0).agent, "", "от 30.06.23", "за Июнь"};
            for (int i = 1, j = 0; i <= row2Words.length; i++, j++) {
                Cell cell = row2.createCell(i);
                cell.setCellValue(row2Words[j]);
                cell.setCellStyle(styles.get("bold"));
            }
            sheet.addMergedRegion(CellRangeAddress.valueOf("$D$2:$E$2"));


            Row row3 = sheet.createRow(2);
            row3.setHeightInPoints((float) 9.8);
            for (int i = 1; i < 11; i++) {
                row3.createCell(i).setCellStyle(styles.get("underline"));
            }

            Row row4 = sheet.createRow(3);
            row4.setHeightInPoints((float) 4.5);

            Row row5 = sheet.createRow(4);
            row5.setHeightInPoints((float) 14.4);
            Cell principal = row5.createCell(1);
            principal.setCellValue("Принципал");
            principal.setCellStyle(styles.get("bold"));
            Cell principalDetails = row5.createCell(2);
            principalDetails.setCellValue("ООО «Почта ЕКОМ», ИНН 9729310818, 119454, г. Москва, проспект Вернадского, д.18, этаж 2, пом.23, тел. +79265731897");
            sheet.addMergedRegion(CellRangeAddress.valueOf("$C$5:$L$5"));

            Row row6 = sheet.createRow(5);
            row6.setHeightInPoints((float) 4.5);

            String agentDet = everyOperations.get(0).sender;
            Row row7 = sheet.createRow(6);
            row7.setHeightInPoints((float) 14.4);
            Cell agent = row7.createCell(1);
            agent.setCellValue("Агент");
            agent.setCellStyle(styles.get("bold"));
            Cell agentDetails = row7.createCell(2);
            agentDetails.setCellValue(agentDet);
            sheet.addMergedRegion(CellRangeAddress.valueOf("$C$7:$L$7"));

            Row row8 = sheet.createRow(7);
            row8.setHeightInPoints(4.5F);

            Row row9 = sheet.createRow(8);
            row9.setHeightInPoints(42);
            for (int i = 0, j = 1; i < titles.length; i++, j++) {
                Cell cell = row9.createCell(j);
                cell.setCellValue(titles[i]);
                cell.setCellStyle(styles.get("title"));
            }

            int rowNum = 9;

            for (int i = 0; i < table.length; i++) {
                Row row = sheet.createRow(rowNum);
                rowNum++;
                for (int j = 0, k = 1; k <= table[0].length; j++, k++) {
                    Cell cell = row.createCell(k);
                    cell.setCellValue(table[i][j]);
                    cell.setCellStyle(styles.get("title"));
                }
            }

            ////// первая строка после таблицы.
            Row rowNextAfterTable1 = sheet.createRow(rowNum);
            rowNextAfterTable1.setHeightInPoints(15);
            Cell cellTotal = rowNextAfterTable1.createCell(1);
            cellTotal.setCellValue("Всего за отчетный период:");
            cellTotal.setCellStyle(styles.get("total"));
            Cell cellSumNotification = rowNextAfterTable1.createCell(7);
            cellSumNotification.setCellValue(dcf.format(summaryNotification) + " ₽");
            cellSumNotification.setCellStyle(styles.get("totalNumbers"));
            Cell cellSumRate = rowNextAfterTable1.createCell(8);
            cellSumRate.setCellValue(dcf.format(summaryRate) + " ₽");
            cellSumRate.setCellStyle(styles.get("totalNumbers"));
            Cell cellSumWeight = rowNextAfterTable1.createCell(9);
            cellSumWeight.setCellValue(dcf.format(summaryWeight) + " кг");
            cellSumWeight.setCellStyle(styles.get("totalNumbers"));
            for (int i = 2; i < 7; i++) {
                rowNextAfterTable1.createCell(i).setCellStyle(styles.get("top-und"));
            }
            int rowNextAfterTableNum1 = rowNum + 1;
            sheet.addMergedRegion(CellRangeAddress.valueOf("$B$" + rowNextAfterTableNum1 + ":$G$" + rowNextAfterTableNum1));

            /////// вторая строка после таблицы
            Row rowNextAfterTable2 = sheet.createRow(rowNum + 1);
            rowNextAfterTable2.setHeightInPoints(15);
            Cell cellMoneyGet = rowNextAfterTable2.createCell(1);
            cellMoneyGet.setCellValue("Итого средств принято:");
            cellMoneyGet.setCellStyle(styles.get("total"));
            Cell cellSumMoneyGet = rowNextAfterTable2.createCell(6);
            cellSumMoneyGet.setCellValue(dcf.format(summaryRate) + " ₽");
            cellSumMoneyGet.setCellStyle(styles.get("totalNumbers"));
            for (int i = 2; i < 6; i++) {
                rowNextAfterTable2.createCell(i).setCellStyle(styles.get("top-und"));
            }
            int rowNextAfterTableNum2 = rowNum + 2;
            sheet.addMergedRegion(CellRangeAddress.valueOf("$B$" + rowNextAfterTableNum2 + ":$F$" + rowNextAfterTableNum2));

            //////// третья строка после таблицы

            Row rowNextAfterTable3 = sheet.createRow(rowNum + 2);
            rowNextAfterTable3.setHeightInPoints(15);
            Cell cellSendToPrincipal = rowNextAfterTable3.createCell(1);
            cellSendToPrincipal.setCellValue("Перечислено на счет принципала:");
            cellSendToPrincipal.setCellStyle(styles.get("total"));
            Cell cellMoneyToSendPrincipal = rowNextAfterTable3.createCell(6);
            cellMoneyToSendPrincipal.setCellStyle(styles.get("totalNumbers"));
            for (int i = 2; i < 6; i++) {
                rowNextAfterTable3.createCell(i).setCellStyle(styles.get("top-und"));
            }
            int rowNextAfterTableNum3 = rowNum + 3;
            sheet.addMergedRegion(CellRangeAddress.valueOf("$B$" + rowNextAfterTableNum3 + ":$F$" + rowNextAfterTableNum3));

            /////// четвертая строка после таблицы
            Row rowNextAfterTable4 = sheet.createRow(rowNum + 3);
            rowNextAfterTable4.setHeightInPoints(15);
            Cell cellTotalNoNDS = rowNextAfterTable4.createCell(1);
            cellTotalNoNDS.setCellValue("Итого средств принято без НДС (20%):");
            cellTotalNoNDS.setCellStyle(styles.get("total"));
            Cell cellSumNoNds = rowNextAfterTable4.createCell(6);
            cellSumNoNds.setCellValue(dcf.format(rateNoNDS) + " ₽");
            cellSumNoNds.setCellStyle(styles.get("totalNumbers"));
            for (int i = 2; i < 6; i++) {
                rowNextAfterTable4.createCell(i).setCellStyle(styles.get("top-und"));
            }
            int rowNextAfterTableNum4 = rowNum + 4;
            sheet.addMergedRegion(CellRangeAddress.valueOf("$B$" + rowNextAfterTableNum4 + ":$F$" + rowNextAfterTableNum4));

            /////// пятая строка после таблицы
            Row rowNextAfterTable5 = sheet.createRow(rowNum + 4);
            rowNextAfterTable5.setHeightInPoints(15);
            Cell cellAgentReward = rowNextAfterTable5.createCell(1);
            cellAgentReward.setCellValue("Вознаграждение агента составило: ");
            cellAgentReward.setCellStyle(styles.get("total"));
            Cell cellTotalAgentReward = rowNextAfterTable5.createCell(3);
            cellTotalAgentReward.setCellValue(dcf.format(totalAgentReward) + " ₽");
            cellTotalAgentReward.setCellStyle(styles.get("totalNumbers"));
            Cell cellPercent = rowNextAfterTable5.createCell(4);
            cellPercent.setCellValue("15 %");
            cellPercent.setCellStyle(styles.get("totalNumbers"));

            for (int i = 2; i < 3; i++) {
                rowNextAfterTable5.createCell(i).setCellStyle(styles.get("top-und"));
            }
            int rowNextAfterTableNum5 = rowNum + 5;
            sheet.addMergedRegion(CellRangeAddress.valueOf("$B$" + rowNextAfterTableNum5 + ":$C$" + rowNextAfterTableNum5));

            /// шестая строка после таблицы
            Row rowNextAfterTable6 = sheet.createRow(rowNum + 5);
            rowNextAfterTable6.setHeightInPoints(7);

            ///седьмая строка после таблицы

            Row rowNextAfterTable7 = sheet.createRow(rowNum + 6);
            rowNextAfterTable7.setHeightInPoints(15);
            Cell cellAllNames = rowNextAfterTable7.createCell(1);
            cellAllNames.setCellValue("Всего наименовений:");
            cellAllNames.setCellStyle(styles.get("total"));
            Cell cellCountSends = rowNextAfterTable7.createCell(3);
            cellCountSends.setCellValue(countSends);
            cellCountSends.setCellStyle(styles.get("totalNumbers"));
            Cell cellSummary = rowNextAfterTable7.createCell(4);
            cellSummary.setCellValue("На сумму");
            cellSummary.setCellStyle(styles.get("total"));
            Cell cellSumAllGetsPackages = rowNextAfterTable7.createCell(5);
            cellSumAllGetsPackages.setCellValue(dcf.format(summaryRate) + " ₽");
            cellSumAllGetsPackages.setCellStyle(styles.get("totalNumbers"));
            for (int i = 2; i < 3; i++) {
                rowNextAfterTable7.createCell(i).setCellStyle(styles.get("top-und"));
            }
            int rowNextAfterTableNum7 = rowNum + 7;
            sheet.addMergedRegion(CellRangeAddress.valueOf("$B$" + rowNextAfterTableNum7 + ":$C$" + rowNextAfterTableNum7));

            /// восьмая строка после таблицы
            Money summaryRateMoney = new Money(summaryRate);
            Row rowNextAfterTable8 = sheet.createRow(rowNum + 7);
            rowNextAfterTable8.setHeightInPoints((float) 14.4);
            Cell cellSummaryRateMoney = rowNextAfterTable8.createCell(1);
            cellSummaryRateMoney.setCellValue(firstUpperCase(summaryRateMoney.num2str()));
            cellSummaryRateMoney.setCellStyle(styles.get("bold"));
            int rowNextAfterTableNum8 = rowNum + 8;
            sheet.addMergedRegion(CellRangeAddress.valueOf("$B$" + rowNextAfterTableNum8 + ":$F$" + rowNextAfterTableNum8));

            /// девятая строка после таблицы
            Row rowNextAfterTable9 = sheet.createRow(rowNum + 8);
            rowNextAfterTable9.setHeightInPoints(7);

            // десятая строка после таблицы
            Money summaryAgentReward = new Money(totalAgentReward);
            Row rowNextAfterTable10 = sheet.createRow(rowNum + 9);
            rowNextAfterTable10.setHeightInPoints((float) 14.4);
            Cell cellSumAgentReward = rowNextAfterTable10.createCell(1);
            cellSumAgentReward.setCellValue("Следует к перечислению "+firstUpperCase(summaryAgentReward.num2str()) + ", НДС не облагается");
            cellSumAgentReward.setCellStyle(styles.get("bold"));
            int rowNextAfterTableNum10 = rowNum + 10;
            sheet.addMergedRegion(CellRangeAddress.valueOf("$B$" + rowNextAfterTableNum10 + ":$F$" + rowNextAfterTableNum10));

            // одиннадцатая строка после таблицы
            Row rowNextAfterTable11 = sheet.createRow(rowNum + 10);
            rowNextAfterTable11.setHeightInPoints(18);
            Cell cellLastText = rowNextAfterTable11.createCell(1);
            cellLastText.setCellValue("Вышеперечисленные   услуги   оказаны  полностью  и  в  срок.  Стороны претензий  по  объему,  качеству и срокам оказания услуг не имеют.");
            cellLastText.setCellStyle(styles.get("bold"));
            int rowNextAfterTableNum11 = rowNum + 11;
            sheet.addMergedRegion(CellRangeAddress.valueOf("$B$" + rowNextAfterTableNum11 + ":$L$" + rowNextAfterTableNum11));

            /// двенадцатая строка после таблицы
            Row rowNextAfterTable12 = sheet.createRow(rowNum + 11);
            rowNextAfterTable12.setHeightInPoints((float) 11.3);
            rowNextAfterTable12.setHeightInPoints((float) 9.8);
            for (int i = 1; i < 11; i++) {
                rowNextAfterTable12.createCell(i).setCellStyle(styles.get("underline"));
            }
            /// Тринадцатая строка после таблицы

            Row rowNextAfterTable13 = sheet.createRow(rowNum + 12);
            rowNextAfterTable13.setHeightInPoints(9.3F);

            //// Четырнадцатая строка после таблицы
            Row rowNextAfterTable14 = sheet.createRow(rowNum + 13);
            rowNextAfterTable14.setHeightInPoints(13);
            Cell agentBottom = rowNextAfterTable14.createCell(1);
            agentBottom.setCellValue("Агент");
            agentBottom.setCellStyle(styles.get("defaultCenter"));
            Cell agentSignature = rowNextAfterTable14.createCell(2);
            agentSignature.setCellValue("_________________________");
            Cell signatory = rowNextAfterTable14.createCell(3);
            signatory.setCellValue(everyOperations.get(0).signatory);
            Cell principalBottom = rowNextAfterTable14.createCell(7);
            principalBottom.setCellValue("Принципал");
            principalBottom.setCellStyle(styles.get("defaultCenter"));
            Cell principalSignature = rowNextAfterTable14.createCell(8);
            principalSignature.setCellValue("_____________________");
            Cell director = rowNextAfterTable14.createCell(9);
            director.setCellValue("Исполнительный директор");
            director.setCellStyle(styles.get("defaultCenter"));
            Cell directorName =  rowNextAfterTable14.createCell(11);
            directorName.setCellValue("Графчиков М.Л.");

            int rowNextAfterTableNum14 = rowNum + 14;
            sheet.addMergedRegion(CellRangeAddress.valueOf("$D$" + rowNextAfterTableNum14 + ":$E$" + rowNextAfterTableNum14));
            sheet.addMergedRegion(CellRangeAddress.valueOf("$J$" + rowNextAfterTableNum14 + ":$K$" + rowNextAfterTableNum14));




            //columnWidth
            sheet.setColumnWidth(0, 1 * 256);
            sheet.setColumnWidth(1, 15 * 256);
            sheet.setColumnWidth(2, 34 * 256);
            sheet.setColumnWidth(3, 16 * 256);
            sheet.setColumnWidth(4, 19 * 256);
            sheet.setColumnWidth(5, 14 * 256);
            sheet.setColumnWidth(6, 17 * 256);
            sheet.setColumnWidth(7, 15 * 256);
            sheet.setColumnWidth(8, 19 * 256);
            sheet.setColumnWidth(9, 19 * 256);
            sheet.setColumnWidth(11, 19 * 256);


            String file = "KhorkinBotActs.xls";
            if (workbookForRecord instanceof XSSFWorkbook) file += "x";
            FileOutputStream fileOut = new FileOutputStream("C:/Users/А/Desktop/Programming/ExcelFiles/Reports/" + file);
            workbookForRecord.write(fileOut);
            fileOut.close();
        }

        System.out.println("Готово");

    }


    private static Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<>();
        CellStyle style;
        Font boldFont = wb.createFont();
        boldFont.setBold(true);
        style = wb.createCellStyle();
        style.setFont(boldFont);
        styles.put("bold", style);

        style = wb.createCellStyle();
        style.setBorderBottom(BorderStyle.MEDIUM);
        styles.put("underline", style);

        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 10);
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setFont(titleFont);
        style.setWrapText(true);
        styles.put("title", style);


        style = wb.createCellStyle();
        style.setFont(boldFont);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderBottom(BorderStyle.MEDIUM);
        styles.put("total", style);

        style = wb.createCellStyle();
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderBottom(BorderStyle.MEDIUM);
        styles.put("top-und", style);

        style = wb.createCellStyle();
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        styles.put("totalNumbers", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        styles.put("defaultCenter", style);



        return styles;
    }

    public static String firstUpperCase(String word) {
        if(word == null || word.isEmpty()) return ""; //или return word;
        return word.substring(0, 1).toUpperCase() + word.substring(1);
    }

}


