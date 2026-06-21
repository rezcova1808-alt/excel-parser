package com.example.demo;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveOutputStream;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Base64;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;

public class UltimateExcelGenerator {

    public static void main(String[] args) throws Exception {
        // 1. Создаём базовый файл через POI
        Path tempFile = Files.createTempFile("temp_excel", ".xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            // ----- Лист DataTypes -----
            XSSFSheet dataSheet = wb.createSheet("DataTypes");
            fillDataSheet(dataSheet, wb);
            // ----- Лист MergeAndHide -----
            XSSFSheet mergeSheet = wb.createSheet("MergeAndHide");
            fillMergeSheet(mergeSheet, wb);
            // ----- Лист CondFormatting -----
            XSSFSheet condSheet = wb.createSheet("CondFormatting");
            fillCondFormatting(condSheet, wb);
            // ----- Лист Table -----
            XSSFSheet tableSheet = wb.createSheet("Table");
            fillTable(tableSheet, wb);
            // ----- Лист Pivot -----
            XSSFSheet pivotSheet = wb.createSheet("Pivot");
            fillPivot(pivotSheet, wb);
            // ----- Лист Picture -----
            XSSFSheet pictureSheet = wb.createSheet("Picture");
            addPicture(pictureSheet, wb);
            // ----- Именованные диапазоны -----
            XSSFName namedRange1 = wb.createName();
            namedRange1.setNameName("MyNamedRange");
            namedRange1.setRefersToFormula("DataTypes!$B$2:$B$5");

            // Ещё один диапазон, ссылающийся на первую строку
            XSSFName namedRange2 = wb.createName();
            namedRange2.setNameName("HeaderRange");
            namedRange2.setRefersToFormula("DataTypes!$1:$1");

            try (FileOutputStream fos = new FileOutputStream(tempFile.toFile())) {
                wb.write(fos);
            }
        }

        // 2. Модифицируем XML внутри ZIP – добавляем CDATA, комментарии, PI
        Path finalFile = Paths.get("UltimateTest.xlsx");
        addRareXmlConstructs(tempFile, finalFile);

        // 3. Удаляем временный файл
        Files.deleteIfExists(tempFile);
        System.out.println("Файл UltimateTest.xlsx успешно создан!");
    }

    private static void fillDataSheet(XSSFSheet sheet, XSSFWorkbook wb) {
        // Стили
        CellStyle dateStyle = wb.createCellStyle();
        dateStyle.setDataFormat(wb.createDataFormat().getFormat("dd.MM.yyyy"));
        CellStyle percentStyle = wb.createCellStyle();
        percentStyle.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
        CellStyle currencyStyle = wb.createCellStyle();
        currencyStyle.setDataFormat(wb.createDataFormat().getFormat("#,##0.00 ₽"));

        Row header = sheet.createRow(0);
        String[] headers = {"Тип", "Значение", "Комментарий", "Гиперссылка", "Валидация"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(createHeaderStyle(wb));
        }

        int rowNum = 1;
        rowNum = addRow(sheet, rowNum, "Текст со спецсимволами", "<Тег> & значение \"кавычки'", null, null);
        rowNum = addRow(sheet, rowNum, "Целое число", 123456, null, null);
        rowNum = addRow(sheet, rowNum, "Дробное число", 123.456, null, null);
        rowNum = addRow(sheet, rowNum, "Отрицательное", -99.99, null, null);
        // Дата
        Row row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Дата");
        Cell dateCell = row.createCell(1);
        dateCell.setCellValue(java.time.LocalDate.of(2025, 3, 15));
        dateCell.setCellStyle(dateStyle);
        addComment(wb, row, 1, "Это дата");
        // Процент
        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Процент");
        Cell percentCell = row.createCell(1);
        percentCell.setCellValue(0.1234);
        percentCell.setCellStyle(percentStyle);
        // Валюта
        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Валюта");
        Cell currencyCell = row.createCell(1);
        currencyCell.setCellValue(1234.56);
        currencyCell.setCellStyle(currencyStyle);
        // Логические
        rowNum = addRow(sheet, rowNum, "Логическое TRUE", true, null, null);
        rowNum = addRow(sheet, rowNum, "Логическое FALSE", false, null, null);
        // Ошибка формулой
        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Ошибка #DIV/0!");
        row.createCell(1).setCellFormula("1/0");
        // Формула SUM
        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Формула SUM");
        row.createCell(1).setCellFormula("SUM(C4:C5)");
        // Формула CONCATENATE
        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Формула CONCAT");
        row.createCell(1).setCellFormula("CONCATENATE(B2, \" \", B3)");
        // Гиперссылки
        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Гиперссылка URL");
        Hyperlink urlLink = wb.getCreationHelper().createHyperlink(HyperlinkType.URL);
        urlLink.setAddress("http://example.com");
        row.createCell(3).setCellValue("Нажми меня");
        row.getCell(3).setHyperlink(urlLink);
        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Ссылка на лист");
        Hyperlink docLink = wb.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
        docLink.setAddress("'MergeAndHide'!A1");
        row.createCell(3).setCellValue("Перейти");
        row.getCell(3).setHyperlink(docLink);

        // Валидация (выпадающий список) на столбце E
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = dvHelper.createExplicitListConstraint(new String[]{"Да", "Нет", "Возможно"});
        CellRangeAddressList addressList = new CellRangeAddressList(1, rowNum, 4, 4);
        DataValidation validation = dvHelper.createValidation(constraint, addressList);
        validation.setShowErrorBox(true);
        validation.createErrorBox("Ошибка", "Выберите из списка");
        sheet.addValidationData(validation);
    }

    private static void fillMergeSheet(XSSFSheet sheet, XSSFWorkbook wb) {
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("Объединённый заголовок");
        titleCell.setCellStyle(createHeaderStyle(wb));
        // Скрытая строка
        Row hiddenRow = sheet.createRow(5);
        hiddenRow.setZeroHeight(true);
        hiddenRow.createCell(0).setCellValue("Эта строка скрыта");
        // Скрытый столбец (C)
        sheet.setColumnHidden(2, true);
        Row colRow = sheet.createRow(2);
        colRow.createCell(2).setCellValue("Невидимый столбец");
        // Комментарий
        Row commentRow = sheet.createRow(10);
        commentRow.createCell(0).setCellValue("Ячейка с комментарием");
        addComment(wb, commentRow, 0, "Это программный комментарий");
    }

    private static void fillCondFormatting(XSSFSheet sheet, XSSFWorkbook wb) {
        for (int i = 0; i < 10; i++) {
            Row r = sheet.createRow(i);
            Cell c = r.createCell(0);
            c.setCellValue(i * 10);
        }
        XSSFSheetConditionalFormatting condFmt = sheet.getSheetConditionalFormatting();
        XSSFConditionalFormattingRule rule = condFmt.createConditionalFormattingRule(ComparisonOperator.GT, "50");
        XSSFFontFormatting fontFmt = rule.createFontFormatting();
        fontFmt.setFontStyle(true, false);
        fontFmt.setFontColor(new XSSFColor(new byte[]{(byte) 255, 0, 0}, null));
        XSSFPatternFormatting patternFmt = rule.createPatternFormatting();
        patternFmt.setFillBackgroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        condFmt.addConditionalFormatting(new CellRangeAddress[]{CellRangeAddress.valueOf("A1:A10")}, rule);
    }

    private static void fillTable(XSSFSheet sheet, XSSFWorkbook wb) {
        String[] headers = {"ID", "Название", "Цена"};
        Row header = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            header.createCell(i).setCellValue(headers[i]);
        }
        Object[][] data = {{1, "Товар А", 100}, {2, "Товар Б", 250}, {3, "Товар В", 75}};
        int rowNum = 1;
        for (Object[] rowData : data) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue((Integer) rowData[0]);
            row.createCell(1).setCellValue((String) rowData[1]);
            row.createCell(2).setCellValue((Integer) rowData[2]);
        }

        // 1. Создаём таблицу
        XSSFTable table = sheet.createTable(null);
        table.setName("TestTable");
        table.setDisplayName("TestTable");

        // 2. Задаём диапазон таблицы (включая заголовок)
        CellRangeAddress range = new CellRangeAddress(0, data.length, 0, headers.length - 1);
        // Параметр false означает, что диапазон не будет включать строку итогов
        String ref = "A1:" + new CellReference(data.length, headers.length - 1).formatAsString();
        AreaReference area = new AreaReference(ref, SpreadsheetVersion.EXCEL2007);
        table.setArea(area);

        // 3. Настройка стиля таблицы через CTTableStyleInfo
        CTTable cttable = table.getCTTable();
        CTTableStyleInfo styleInfo = cttable.addNewTableStyleInfo();

        // Выбираем любой понравившийся встроенный стиль, например "TableStyleMedium2"
        styleInfo.setName("TableStyleMedium2");
        // Включаем чередование строк
        styleInfo.setShowRowStripes(true);
        // Можно включить или выключить чередование столбцов
        styleInfo.setShowColumnStripes(false);
        // Показывать первую колонку
        styleInfo.setShowFirstColumn(false);
        // Показывать последнюю колонку
        styleInfo.setShowLastColumn(false);
    }

    private static void fillPivot(XSSFSheet sheet, XSSFWorkbook wb) {
        XSSFSheet sourceSheet = wb.getSheet("Table");
        if (sourceSheet == null) return;

        int lastRow = sourceSheet.getLastRowNum();
        int lastCol = sourceSheet.getRow(0).getLastCellNum() - 1;

        AreaReference sourceArea = new AreaReference(
                new CellReference(0, 0),
                new CellReference(lastRow, lastCol),
                SpreadsheetVersion.EXCEL2007
        );

        XSSFPivotTable pivot = sheet.createPivotTable(sourceArea, new CellReference("A3"), sourceSheet);
        pivot.addRowLabel(0);
        pivot.addRowLabel(1);
        pivot.addColumnLabel(DataConsolidateFunction.SUM, 2, "Сумма цен");
    }
    private static void addPicture(XSSFSheet sheet, XSSFWorkbook wb) throws IOException {
        // 1x1 прозрачный PNG
        byte[] imgData = Base64.getDecoder().decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");
        int pictureIdx = wb.addPicture(imgData, Workbook.PICTURE_TYPE_PNG);
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 2, 2, 4, 6);
        drawing.createPicture(anchor, pictureIdx);
    }

    private static CellStyle createHeaderStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static void addComment(Workbook wb, Row row, int col, String text) {
        if (!(wb instanceof XSSFWorkbook)) return;
        XSSFWorkbook xwb = (XSSFWorkbook) wb;
        XSSFSheet sheet = (XSSFSheet) row.getSheet();
        XSSFCell cell = (XSSFCell) row.createCell(col);
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFComment comment = drawing.createCellComment(
                new XSSFClientAnchor(0, 0, 0, 0, col, row.getRowNum(), col + 2, row.getRowNum() + 2));
        comment.setString(new XSSFRichTextString(text));
        cell.setCellComment(comment);
    }

    private static int addRow(XSSFSheet sheet, int rowNum, String type, Object value, String comment, Hyperlink link) {
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(type);
        if (value instanceof String) row.createCell(1).setCellValue((String) value);
        else if (value instanceof Integer) row.createCell(1).setCellValue((Integer) value);
        else if (value instanceof Double) row.createCell(1).setCellValue((Double) value);
        else if (value instanceof Boolean) row.createCell(1).setCellValue((Boolean) value);
        if (comment != null) addComment(sheet.getWorkbook(), row, 1, comment);
        if (link != null) row.createCell(3).setHyperlink(link);
        return rowNum + 1;
    }

    // ========== Внесение редких XML-конструкций ==========
    private static void addRareXmlConstructs(Path source, Path target) throws Exception {
        Map<String, byte[]> entries = new HashMap<>();
        try (ZipFile zip = ZipFile.builder().setFile(source.toFile()).get()) {
            Enumeration<ZipArchiveEntry> en = zip.getEntries();
            while (en.hasMoreElements()) {
                ZipArchiveEntry entry = en.nextElement();
                if (!entry.isDirectory()) {
                    ByteArrayOutputStream baos = new ByteArrayOutputStream();
                    try (InputStream is = zip.getInputStream(entry)) {
                        IOUtils.copy(is, baos);
                    }
                    entries.put(entry.getName(), baos.toByteArray());
                }
            }
        }
        // Модифицируем sheet1.xml (DataTypes)
        modifySheetXml(entries, "xl/worksheets/sheet1.xml");
        // Собираем обратно
        try (ZipArchiveOutputStream zos = new ZipArchiveOutputStream(Files.newOutputStream(target))) {
            for (Map.Entry<String, byte[]> e : entries.entrySet()) {
                ZipArchiveEntry entry = new ZipArchiveEntry(e.getKey());
                zos.putArchiveEntry(entry);
                zos.write(e.getValue());
                zos.closeArchiveEntry();
            }
        }
    }

    private static void modifySheetXml(Map<String, byte[]> entries, String sheetPath) {
        if (!entries.containsKey(sheetPath)) return;
        String xml = new String(entries.get(sheetPath), StandardCharsets.UTF_8);
        // Ищем ячейку с текстом "<Тег> & значение \"кавычки'" и заменяем на CDATA
        String targetText = "<Тег> & значение \"кавычки'";
        String cdataBlock = "<![CDATA[<Тег> & значение \"кавычки' и специальные символы < > &]]>";
        xml = xml.replace(targetText, cdataBlock);
        // Вставляем комментарий после декларации
        String comment = "<!-- Это тестовый комментарий в XML -->\n";
        xml = xml.replaceFirst("(<\\?xml[^?]+\\?>\\s*)", "$1" + comment);
        // Вставляем processing instruction после <worksheet ...>
        String pi = "<?myInstruction version=\"1.0\"?>\n";
        xml = xml.replaceFirst("(<worksheet[^>]*>)", "$1\n" + pi);
        entries.put(sheetPath, xml.getBytes(StandardCharsets.UTF_8));
    }
}