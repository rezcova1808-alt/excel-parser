package com.example.demo;

public class TestExcelParser {

    public static void main(String[] args) {
        // Твой XML в виде ОДНОЙ СТРОКИ (я сжал его, но это не важно)
        String xml = "<worksheet Ignorable=\"x14ac xr xr2 xr3\" uid=\"{00000000-0001-0000-0000-000000000000}\"><dimension ref=\"A1:J18\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"B18\" sqref=\"B18\"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/><cols><col min=\"1\" max=\"1\" width=\"39.28515625\" customWidth=\"1\"/><col min=\"2\" max=\"2\" width=\"40.85546875\" customWidth=\"1\"/><col min=\"3\" max=\"3\" width=\"20.140625\" customWidth=\"1\"/><col min=\"4\" max=\"4\" width=\"28.140625\" customWidth=\"1\"/><col min=\"6\" max=\"6\" width=\"35.140625\" customWidth=\"1\"/><col min=\"7\" max=\"7\" width=\"25\" customWidth=\"1\"/><col min=\"8\" max=\"8\" width=\"22.7109375\" customWidth=\"1\"/><col min=\"9\" max=\"9\" width=\"27.42578125\" customWidth=\"1\"/><col min=\"10\" max=\"10\" width=\"24.42578125\" customWidth=\"1\"/></cols><sheetData><row r=\"1\" spans=\"1:10\"><c r=\"A1\" s=\"3\" t=\"s\"><v>0</v></c><c r=\"B1\" s=\"4\" t=\"s\"><v>1</v></c><c r=\"C1\" s=\"5\" t=\"s\"><v>2</v></c><c r=\"D1\" s=\"6\" t=\"s\"><v>3</v></c><c r=\"E1\" s=\"7\" t=\"s\"><v>4</v></c><c r=\"F1\" s=\"15\" t=\"s\"><v>36</v></c><c r=\"G1\" s=\"16\"/><c r=\"H1\" s=\"19\" t=\"s\"><v>36</v></c><c r=\"I1\" s=\"20\" t=\"s\"><v>36</v></c><c r=\"J1\" s=\"20\"/></row><row r=\"2\" spans=\"1:10\"><c r=\"A2\" t=\"s\"><v>5</v></c><c r=\"B2\" t=\"s\"><v>6</v></c><c r=\"F2\" s=\"17\" t=\"s\"><v>35</v></c><c r=\"G2\" s=\"14\"/><c r=\"H2\" s=\"19\"/><c r=\"I2\" s=\"17\" t=\"s\"><v>35</v></c><c r=\"J2\" s=\"17\" t=\"s\"><v>35</v></c></row></sheetData></worksheet>";

        // Запускаем обработку
        processTest(xml);
    }

    private static void processTest(String content) {
        StringBuilder result = new StringBuilder();
        int pos = 0;
        int totalCells = 0;
        int totalSkipped = 0;

        while (pos < content.length()) {
            // Ищем <c
            int cStart = content.indexOf("<c", pos);
            if (cStart == -1) {
                result.append(content, pos, content.length());
                break;
            }

            // Копируем текст до <c
            result.append(content, pos, cStart);

            // Проверяем, что это ячейка (после c должен быть пробел или >)
            boolean isCell = false;
            if (cStart + 2 < content.length()) {
                char next = content.charAt(cStart + 2);
                if (next == ' ' || next == '>') {
                    isCell = true;
                }
            }

            if (isCell) {
                // Это ячейка – ищем закрывающий </c>
                int cEnd = content.indexOf("</c>", cStart);
                if (cEnd == -1) {
                    result.append(content, cStart, content.length());
                    break;
                }
                String cellXml = content.substring(cStart, cEnd + 4);
                totalCells++;
                System.out.println("Найдена ячейка #" + totalCells + ":");
                System.out.println(cellXml);
                System.out.println("-----");

                // В реальном коде здесь была бы замена через processCellXml
                // Мы просто копируем ячейку как есть
                result.append(cellXml);
                pos = cEnd + 4;
            } else {
                // Это не ячейка (например <col>) – копируем тег целиком
                int tagEnd = content.indexOf('>', cStart);
                if (tagEnd == -1) {
                    result.append(content, cStart, content.length());
                    break;
                }
                String tag = content.substring(cStart, tagEnd + 1);
                totalSkipped++;
                System.out.println("Пропускаем тег (не ячейка): " + tag);
                result.append(tag);
                pos = tagEnd + 1;
            }
        }

        System.out.println("\nВсего найдено ячеек: " + totalCells);
        System.out.println("Всего пропущено других тегов: " + totalSkipped);
        System.out.println("\nИтоговый XML (без изменений):");
        System.out.println(result);
    }
}