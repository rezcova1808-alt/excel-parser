package com.example.demo;

import java.util.*;

public class TestExcelParser2 {

    public static void main(String[] args) {
        // Твой XML в виде одной строки (я сжал его для компактности)
        String xml = "<worksheet Ignorable=\"x14ac xr xr2 xr3\" uid=\"{00000000-0001-0000-0000-000000000000}\"><dimension ref=\"A1:J18\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"B18\" sqref=\"B18\"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/><cols><col min=\"1\" max=\"1\" width=\"39.28515625\" customWidth=\"1\"/><col min=\"2\" max=\"2\" width=\"40.85546875\" customWidth=\"1\"/><col min=\"3\" max=\"3\" width=\"20.140625\" customWidth=\"1\"/><col min=\"4\" max=\"4\" width=\"28.140625\" customWidth=\"1\"/><col min=\"6\" max=\"6\" width=\"35.140625\" customWidth=\"1\"/><col min=\"7\" max=\"7\" width=\"25\" customWidth=\"1\"/><col min=\"8\" max=\"8\" width=\"22.7109375\" customWidth=\"1\"/><col min=\"9\" max=\"9\" width=\"27.42578125\" customWidth=\"1\"/><col min=\"10\" max=\"10\" width=\"24.42578125\" customWidth=\"1\"/></cols><sheetData><row r=\"1\" spans=\"1:10\"><c r=\"A1\" s=\"3\" t=\"s\"><v>0</v></c><c r=\"B1\" s=\"4\" t=\"s\"><v>1</v></c><c r=\"C1\" s=\"5\" t=\"s\"><v>2</v></c><c r=\"D1\" s=\"6\" t=\"s\"><v>3</v></c><c r=\"E1\" s=\"7\" t=\"s\"><v>4</v></c><c r=\"F1\" s=\"15\" t=\"s\"><v>36</v></c><c r=\"G1\" s=\"16\"/><c r=\"H1\" s=\"19\" t=\"s\"><v>36</v></c><c r=\"I1\" s=\"20\" t=\"s\"><v>36</v></c><c r=\"J1\" s=\"20\"/></row><row r=\"2\" spans=\"1:10\"><c r=\"A2\" t=\"s\"><v>5</v></c><c r=\"B2\" t=\"s\"><v>6</v></c><c r=\"F2\" s=\"17\" t=\"s\"><v>35</v></c><c r=\"G2\" s=\"14\"/><c r=\"H2\" s=\"19\"/><c r=\"I2\" s=\"17\" t=\"s\"><v>35</v></c><c r=\"J2\" s=\"17\" t=\"s\"><v>35</v></c></row></sheetData></worksheet>";

        // --- Конфигурация замены ---
        // Карта: что меняем -> на что меняем
        Map<String, String> replacementMap = new HashMap<>();
        replacementMap.put("0", "Zero");
        replacementMap.put("1", "One");
        replacementMap.put("36", "ThirtySix");

        // Исключённые колонки (номера 1-базированные). Например, колонка A (1) и C (3) не меняем.
        Set<Integer> excludedCols = new HashSet<>(Arrays.asList(1, 3));

        // Запускаем обработку с заменой
        String resultXml = processWithReplacement(xml, replacementMap, excludedCols);

        // Вывод результата
        System.out.println("\n=== Итоговый XML после замены ===\n" + resultXml);
    }

    /**
     * Основной метод: разбирает XML, находит все ячейки <c>, заменяет текст
     * в inlineStr и plain, НО НЕ ТРОГАЕТ SHARED (t="s"), с учётом исключённых колонок.
     */
    private static String processWithReplacement(String content,
                                                 Map<String, String> replacementMap,
                                                 Set<Integer> excludedCols) {
        StringBuilder result = new StringBuilder(content.length() + 1024);
        int pos = 0;
        int totalCells = 0;
        int modifiedCells = 0;

        while (pos < content.length()) {
            int cStart = content.indexOf("<c", pos);
            if (cStart == -1) {
                result.append(content, pos, content.length());
                break;
            }

            // Копируем текст до <c
            result.append(content, pos, cStart);

            // Проверяем, что это именно ячейка (после "c" пробел или ">")
            boolean isCell = false;
            if (cStart + 2 < content.length()) {
                char next = content.charAt(cStart + 2);
                if (next == ' ' || next == '>') {
                    isCell = true;
                }
            }

            if (!isCell) {
                // Это не ячейка (например <col> или <c в другом контексте) – копируем как есть
                int tagEnd = content.indexOf('>', cStart);
                if (tagEnd == -1) {
                    result.append(content, cStart, content.length());
                    break;
                }
                result.append(content, cStart, tagEnd + 1);
                pos = tagEnd + 1;
                continue;
            }

            // Это ячейка – ищем закрывающий </c>
            int cEnd = content.indexOf("</c>", cStart);
            if (cEnd == -1) {
                result.append(content, cStart, content.length());
                break;
            }

            String cellTag = content.substring(cStart, cEnd + 4); // +4 за "</c>"
            totalCells++;

            // Обрабатываем ячейку (заменяем текст, если нужно)
            String processedCell = processCell(cellTag, replacementMap, excludedCols);
            if (!processedCell.equals(cellTag)) {
                modifiedCells++;
            }

            result.append(processedCell);
            pos = cEnd + 4;
        }

        System.out.println("Всего ячеек: " + totalCells + ", изменено: " + modifiedCells);
        return result.toString();
    }

    /**
     * Обрабатывает одну ячейку: определяет тип и при необходимости заменяет текст.
     * Shared (t="s") – пропускаем (не меняем).
     * InlineStr (t="inlineStr") – меняем текст внутри <is><t>...</t></is>.
     * Plain (без t или t="str") – меняем текст внутри <v>...</v> (но только если это не число).
     */
    private static String processCell(String cellTag,
                                      Map<String, String> replacementMap,
                                      Set<Integer> excludedCols) {
        // Извлекаем номер колонки
        String ref = extractAttribute(cellTag, "r");
        int col = extractColumnFromCellRef(ref);
        if (col > 0 && excludedCols.contains(col)) {
            // Колонка исключена – возвращаем без изменений
            return cellTag;
        }

        // Определяем тип
        String type = extractAttribute(cellTag, "t");

        // Shared – не меняем
        if ("s".equals(type)) {
            return cellTag;
        }

        // InlineStr – меняем текст внутри <is><t>
        if ("inlineStr".equals(type)) {
            return replaceInlineText(cellTag, replacementMap);
        }

        // Plain (без t или t="str") – меняем текст внутри <v>
        // Но если это число – не меняем (опционально)
        return replacePlainText(cellTag, replacementMap);
    }

    /**
     * Заменяет текст внутри <is><t>...</t></is>.
     */
    private static String replaceInlineText(String cellTag, Map<String, String> replacementMap) {
        int isStart = cellTag.indexOf("<is>");
        if (isStart == -1) return cellTag;
        int isEnd = cellTag.indexOf("</is>", isStart);
        if (isEnd == -1) return cellTag;

        String beforeIs = cellTag.substring(0, isStart);
        String insideIs = cellTag.substring(isStart + 4, isEnd);
        String afterIs = cellTag.substring(isEnd + 5);

        int tStart = insideIs.indexOf("<t>");
        if (tStart == -1) return cellTag;
        int tEnd = insideIs.indexOf("</t>", tStart);
        if (tEnd == -1) return cellTag;

        String beforeT = insideIs.substring(0, tStart + 3);
        String text = insideIs.substring(tStart + 3, tEnd);
        String afterT = insideIs.substring(tEnd); // включает </t>

        if (replacementMap.containsKey(text)) {
            String newText = replacementMap.get(text);
            newText = escapeXml(newText);
            return beforeIs + beforeT + newText + afterT + afterIs;
        }
        return cellTag;
    }

    /**
     * Заменяет текст внутри <v>...</v> для plain-ячеек.
     * Не заменяет, если значение – число (можно убрать это условие).
     */
    private static String replacePlainText(String cellTag, Map<String, String> replacementMap) {
        int vStart = cellTag.indexOf("<v>");
        if (vStart == -1) return cellTag;
        int vEnd = cellTag.indexOf("</v>", vStart);
        if (vEnd == -1) return cellTag;

        String beforeV = cellTag.substring(0, vStart + 3);
        String value = cellTag.substring(vStart + 3, vEnd);
        String afterV = cellTag.substring(vEnd);

        // Если это число – не заменяем (чтобы не испортить числа)
        if (value.matches("-?\\d+(\\.\\d+)?")) {
            return cellTag;
        }

        if (replacementMap.containsKey(value)) {
            String newValue = replacementMap.get(value);
            newValue = escapeXml(newValue);
            return beforeV + newValue + afterV;
        }
        return cellTag;
    }

    // ------------------- Вспомогательные методы -------------------

    /**
     * Извлекает значение атрибута по имени (регистр учитывается).
     */
    private static String extractAttribute(String tag, String attrName) {
        String pattern = attrName + "=\"";
        int start = tag.indexOf(pattern);
        if (start == -1) {
            pattern = attrName + "='";
            start = tag.indexOf(pattern);
        }
        if (start == -1) return null;
        start += pattern.length();
        int end = tag.indexOf('"', start);
        if (end == -1) end = tag.indexOf("'", start);
        if (end == -1) return null;
        return tag.substring(start, end);
    }

    /**
     * Из ссылки вида "A1" или "AB123" извлекает номер колонки (1‑based).
     */
    private static int extractColumnFromCellRef(String cellRef) {
        if (cellRef == null || cellRef.isEmpty()) return 0;
        int col = 0;
        for (int i = 0; i < cellRef.length(); i++) {
            char ch = cellRef.charAt(i);
            if (Character.isDigit(ch)) break;
            if (ch >= 'A' && ch <= 'Z') {
                col = col * 26 + (ch - 'A' + 1);
            }
        }
        return col;
    }

    /**
     * Экранирование XML-спецсимволов.
     */
    private static String escapeXml(String s) {
        if (s == null) return "";
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }
}