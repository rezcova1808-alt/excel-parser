package com.example.demo;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamReader;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class ExcelDPANReplacer {

    private final DPANPatternMatcher patternMatcher;

    public ExcelDPANReplacer() {
        this.patternMatcher = new DPANPatternMatcher();
    }

    public static final int BUFFER_SIZE = 256 * 1024;

    public void replaceAllDpans(Path inputPath, Path outputPath,
                                Map<String, String> replacementMap,
                                Map<String, Set<Integer>> excludedColsPerSheet) throws Exception {

        Path tempDir = Files.createTempDirectory("xlsx_inline_");
        try {
            FileUtils.unzip(inputPath, tempDir, BUFFER_SIZE);

            Map<Integer, SharedStringData> sharedStrings = loadSharedStrings(tempDir);
            System.out.println("Loaded " + sharedStrings.size() + " shared strings.");

            Map<String, String> sheetFileToVisible = getSheetFileToVisibleName(inputPath);

            processWorksheets(tempDir, replacementMap, excludedColsPerSheet, sheetFileToVisible, sharedStrings);

            FileUtils.zip(tempDir, outputPath, BUFFER_SIZE);
        } finally {
            FileUtils.deleteDirectory(tempDir.toFile().toPath());
        }
    }

    // ==================== ЗАГРУЗКА SHARED STRINGS ====================
    private Map<Integer, SharedStringData> loadSharedStrings(Path tempDir) throws Exception {
        Map<Integer, SharedStringData> map = new HashMap<>();
        Path shared = tempDir.resolve("xl/sharedStrings.xml");
        if (!Files.exists(shared)) return map;

        XMLInputFactory factory = XMLInputFactory.newInstance();
        factory.setProperty(XMLInputFactory.SUPPORT_DTD, false);
        try (InputStream is = Files.newInputStream(shared)) {
            XMLStreamReader reader = factory.createXMLStreamReader(is);
            int index = 0;
            StringBuilder rawBuilder = new StringBuilder();
            StringBuilder plainBuilder = new StringBuilder();
            boolean inSi = false, inT = false;

            while (reader.hasNext()) {
                int event = reader.next();
                switch (event) {
                    case XMLStreamReader.START_ELEMENT:
                        if ("si".equals(reader.getLocalName())) {
                            inSi = true;
                            rawBuilder.setLength(0);
                            plainBuilder.setLength(0);
                        } else if (inSi) {
                            rawBuilder.append('<').append(reader.getLocalName());
                            for (int i = 0; i < reader.getAttributeCount(); i++) {
                                rawBuilder.append(' ')
                                        .append(reader.getAttributeLocalName(i))
                                        .append("=\"")
                                        .append(escapeAttr(reader.getAttributeValue(i)))
                                        .append('"');
                            }
                            rawBuilder.append('>');
                            if ("t".equals(reader.getLocalName())) inT = true;
                        }
                        break;

                    case XMLStreamReader.END_ELEMENT:
                        if ("si".equals(reader.getLocalName())) {
                            map.put(index++, new SharedStringData(rawBuilder.toString(), plainBuilder.toString()));
                            inSi = false;
                        } else if (inSi) {
                            rawBuilder.append("</").append(reader.getLocalName()).append('>');
                            if ("t".equals(reader.getLocalName())) inT = false;
                        }
                        break;

                    case XMLStreamReader.CHARACTERS:
                    case XMLStreamReader.CDATA:
                        if (inSi) {
                            String text = reader.getText();
                            rawBuilder.append(escapeXml(text));
                            if (inT) plainBuilder.append(text);
                        }
                        break;
                }
            }
        }
        return map;
    }

    // ==================== ПОЛУЧЕНИЕ ИМЁН ЛИСТОВ ====================
    public Map<String, String> getSheetFileToVisibleName(Path srcPath) throws Exception {
        Map<String, String> result = new HashMap<>();
        try (ZipFile zipFile = new ZipFile(srcPath.toFile())) {
            Map<String, String> relIdToName = new HashMap<>();
            ZipEntry workbookEntry = zipFile.getEntry("xl/workbook.xml");
            if (workbookEntry != null) {
                try (InputStream is = zipFile.getInputStream(workbookEntry)) {
                    XMLStreamReader reader = XMLInputFactory.newInstance().createXMLStreamReader(is);
                    while (reader.hasNext()) {
                        int event = reader.next();
                        if (event == XMLStreamReader.START_ELEMENT && "sheet".equals(reader.getLocalName())) {
                            String name = reader.getAttributeValue(null, "name");
                            String ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                            String id = reader.getAttributeValue(ns, "id");
                            if (name != null && id != null) relIdToName.put(id, name);
                        }
                    }
                    reader.close();
                }
            }

            ZipEntry relsEntry = zipFile.getEntry("xl/_rels/workbook.xml.rels");
            if (relsEntry != null) {
                try (InputStream is = zipFile.getInputStream(relsEntry)) {
                    XMLStreamReader reader = XMLInputFactory.newInstance().createXMLStreamReader(is);
                    while (reader.hasNext()) {
                        int event = reader.next();
                        if (event == XMLStreamReader.START_ELEMENT && "Relationship".equals(reader.getLocalName())) {
                            String id = reader.getAttributeValue(null, "Id");
                            String target = reader.getAttributeValue(null, "Target");
                            if (id != null && target != null && target.startsWith("worksheets/")) {
                                String sheetFileName = target.substring(target.lastIndexOf('/') + 1);
                                String sheetName = relIdToName.get(id);
                                if (sheetName != null) result.put(sheetFileName, sheetName);
                            }
                        }
                    }
                    reader.close();
                }
            }
        }
        return result;
    }

    // ==================== ОБРАБОТКА ВСЕХ ЛИСТОВ ====================
    private void processWorksheets(Path tempDir, Map<String, String> replacementMap,
                                   Map<String, Set<Integer>> excludedColsPerSheet,
                                   Map<String, String> sheetFileToVisible,
                                   Map<Integer, SharedStringData> sharedStrings) throws Exception {
        Path sheetsDir = tempDir.resolve("xl/worksheets");
        if (!Files.exists(sheetsDir)) return;

        try (DirectoryStream<Path> stream = Files.newDirectoryStream(sheetsDir, "*.xml")) {
            for (Path sheetFile : stream) {
                String sheetFileName = sheetFile.getFileName().toString();
                String visibleName = sheetFileToVisible.getOrDefault(sheetFileName, sheetFileName);
                Set<Integer> excluded = excludedColsPerSheet.getOrDefault(visibleName, Collections.emptySet());
                System.out.println("Processing " + sheetFileName + " (" + visibleName + ") excluded cols: " + excluded);
                processSingleSheet(sheetFile, replacementMap, excluded, sharedStrings);
            }
        }
    }

    // ==================== ОСНОВНОЙ МЕТОД - ПОТОКОВАЯ ОБРАБОТКА ЛИСТА (БЫСТРО И БЕЗ ЖОРА ПАМЯТИ) ====================
    // ==================== ОБРАБОТКА ОДНОГО ЛИСТА (ПОТОКОВО, БЕЗ ЖОРА ПАМЯТИ) ====================
    private void processSingleSheet(Path sheetFile, Map<String, String> replacementMap,
                                    Set<Integer> excluded, Map<Integer, SharedStringData> sharedStrings) throws Exception {
        Path tempSheet = sheetFile.getParent().resolve(sheetFile.getFileName() + ".tmp");

        try (InputStream in = Files.newInputStream(sheetFile);
             OutputStream out = Files.newOutputStream(tempSheet)) {

            XMLInputFactory factory = XMLInputFactory.newInstance();
            factory.setProperty(XMLInputFactory.SUPPORT_DTD, false);
            XMLStreamReader reader = factory.createXMLStreamReader(in);

            OutputStreamWriter writer = new OutputStreamWriter(out, StandardCharsets.UTF_8);

            StringBuilder cellBuffer = new StringBuilder();
            boolean inCell = false;

            while (reader.hasNext()) {
                int event = reader.next();

                switch (event) {
                    case XMLStreamReader.START_ELEMENT:
                        String localName = reader.getLocalName();

                        if ("c".equals(localName)) {
                            // Начинаем собирать ячейку
                            inCell = true;
                            cellBuffer.setLength(0);
                            cellBuffer.append('<').append(localName);
                            for (int i = 0; i < reader.getAttributeCount(); i++) {
                                cellBuffer.append(' ')
                                        .append(reader.getAttributeLocalName(i))
                                        .append("=\"")
                                        .append(escapeAttr(reader.getAttributeValue(i)))
                                        .append('"');
                            }
                            cellBuffer.append('>');
                        } else if (inCell) {
                            // Внутри ячейки – копируем теги в буфер
                            cellBuffer.append('<').append(localName);
                            for (int i = 0; i < reader.getAttributeCount(); i++) {
                                cellBuffer.append(' ')
                                        .append(reader.getAttributeLocalName(i))
                                        .append("=\"")
                                        .append(escapeAttr(reader.getAttributeValue(i)))
                                        .append('"');
                            }
                            cellBuffer.append('>');
                        } else {
                            // Вне ячейки – пишем сразу
                            writer.write('<');
                            writer.write(localName);
                            for (int i = 0; i < reader.getAttributeCount(); i++) {
                                writer.write(' ');
                                writer.write(reader.getAttributeLocalName(i));
                                writer.write("=\"");
                                writer.write(escapeAttr(reader.getAttributeValue(i)));
                                writer.write('"');
                            }
                            writer.write('>');
                        }
                        break;

                    case XMLStreamReader.CHARACTERS:
                    case XMLStreamReader.CDATA:
                        String text = reader.getText();
                        if (inCell) {
                            cellBuffer.append(escapeXml(text));
                        } else {
                            writer.write(escapeXml(text));
                        }
                        break;

                    case XMLStreamReader.END_ELEMENT:
                        localName = reader.getLocalName();

                        if ("c".equals(localName)) {
                            cellBuffer.append("</").append(localName).append('>');
                            // Обрабатываем накопленную ячейку
                            String processed = processCellBuffer(cellBuffer.toString(), replacementMap, excluded, sharedStrings);
                            writer.write(processed);
                            inCell = false;
                            cellBuffer.setLength(0);
                        } else if (inCell) {
                            cellBuffer.append("</").append(localName).append('>');
                        } else {
                            writer.write("</");
                            writer.write(localName);
                            writer.write('>');
                        }
                        break;
                }
            }

            reader.close();
            writer.flush();
            writer.close();
        }

        Files.move(tempSheet, sheetFile, StandardCopyOption.REPLACE_EXISTING);
    }


    // ==================== ОБРАБОТКА БУФЕРА ЯЧЕЙКИ ====================
    private String processCellBuffer(String cellXml, Map<String, String> replacementMap,
                                     Set<Integer> excluded, Map<Integer, SharedStringData> sharedStrings) {
        String r = extractAttributeSimple(cellXml, "r");
        int col = getColumnNumber(r);
        String t = extractAttributeSimple(cellXml, "t");

        // --- Shared ---
        if ("s".equals(t)) {
            if (excluded.contains(col)) return cellXml;

            int vStart = cellXml.indexOf("<v>");
            if (vStart == -1) return cellXml;
            int vEnd = cellXml.indexOf("</v>", vStart);
            if (vEnd == -1) return cellXml;

            String indexStr = cellXml.substring(vStart + 3, vEnd);
            int idx;
            try {
                idx = Integer.parseInt(indexStr.trim());
            } catch (NumberFormatException e) {
                return cellXml;
            }

            SharedStringData data = sharedStrings.get(idx);
            if (data == null) return cellXml;

            String oldText = data.plain;
            if (!replacementMap.containsKey(oldText)) return cellXml;
            String newText = replacementMap.get(oldText);

            // 1. Меняем t="s" на t="inlineStr"
            String newCell = cellXml.replace(" t=\"s\"", " t=\"inlineStr\"")
                    .replace("t=\"s\"", "t=\"inlineStr\"");

            // 2. Заменяем <v>...</v> на <is><t>...</t></is>
            String beforeV = cellXml.substring(0, vStart);
            String afterV = cellXml.substring(vEnd + 4);
            newCell = beforeV + "<is><t>" + escapeXml(newText) + "</t></is>" + afterV;

            System.out.println("✅ Shared " + r + ": " + oldText + " -> " + newText);
            return newCell;
        }

        // --- Inline ---
        if ("inlineStr".equals(t)) {
            if (excluded.contains(col)) return cellXml;

            String insideIs = extractContentBetween(cellXml, "<is>", "</is>");
            if (insideIs == null) return cellXml;

            String oldText = extractContentBetween(insideIs, "<t>", "</t>");
            if (oldText == null) return cellXml;

            if (!replacementMap.containsKey(oldText)) return cellXml;
            String newText = replacementMap.get(oldText);

            int tStart = insideIs.indexOf("<t>");
            int tEnd = insideIs.indexOf("</t>", tStart);
            String newInside = insideIs.substring(0, tStart + 3) + escapeXml(newText) + insideIs.substring(tEnd);

            int isStart = cellXml.indexOf("<is>");
            int isEnd = cellXml.indexOf("</is>", isStart);
            String newCell = cellXml.substring(0, isStart + 4) + newInside + "</is>" + cellXml.substring(isEnd + 5);

            System.out.println("🔄 Inline " + r + ": " + oldText + " -> " + newText);
            return newCell;
        }

        return cellXml;
    }


    // ==================== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ====================
    private static class SharedStringData {
        final String raw;
        final String plain;
        SharedStringData(String raw, String plain) {
            this.raw = raw;
            this.plain = plain;
        }
    }

    private String replaceExactInSiFragment(String rawSi, String oldText, String newText) {
        if (rawSi == null || rawSi.isEmpty()) return rawSi;
        StringBuilder sb = new StringBuilder(rawSi);
        int pos = 0;
        while (true) {
            int tStart = sb.indexOf("<t", pos);
            if (tStart == -1) break;
            int tagEnd = sb.indexOf(">", tStart);
            if (tagEnd == -1) break;
            int tEnd = sb.indexOf("</t>", tagEnd);
            if (tEnd == -1) break;
            String content = sb.substring(tagEnd + 1, tEnd);
            if (content.equals(oldText)) {
                sb.replace(tagEnd + 1, tEnd, escapeXml(newText));
                break;
            }
            pos = tEnd + 4;
        }
        return sb.toString();
    }

    private String extractAttributeSimple(String tag, String attrName) {
        String pattern = attrName + "=\"";
        int start = tag.indexOf(pattern);
        if (start != -1) {
            start += pattern.length();
            int end = tag.indexOf('"', start);
            if (end != -1) return tag.substring(start, end);
        }
        pattern = attrName + "='";
        start = tag.indexOf(pattern);
        if (start != -1) {
            start += pattern.length();
            int end = tag.indexOf('\'', start);
            if (end != -1) return tag.substring(start, end);
        }
        return "";
    }

    private String extractContentBetween(String xml, String openTag, String closeTag) {
        int start = xml.indexOf(openTag);
        if (start == -1) return null;
        int end = xml.indexOf(closeTag, start);
        if (end == -1) return null;
        return xml.substring(start + openTag.length(), end);
    }

    private int getColumnNumber(String cellRef) {
        if (cellRef == null || cellRef.isEmpty()) return 0;
        int col = 0;
        for (char ch : cellRef.toCharArray()) {
            if (Character.isDigit(ch)) break;
            if (ch >= 'A' && ch <= 'Z') {
                col = col * 26 + (ch - 'A' + 1);
            }
        }
        return col;
    }

    private String escapeAttr(String s) {
        if (s == null) return "";
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }

    private String escapeXml(String s) {
        if (s == null) return "";
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;");
    }
}