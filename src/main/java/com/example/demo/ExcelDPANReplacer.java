package com.example.demo;

import javax.xml.namespace.QName;
import javax.xml.stream.*;
import javax.xml.stream.events.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class ExcelDPANReplacer {

    private final XMLInputFactory xmlFactory = XMLInputFactory.newInstance();
    private final XMLOutputFactory outFactory = XMLOutputFactory.newInstance();

    private final DPANPatternMatcher patternMatcher;

    public ExcelDPANReplacer() {
        this.patternMatcher = new DPANPatternMatcher();
        xmlFactory.setProperty(XMLInputFactory.SUPPORT_DTD, false);
        xmlFactory.setProperty(XMLInputFactory.IS_NAMESPACE_AWARE, true); // <- важно
        xmlFactory.setProperty(XMLInputFactory.IS_VALIDATING, false);
        outFactory.setProperty(XMLOutputFactory.IS_REPAIRING_NAMESPACES, false);
    }

    private static class SharedStringData {
        final String raw;
        final String plain;
        final List<Integer> textLengths;

        SharedStringData(String raw, String plain, List<Integer> textLengths) {
            this.raw = raw;
            this.plain = plain;
            this.textLengths = textLengths;
        }
    }

    public void replaceAllDpans(Path inputPath,
                                Map<String, String> replacementMap,
                                Map<String, Set<Integer>> excludedColsPerSheet, Path tempDir, int bufferSize) throws Exception {
        Path sharedPath = tempDir.resolve("xl/sharedStrings.xml");
        Map<Integer, SharedStringData> sharedStringsFilteredCandidates = loadSharedStringsFilteredCandidates(sharedPath);
        Map<String, String> sheetFileToVisible = getSheetFileToVisibleName(inputPath);
        processWorksheets(tempDir, replacementMap, excludedColsPerSheet, sheetFileToVisible, sharedStringsFilteredCandidates, bufferSize);
    }

    private Map<Integer, SharedStringData> loadSharedStringsFilteredCandidates(Path shared) throws Exception {
        Map<Integer, SharedStringData> map = new HashMap<>();
        if (shared == null || !Files.exists(shared)) return map;

        try (InputStream is = Files.newInputStream(shared)) {
            XMLStreamReader reader = xmlFactory.createXMLStreamReader(is, StandardCharsets.UTF_8.name());
            StringBuilder rawBuilder = new StringBuilder();
            StringBuilder plainBuilder = new StringBuilder();
            List<Integer> textLengths = new ArrayList<>();
            boolean inSi = false, inT = false;
            int idx = 0;
            int currentTextLength = 0;

            while (reader.hasNext()) {
                int ev = reader.next();
                if (ev == XMLStreamConstants.START_ELEMENT) {
                    String local = reader.getLocalName();
                    if ("si".equals(local)) {
                        inSi = true;
                        rawBuilder.setLength(0);
                        plainBuilder.setLength(0);
                        textLengths.clear();
                    } else if (inSi) {
                        rawBuilder.append('<').append(local);
                        for (int i = 0; i < reader.getAttributeCount(); i++) {
                            rawBuilder.append(' ')
                                    .append(reader.getAttributeLocalName(i))
                                    .append("=\"")
                                    .append(escapeAttr(reader.getAttributeValue(i)))
                                    .append('"');
                        }
                        rawBuilder.append('>');
                        if ("t".equals(local)) {
                            inT = true;
                            currentTextLength = 0;
                        }
                    }
                } else if (ev == XMLStreamConstants.CHARACTERS || ev == XMLStreamConstants.CDATA) {
                    if (inSi) {
                        String txt = reader.getText();
                        rawBuilder.append(escapeXml(txt));
                        if (inT) {
                            plainBuilder.append(txt);
                            currentTextLength += txt.length();
                        }
                    }
                } else if (ev == XMLStreamConstants.END_ELEMENT) {
                    String local = reader.getLocalName();
                    if ("si".equals(local)) {
                        String plainText = plainBuilder.toString();
                        if (patternMatcher.containsDPANCandidate(plainText)) {
                            // System.out.println("plainText, "+ plainText +", index:"+ idx);
                            map.put(idx, new SharedStringData(
                                    rawBuilder.toString(),
                                    plainText,
                                    new ArrayList<>(textLengths)
                            ));
                        }
                        idx++; // ВАЖНО: увеличиваем для следующего элемента
                        inSi = false;
                        inT = false;
                    } else if (inSi) {
                        rawBuilder.append("</").append(local).append('>');
                        if ("t".equals(local)) {
                            inT = false;
                            textLengths.add(currentTextLength);
                        }
                    }
                }
            }
            reader.close();
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
                                   Map<Integer, SharedStringData> sharedStrings, int bufferSize) throws Exception {

        Path sheetsDir = tempDir.resolve("xl/worksheets");
        if (!Files.exists(sheetsDir)) return;

        try (DirectoryStream<Path> stream = Files.newDirectoryStream(sheetsDir, "*.xml")) {
            for (Path sheetFile : stream) {
                String sheetFileName = sheetFile.getFileName().toString();
                String visibleName = sheetFileToVisible.getOrDefault(sheetFileName, sheetFileName);
                Set<Integer> excluded = excludedColsPerSheet.getOrDefault(visibleName, Collections.emptySet());
                System.out.println("excluded:" + excluded);
                processSingleSheet(sheetFile, replacementMap, excluded, sharedStrings, bufferSize);
            }
        }
    }

    // ==================== ОСНОВНОЙ МЕТОД - ПОТОКОВАЯ ОБРАБОТКА ЛИСТА (БЫСТРО И БЕЗ ЖОРА ПАМЯТИ) ====================
    private void processSingleSheet(Path sheetFile, Map<String, String> replacementMap,
                                    Set<Integer> excluded, Map<Integer, SharedStringData> sharedStringsFiltered, int bufferSize) throws Exception {

        Path tmp = sheetFile.getParent().resolve(sheetFile.getFileName().toString() + ".tmp");

        XMLEventReader reader = null;
        XMLStreamWriter writer = null;

        try (InputStream is = new BufferedInputStream(Files.newInputStream(sheetFile), bufferSize);
             OutputStream os = new BufferedOutputStream(Files.newOutputStream(tmp), bufferSize)) {

            reader = xmlFactory.createXMLEventReader(is, StandardCharsets.UTF_8.name());
            writer = outFactory.createXMLStreamWriter(os, StandardCharsets.UTF_8.name());

            // <-- НОВОЕ: кэш модифицированных raw-строк по индексу shared string
            Map<Integer, String> modifiedRawCache = new HashMap<>();

            List<XMLEvent> cellEvents = new ArrayList<>();
            boolean insideCell = false;
            boolean insideShared = false;
            boolean insideV = false;
            String cellRef = null;
            StringBuilder vContent = new StringBuilder();
            Map<String, String> cellAttrs = new LinkedHashMap<>();

            while (reader.hasNext()) {
                XMLEvent event = reader.nextEvent();

                if (event.isStartElement()) {
                    StartElement start = event.asStartElement();
                    String local = start.getName().getLocalPart();

                    if ("c".equals(local)) {
                        insideCell = true;
                        cellEvents.clear();
                        cellAttrs.clear();
                        cellRef = null;
                        insideShared = false;
                        vContent.setLength(0);

                        cellEvents.add(event);
                        Iterator<?> attrs = start.getAttributes();
                        while (attrs.hasNext()) {
                            Attribute attr = (Attribute) attrs.next();
                            String name = attr.getName().getLocalPart();
                            String value = attr.getValue();
                            if ("r".equals(name)) cellRef = value;
                            if ("t".equals(name) && "s".equals(value)) insideShared = true;
                            cellAttrs.put(name, value);
                        }
                        continue;
                    } else if (insideCell) {
                        if (insideShared && "v".equals(local)) insideV = true;
                        cellEvents.add(event);
                        continue;
                    } else {
                        copyStartElement(start, writer);
                        continue;
                    }
                }

                if (event.isEndElement()) {
                    EndElement end = event.asEndElement();
                    String local = end.getName().getLocalPart();

                    if ("c".equals(local) && insideCell) {
                        cellEvents.add(event);
                        boolean replaced = false;
                        if (insideShared && cellRef != null && !vContent.toString().trim().isEmpty()) {
                            try {
                                int idx = Integer.parseInt(vContent.toString().trim());
                                if (!excluded.contains(getColumnIndex(cellRef))) {
                                    SharedStringData data = sharedStringsFiltered.get(idx);
                                    if (data != null) {
                                        // <-- НОВОЕ: используем кэш
                                        String modifiedRaw = modifiedRawCache.get(idx);
                                        if (modifiedRaw == null) {
                                            // Первый раз – вычисляем
                                            String plainText = data.plain;
                                            String replacedPlain = patternMatcher.replaceDPANsWithMap(plainText, replacementMap);
                                            if (!plainText.equals(replacedPlain)) {
                                                //модификация текста для shared строки insideShared==true
                                                //Исходный фрагмент внутри <si>: <r><rPr><b/></rPr><t>DPAN_12345</t></r> <t>(проверка)</t>
                                                //После замены: <r><rPr><b/></rPr><t>NEW_67890</t></r> <t>(проверка)</t>
                                                modifiedRaw = replaceTextInRawXml(data, replacedPlain);
                                                modifiedRawCache.put(idx, modifiedRaw);
                                            } else {
                                                // Замены нет – сохраняем пустую строку как маркер
                                                modifiedRawCache.put(idx, "");
                                            }
                                        }
                                        // Если modifiedRaw не пустая – заменяем (Меняем t="s" на t="inlineStr", остальные атрибуты сохраняем)
                                        if (modifiedRaw != null && !modifiedRaw.isEmpty()) {
                                            writeInlineCell(writer, modifiedRaw, cellAttrs);
                                            replaced = true;
                                        }
                                    }
                                }
                            } catch (NumberFormatException ignored) {
                            }
                        }
                        if (!replaced) {
                            for (XMLEvent ev : cellEvents) {
                                if (ev.isStartElement()) {
                                    copyStartElement(ev.asStartElement(), writer);
                                } else if (ev.isEndElement()) {
                                    writer.writeEndElement();
                                } else if (ev.isCharacters()) {
                                    writer.writeCharacters(ev.asCharacters().getData());
                                } else if (ev.getEventType() == XMLStreamConstants.CDATA) {
                                    writer.writeCData(ev.asCharacters().getData());
                                }
                            }
                        }
                        insideCell = false;
                        insideShared = false;
                        insideV = false;
                        cellRef = null;
                        continue;
                    } else if (insideCell) {
                        cellEvents.add(event);
                        if ("v".equals(local)) insideV = false;
                        continue;
                    } else {
                        writer.writeEndElement();
                        continue;
                    }
                }

                if (event.isCharacters()) {
                    Characters chars = event.asCharacters();
                    if (insideCell) {
                        if (insideV && insideShared) {
                            vContent.append(chars.getData());
                        }
                        cellEvents.add(event);
                        continue;
                    } else {
                        writer.writeCharacters(chars.getData());
                        continue;
                    }
                }
                if (event.getEventType() == XMLStreamConstants.CDATA) {
                    if (insideCell) {
                        if (insideV && insideShared) {
                            vContent.append(event.asCharacters().getData());
                        }
                        cellEvents.add(event);
                    } else {
                        writer.writeCData(event.asCharacters().getData());
                    }
                }
            }

            writer.flush();
        } finally {
            try {
                if (writer != null) writer.close();
            } catch (Exception ignored) {
            }
            try {
                if (reader != null) reader.close();
            } catch (Exception ignored) {
            }
        }
        Files.move(tmp, sheetFile, StandardCopyOption.REPLACE_EXISTING);
    }

    private void copyStartElement(StartElement start, XMLStreamWriter writer) throws XMLStreamException {
        QName name = start.getName();
        writer.writeStartElement(name.getPrefix(), name.getLocalPart(), name.getNamespaceURI());
        Iterator<?> attrs = start.getAttributes();
        while (attrs.hasNext()) {
            Attribute attr = (Attribute) attrs.next();
            QName attrName = attr.getName();
            writer.writeAttribute(attrName.getPrefix(), attrName.getNamespaceURI(),
                    attrName.getLocalPart(), attr.getValue());
        }
        Iterator<?> nsIter = start.getNamespaces();
        while (nsIter.hasNext()) {
            Namespace ns = (Namespace) nsIter.next();
            writer.writeNamespace(ns.getPrefix(), ns.getNamespaceURI());
        }
    }

    private void writeInlineCell(XMLStreamWriter writer, String modifiedRaw,
                                 Map<String, String> attrs) throws Exception {
        writer.writeStartElement("c");
        boolean hasT = false;
        for (Map.Entry<String, String> entry : attrs.entrySet()) {
            String name = entry.getKey();
            String value = entry.getValue();
            if ("t".equals(name)) {
                writer.writeAttribute("t", "inlineStr");
                hasT = true;
            } else {
                writer.writeAttribute(name, value);
            }
        }
        // 1. Меняем t="s" на t="inlineStr"
        if (!hasT) writer.writeAttribute("t", "inlineStr");
        writer.writeStartElement("is");
        // Сохраняем пробелы
        writer.writeAttribute("xml", "http://www.w3.org/XML/1998/namespace", "space", "preserve");
        writeXmlFragment(writer, modifiedRaw);
        writer.writeEndElement(); // </is>
        writer.writeEndElement(); // </c>
    }

    private void writeXmlFragment(XMLStreamWriter writer, String fragment) throws Exception {
        // Оптимизация: если фрагмент простой, пишем без парсинга
        if (fragment.startsWith("<t>") && fragment.endsWith("</t>") && !fragment.contains("<r>")) {
            String text = fragment.substring(3, fragment.length() - 4);
            writer.writeStartElement("t");
            writer.writeCharacters(text);
            writer.writeEndElement();
            return;
        }

        XMLStreamReader reader = null;
        try (StringReader sr = new StringReader("<root>" + fragment + "</root>")) {
            reader = xmlFactory.createXMLStreamReader(sr);
            while (reader.hasNext()) {
                int ev = reader.next();
                switch (ev) {
                    case XMLStreamConstants.START_ELEMENT:
                        String local = reader.getLocalName();
                        if (!"root".equals(local)) {
                            String prefix = reader.getPrefix();
                            String ns = reader.getNamespaceURI();
                            if (prefix != null && !prefix.isEmpty()) {
                                writer.writeStartElement(prefix, local, ns);
                            } else {
                                writer.writeStartElement(local);
                            }
                            for (int i = 0; i < reader.getAttributeCount(); i++) {
                                String attrPrefix = reader.getAttributePrefix(i);
                                String attrNs = reader.getAttributeNamespace(i);
                                String attrLocal = reader.getAttributeLocalName(i);
                                String value = reader.getAttributeValue(i);
                                if (attrPrefix != null && !attrPrefix.isEmpty()) {
                                    writer.writeAttribute(attrPrefix, attrNs, attrLocal, value);
                                } else {
                                    writer.writeAttribute(attrLocal, value);
                                }
                            }
                        }
                        break;
                    case XMLStreamConstants.END_ELEMENT:
                        local = reader.getLocalName();
                        if (!"root".equals(local)) {
                            writer.writeEndElement();
                        }
                        break;
                    case XMLStreamConstants.CHARACTERS:
                        writer.writeCharacters(reader.getText());
                        break;
                    case XMLStreamConstants.CDATA:
                        writer.writeCData(reader.getText());
                        break;
                }
            }
        } finally {
            if (reader != null) {
                try {
                    reader.close();
                } catch (XMLStreamException ignored) {
                }
            }
        }
    }

    private String replaceTextInRawXml(SharedStringData data, String newText) {
        List<Integer> textLengths = data.textLengths;
        // Если нет ни одного <t>, то нечего менять – возвращаем оригинал
        if (textLengths.isEmpty()) return data.raw;

        // 1. Разбиваем новый текст на части по длинам исходных <t>
        String[] parts = new String[textLengths.size()];
        int pos = 0;
        int newLen = newText.length();
        for (int i = 0; i < parts.length; i++) {
            int len = textLengths.get(i);// сколько символов было в i-м <t> в оригинале
            if (pos < newLen) {
                // Берём от newText не более len символов, но не выходя за границу
                int take = Math.min(len, newLen - pos);
                parts[i] = newText.substring(pos, pos + take);
                pos = pos + take;
            } else {
                // Если новый текст уже закончился, остальные части остаются пустыми
                parts[i] = "";
            }
        }
        // Если новый текст оказался длиннее суммы всех оригинальных длин,
        // то оставшуюся часть добавляем к последнему куску
        if (pos < newLen) {
            parts[parts.length - 1] = parts[parts.length - 1] + newText.substring(pos);
        }

        // 2. Заменяем текст внутри каждого <t> в исходном raw
        StringBuilder sb = new StringBuilder(data.raw);
        int i = 0;
        int partIdx = 0;
        while (i < sb.length() && partIdx < parts.length) {
            // Ищем открывающий тег <t (возможно, с атрибутами)
            int openStart = sb.indexOf("<t", i);
            if (openStart == -1) break; // нет больше <t
            int openEnd = sb.indexOf(">", openStart);// позиция закрывающей >
            if (openEnd == -1) break;
            int closeStart = sb.indexOf("</t>", openEnd); // позиция начала закрывающего </t>
            if (closeStart == -1) break;
            // Позиция, где начинается текст (сразу после >)
            int textStart = openEnd + 1;
            // Заменяем текст между > и </t> на экранированную версию части нового текста
            sb.replace(textStart, closeStart, escapeXml(parts[partIdx]));
            // Перемещаем указатель за закрывающий </t> и переходим к следующей части
            i = closeStart + 4;// длина "</t>" = 4
            partIdx++;
        }
        // Возвращаем изменённый XML-фрагмент
        return sb.toString();
    }

    private static int getColumnIndex(String cellRef) {
        String colLetters = cellRef.replaceAll("[0-9$]", "").toUpperCase();
        int idx = 0;
        for (char c : colLetters.toCharArray()) {
            idx = idx * 26 + (c - 'A' + 1);
        }
        return idx - 1;
    }

    private static String escapeXml(String s) {
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
    }

    private static String escapeAttr(String s) {
        return s.replace("&", "&amp;").replace("\"", "&quot;").replace("<", "&lt;").replace(">", "&gt;");
    }
}