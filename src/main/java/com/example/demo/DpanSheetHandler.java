package com.example.demo;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.SharedStrings;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class DpanSheetHandler extends DefaultHandler {
    private final Set<Integer> excludedCols;
    private final DPANPatternMatcher patternMatcher;
    private final SharedStrings sharedStrings;
    private final Set<String> values = new HashSet<>();
    //Сбор кандидатов по адресу ячейки и его кандидаты
    // может эта коллекция и не потребуется!!!! много памяти ест
    //private final Map<String, Set<String>> candidateToAddresses = new HashMap<>();

    private String currentCellRef = "";
    private String cellType = "";
    private final StringBuilder rawValue = new StringBuilder();
    private boolean inCell = false;
    private boolean inValue = false;
    private boolean inInlineString = false;

    public DpanSheetHandler(SharedStrings sharedStrings, Set<Integer> excludedCols, DPANPatternMatcher patternMatcher) {
        this.sharedStrings = sharedStrings;
        this.excludedCols = excludedCols;
        this.patternMatcher = patternMatcher;
    }

    public Set<String> getValues() {
        return values;
    }

    /*public Map<String, Set<String>> getCandidateToAddresses() {
        return candidateToAddresses;
    }*/

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) {
        String name = localName.isEmpty() ? qName : localName;
        switch (name) {
            case "c":
                inCell = true;
                currentCellRef = attributes.getValue("r");
                cellType = attributes.getValue("t") == null ? "" : attributes.getValue("t");
                rawValue.setLength(0);
                break;
            case "v":
                if (inCell) {
                    inValue = true;
                }
                break;
            case "is":
                inInlineString = true;
                break;
            case "t":
                if (inInlineString || inCell) {
                    inValue = true;
                }
                break;
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) {
        if (inValue) rawValue.append(ch, start, length);
    }

    @Override
    public void endElement(String uri, String localName, String qName) {
        String name = localName.isEmpty() ? qName : localName;
        switch (name) {
            case "c":
                if (inCell) {
                    String raw = rawValue.toString().trim();
                    String value = convertRawToValue(raw, cellType);
                    //System.out.println("ref:" + currentCellRef + "," + value);
                    if (!value.isEmpty()) {
                        int col = new CellReference(currentCellRef).getCol();
                        if (!excludedCols.contains(col)) {
                            // Преобразуем raw в финальное строковое значение
                            List<String> candidates = patternMatcher.findDPANCandidates(value);
                            for (String candidate : candidates) {
                                if (patternMatcher.isExactDPAN(candidate)) {
                                    values.add(candidate);
                                    //candidateToAddresses.computeIfAbsent(candidate, k -> new HashSet<>()).add(cellRef);// может не нужно
                                }
                            }

                        }
                    }
                }
                inCell = false;
                inValue = false;
                inInlineString = false;
                break;
            case "v":
            case "t":
                inValue = false;
                break;
            case "is":
                inInlineString = false;
                break;
        }
    }

    // Преобразование сырого значения в строку с учётом типа ячейки
    private String convertRawToValue(String raw, String cellType) {
        if (raw.isEmpty()) {
            return "";
        }
        switch (cellType) {
            case "s": // shared string
                int idx = Integer.parseInt(raw);
                String str = sharedStrings.getItemAt(idx).getString();
                return str == null ? "" : str.trim();
            //"inlineStr" – строка, хранящаяся прямо внутри ячейки в теге <is><t>...</t></is>.
            //"str" – формульная строка
            case "inlineStr":
            case "str":
                return raw;
            case "b": // boolean
                return "1".equals(raw) ? "true" : "false";
            default:
                return raw;
        }
    }
}