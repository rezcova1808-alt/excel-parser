package com.example.demo;


import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.SharedStrings;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.IntStream;

public class MarkersHandler extends DefaultHandler {
    private final SharedStrings sharedStrings;
    private final List<CellRangeAddress> allRegions;
    private final Set<String> markers;
    private final int rowStart;
    private final int rowEnd;
    private final Set<Integer> excluded = new HashSet<>();

    private final StringBuilder valueBuffer = new StringBuilder();
    private String currentCellRef;
    private String cellType;
    private int currentRow0 = -1;
    private boolean inCell = false;
    private boolean inValue = false;
    private boolean inInlineString = false;

    public MarkersHandler(SharedStrings sharedStrings,
                          List<CellRangeAddress> allRegions,
                          Set<String> markers,
                          int rowStart,
                          int rowEnd) {
        this.sharedStrings = sharedStrings;
        this.allRegions = allRegions;
        this.markers = markers;
        this.rowStart = rowStart;
        this.rowEnd = rowEnd;
    }

    public Set<Integer> getExcludedColumns() {
        return excluded;
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) {
        String name = localName.isEmpty() ? qName : localName;
        switch (name) {
            case "c":
                inCell = true;
                currentCellRef = attributes.getValue("r");
                cellType = attributes.getValue("t") == null ? "" : attributes.getValue("t");
                valueBuffer.setLength(0);
                currentRow0 = new CellReference(currentCellRef).getRow();
                if (currentRow0 > rowEnd) {
                    System.out.println("Stop parsing: reached rowEnd");
                    throw new StopParsingException();
                }
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
        if (inValue) {
            valueBuffer.append(ch, start, length);
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName) {
        String name = localName.isEmpty() ? qName : localName;
        switch (name) {
            case "c":
                if (inCell) {
                    String raw = valueBuffer.toString().trim();
                    String value = convertRawToValue(raw, cellType);
                    if (!value.isEmpty() && markers.stream().anyMatch(m -> m.trim().equalsIgnoreCase(value))) {
                        int col = new CellReference(currentCellRef).getCol();
                        CellRangeAddress region = allRegions.stream()
                                .filter(r -> r.isInRange(currentRow0, col))
                                .findFirst()
                                .orElse(null);
                        if (region != null) {
                            IntStream.rangeClosed(region.getFirstColumn(), region.getLastColumn()).forEach(excluded::add);
                        } else {
                            excluded.add(col);
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