package com.example.excelparesernew;

import org.apache.poi.ss.util.CellRangeAddress;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.List;

public class RegionsHandler extends DefaultHandler {
    private final List<CellRangeAddress> regions = new ArrayList<>();

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attrs) {
        if ("mergeCell".equals(localName) || "mergeCell".equals(qName)) {
            String ref = attrs.getValue("ref");
            if (ref != null && ref.contains(":")) {
                regions.add(CellRangeAddress.valueOf(ref));
            }
        }
    }

    public List<CellRangeAddress> getRegions() {
        return regions;
    }
}
