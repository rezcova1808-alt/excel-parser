package com.example.demo;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStrings;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;


public class ExcelDPANCollector {

    private final DPANPatternMatcher patternMatcher;
    public static final Set<String> EXCLUDED_HEADERS = new HashSet<>(Arrays.asList("Назначение платежа"));

    public ExcelDPANCollector() {
        this.patternMatcher = new DPANPatternMatcher();
    }

    public ProcessFileResult collectAllDPANCandidatesMemory(File excelFile, boolean skipColumns) {
        ProcessFileResult processFileResult = new ProcessFileResult();
        System.out.println("start process columns and regions: " + LocalDateTime.now());
        Map<String, Set<Integer>> excludedColsPerSheet = skipColumns
                ? findExcludedColumnsFromMarkers(excelFile, EXCLUDED_HEADERS, 0, 0)
                : Collections.emptyMap();
        Map<String, Set<String>> collectValuesFromSheet = collectValuesPerSheet(excelFile, excludedColsPerSheet, patternMatcher);
        Set<String> allValues = collectValuesFromSheet.values()
                .stream()
                .filter(Objects::nonNull)
                .flatMap(Set::stream)
                .collect(Collectors.toSet());
        System.out.println("end process values: " + LocalDateTime.now());
        processFileResult.setExcludedColsPerSheet(excludedColsPerSheet);
        processFileResult.setCollectValuesFromFile(allValues);
        return processFileResult;
    }

    private Map<String, Set<Integer>> findExcludedColumnsFromMarkers(File file,
                                                                     Set<String> markers,
                                                                     int rowStart,
                                                                     int rowEnd) {
        if (markers == null || markers.isEmpty() || rowStart < 0) {
            return Collections.emptyMap();
        }
        System.out.println("start process regions: " + LocalDateTime.now());
        Map<String, List<CellRangeAddress>> sheetRegions = loadAllRegions(file);
        System.out.println("end process regions: " + LocalDateTime.now());
        Map<String, Set<Integer>> excludedColumnsMap = new LinkedHashMap<>();
        try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ)) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStrings sst = reader.getSharedStringsTable();
            SAXParserFactory factory = SAXParserFactory.newInstance();
            factory.setNamespaceAware(true);
            XMLReader parser = factory.newSAXParser().getXMLReader();

            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
            while (sheets.hasNext()) {
                try (InputStream sheetStream = sheets.next()) {
                    String sheetName = sheets.getSheetName();
                    List<CellRangeAddress> regions = sheetRegions.get(sheetName);
                    MarkersHandler handler = new MarkersHandler(sst, regions, markers, rowStart, rowEnd);
                    parser.setContentHandler(handler);
                    try {
                        parser.parse(new InputSource(sheetStream));
                    } catch (StopParsingException e) { /* нормальная остановка */ }
                    excludedColumnsMap.put(sheetName, handler.getExcludedColumns());
                }
            }
        } catch (IOException | OpenXML4JException | ParserConfigurationException | SAXException e) {
            throw new IllegalStateException(e);
        }
        return excludedColumnsMap;
    }

    private static Map<String, List<CellRangeAddress>> loadAllRegions(File file) {
        Map<String, List<CellRangeAddress>> regionsMap = new LinkedHashMap<>();
        try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ)) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStrings sst = reader.getSharedStringsTable();
            SAXParserFactory factory = SAXParserFactory.newInstance();
            factory.setNamespaceAware(true);
            XMLReader parser = factory.newSAXParser().getXMLReader();

            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
            while (sheets.hasNext()) {
                try (InputStream sheetStream = sheets.next()) {
                    String sheetName = sheets.getSheetName();
                    RegionsHandler handler = new RegionsHandler();
                    parser.setContentHandler(handler);
                    parser.parse(new InputSource(sheetStream));
                    regionsMap.put(sheetName, handler.getRegions());
                }
            }
        } catch (IOException | OpenXML4JException | ParserConfigurationException | SAXException e) {
            throw new IllegalStateException(e);
        }
        return regionsMap;
    }

    private static Map<String, Set<String>> collectValuesPerSheet(File file,
                                                                  Map<String, Set<Integer>> excludedColsPerSheet,
                                                                  DPANPatternMatcher patternMatcher) {
        Map<String, Set<String>> valuesPerSheet = new LinkedHashMap<>();
        try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ)) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStrings sst = reader.getSharedStringsTable();
            SAXParserFactory factory = SAXParserFactory.newInstance();
            factory.setNamespaceAware(true);
            XMLReader parser = factory.newSAXParser().getXMLReader();

            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
            while (sheets.hasNext()) {
                try (InputStream sheetStream = sheets.next()) {
                    String sheetName = sheets.getSheetName();
                    Set<String> sheetValues = new HashSet<>();

                    // Обработка имени листа
                    if (sheetName != null && !sheetName.isEmpty()) {
                        List<String> candidatesFromName = patternMatcher.findDPANCandidates(sheetName);
                        for (String candidate : candidatesFromName) {
                            if (patternMatcher.isExactDPAN(candidate)) {
                                sheetValues.add(candidate);
                            }
                        }
                    }

                    // Обработка ячеек листа
                    Set<Integer> excluded = excludedColsPerSheet.getOrDefault(sheetName, Collections.emptySet());
                    DpanSheetHandler handler = new DpanSheetHandler(sst, excluded, patternMatcher);
                    parser.setContentHandler(handler);
                    parser.parse(new InputSource(sheetStream));
                    sheetValues.addAll(handler.getValues());
                    valuesPerSheet.put(sheetName, sheetValues);
                }
            }
        } catch (IOException | OpenXML4JException | ParserConfigurationException | SAXException e) {
            throw new IllegalStateException(e);
        }
        return valuesPerSheet;
    }

}
