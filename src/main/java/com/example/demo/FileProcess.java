package com.example.demo;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.IntStream;

public class FileProcess {
    public static void main(String[] args) throws Exception {
        ExcelDPANCollector collector = new ExcelDPANCollector();
        Path in = Paths.get("src/main/resources/testdetokenText_big.xlsx");
        //Path in = Paths.get("src/main/resources/UltimateTest.xlsx");
        Path out = Paths.get("src/main/resources/testdetokenText_result.xlsx");
        System.out.println("start collect:" + LocalDateTime.now());
        ProcessFileResult result = collector.collectAllDPANCandidatesMemory(in.toFile(), true);
        Set<String> candidates = result.getCollectValuesFromFile();
        System.out.println(candidates);
        Map<String, String> replacementMap = collectReplacementMap(candidates);
        Map<String, Set<Integer>> excludedColsPerSheet = result.getExcludedColsPerSheet();
        ExcelDPANReplacer excelDPANReplacer = new ExcelDPANReplacer();
        System.out.println("start replace:" + LocalDateTime.now());
        excelDPANReplacer.replaceAllDpans(in, out, replacementMap, excludedColsPerSheet);
        System.out.println("end replace:" + LocalDateTime.now());
    }

    static Map<String, String> collectReplacementMap(Set<String> candidates) {
        List<String> originalDpans = new ArrayList<>(candidates);
        String[] dpans = originalDpans.stream()
                .map(dpan -> dpan.replaceAll("\\s+", ""))//если не очистить по пробелам, то мы можем влететь в ошибку!!!
                .toArray(String[]::new);
        String[] pansArray = getPans(dpans);

        Map<String, String> replacementMap = new HashMap<>();
        IntStream.range(0, pansArray.length).forEach(i -> {
            String originalDpan = originalDpans.get(i);
            String pan = pansArray[i];
            if (originalDpan != null && pan != null && !pan.isEmpty() && !pan.equals(originalDpans.get(i).replaceAll("\\s+", ""))) {
                String panWithSpaces = restoreSpaces2(originalDpan, pan);
                replacementMap.put(originalDpan, panWithSpaces);
            }
        });
        return replacementMap;
    }

    public static String[] getPans(String[] dpans) {
        System.out.println("in dpan method");
        if (dpans == null) {
            return new String[0];
        } else {
            List<String> list = new ArrayList<>();
            for (String dpan : dpans) {
                if (dpan == null) {
                    list.add(null);
                } else {
                    String cleanedDpan = dpan.trim();
                    if (cleanedDpan.length() >= 4) {
                        cleanedDpan = "pan_" + cleanedDpan.substring(4);
                    }
                    list.add(cleanedDpan);
                }
            }
            return list.toArray(new String[0]);
        }
    }

    public static String restoreSpaces2(String originalWithSpaces, String valueWithoutSpaces) {
        if (originalWithSpaces == null || valueWithoutSpaces == null) {
            return valueWithoutSpaces;
        }

        StringBuilder result = new StringBuilder();
        int valueIndex = 0;

        for (int i = 0; i < originalWithSpaces.length(); i++) {
            char c = originalWithSpaces.charAt(i);

            if (Character.isWhitespace(c)) {
                result.append(c);
            } else if (valueIndex < valueWithoutSpaces.length()) {
                result.append(valueWithoutSpaces.charAt(valueIndex));
                valueIndex++;
            } else {
                break;
            }
        }

        if (valueIndex < valueWithoutSpaces.length()) {
            result.append(valueWithoutSpaces.substring(valueIndex));
        }

        return result.toString();
    }
}