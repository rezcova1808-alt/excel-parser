package com.example.excelparesernew;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DPANPatternMatcher {

    public static final String DPAN_CANDIDATE_PATTERN = "\\b[\\da-fA-F\\s]{10,}[ABC]\\s?\\d{4}\\b";
    public static final Pattern dpanCandidatePattern = Pattern.compile(DPAN_CANDIDATE_PATTERN);
    public static final String DPAN_EXACT_PATTERN = "\\b\\d{6,}[a-fA-F\\d]{5}[ABC]\\d{4}\\b";
    public static final Pattern dpanPattern = Pattern.compile(DPAN_EXACT_PATTERN);

    public List<String> findDPANCandidates(String text) {
        List<String> candidates = new ArrayList<>();

        if (text == null || text.trim().isEmpty()) {
            return candidates;
        }

        Matcher matcher = dpanCandidatePattern.matcher(text);

        while (matcher.find()) {
            String candidate = matcher.group();
            candidates.add(candidate);
        }

        return candidates;
    }

    public boolean isExactDPAN(String candidate) {
        if (candidate == null) return false;
        String clean = candidate.replaceAll("\\s+", "");
        //return clean.matches(String.valueOf(dpanPattern));
        return dpanPattern.matcher(clean).matches();
    }

    public String replaceDPANsWithMap(String text, Map<String, String> replacementMap) {
        if (text == null || replacementMap == null || replacementMap.isEmpty()) {
            return text;
        }

        String result = text;
        List<String> candidates = findDPANCandidates(text);

        for (String candidate : candidates) {
            if (!isExactDPAN(candidate)) {
                continue;
            }

            String replacement = findReplacementInMap(candidate, replacementMap);
            if (replacement != null) {
                result = result.replace(candidate, replacement);
            }
        }

        return result;
    }

    private String findReplacementInMap(String dpan, Map<String, String> replacementMap) {
        if (replacementMap.containsKey(dpan)) {
            return replacementMap.get(dpan);
        }

        String cleanDPAN = dpan.replaceAll("\\s+", "");
        if (replacementMap.containsKey(cleanDPAN)) {
            return replacementMap.get(cleanDPAN);
        }

        return null;
    }

    /**
     * Быстрая проверка: содержит ли текст хотя бы один кандидат в DPAN.
     * Использует общий паттерн dpanCandidatePattern, НЕ создаёт список.
     * Останавливается на первом же совпадении.
     */
    public boolean containsDPANCandidate(String text) {
        if (text == null || text.isEmpty()) {
            return false;
        }
        return dpanCandidatePattern.matcher(text).find();
    }
}