package com.abysstutu.metadata;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.CustomProperties;
import org.apache.poi.hpsf.PropertySet;
import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.Entry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/** Small command-line bridge for legacy Office standard metadata streams. */
public final class LegacyMetadataHelper {
    private static final List<String> KEYS = Arrays.asList(
        "title", "subject", "creator", "lastModifiedBy", "keywords", "description",
        "revision", "created", "modified", "lastPrinted", "totalTime", "application",
        "category", "company", "manager", "hyperlinkBase", "contentStatus", "template"
    );
    private static final DateTimeFormatter TIME = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    private static final String LINKBASE = "_PID_LINKBASE";

    private LegacyMetadataHelper() { }

    public static void main(String[] args) throws Exception {
        if (args.length < 2) throw new IllegalArgumentException("Usage: read <input> | write <input> <output> <clear>");
        if ("read".equals(args[0])) {
            read(Paths.get(args[1]));
            return;
        }
        if ("write".equals(args[0]) && args.length == 4) {
            write(Paths.get(args[1]), Paths.get(args[2]), Boolean.parseBoolean(args[3]), readInput());
            return;
        }
        throw new IllegalArgumentException("Invalid command");
    }

    private static void read(Path input) throws Exception {
        try (InputStream file = Files.newInputStream(input); POIFSFileSystem fs = new POIFSFileSystem(file)) {
            DirectoryNode root = fs.getRoot();
            SummaryInformation summary = summary(root);
            DocumentSummaryInformation docSummary = documentSummary(root);
            Map<String, String> values = new LinkedHashMap<>();
            values.put("title", string(summary == null ? null : summary.getTitle()));
            values.put("subject", string(summary == null ? null : summary.getSubject()));
            values.put("creator", string(summary == null ? null : summary.getAuthor()));
            values.put("lastModifiedBy", string(summary == null ? null : summary.getLastAuthor()));
            values.put("keywords", string(summary == null ? null : summary.getKeywords()));
            values.put("description", string(summary == null ? null : summary.getComments()));
            values.put("revision", string(summary == null ? null : summary.getRevNumber()));
            values.put("created", time(summary == null ? null : summary.getCreateDateTime()));
            values.put("modified", time(summary == null ? null : summary.getLastSaveDateTime()));
            values.put("lastPrinted", time(summary == null ? null : summary.getLastPrinted()));
            values.put("totalTime", summary == null || summary.getEditTime() == 0 ? "" : String.valueOf(summary.getEditTime() / 600000000L));
            values.put("application", string(summary == null ? null : summary.getApplicationName()));
            values.put("category", string(docSummary == null ? null : docSummary.getCategory()));
            values.put("company", string(docSummary == null ? null : docSummary.getCompany()));
            values.put("manager", string(docSummary == null ? null : docSummary.getManager()));
            CustomProperties custom = docSummary == null ? null : docSummary.getCustomProperties();
            values.put("hyperlinkBase", string(custom == null ? null : custom.get(LINKBASE)));
            values.put("contentStatus", string(docSummary == null ? null : docSummary.getContentStatus()));
            values.put("template", string(summary == null ? null : summary.getTemplate()));
            for (String key : KEYS) emit(key, values.getOrDefault(key, ""));
            if (custom != null) for (String name : custom.nameSet()) {
                if (!LINKBASE.equals(name)) emit(customKey(name), string(custom.get(name)));
            }
        }
    }

    private static void write(Path input, Path output, boolean clear, Map<String, String> values) throws Exception {
        try (InputStream file = Files.newInputStream(input); POIFSFileSystem fs = new POIFSFileSystem(file)) {
            DirectoryNode root = fs.getRoot();
            SummaryInformation summary = clear ? PropertySetFactory.newSummaryInformation() : summary(root);
            DocumentSummaryInformation docSummary = clear ? PropertySetFactory.newDocumentSummaryInformation() : documentSummary(root);
            if (summary == null) summary = PropertySetFactory.newSummaryInformation();
            if (docSummary == null) docSummary = PropertySetFactory.newDocumentSummaryInformation();

            summary.setTitle(value(values, "title", clear));
            summary.setSubject(value(values, "subject", clear));
            summary.setAuthor(value(values, "creator", clear));
            summary.setLastAuthor(value(values, "lastModifiedBy", clear));
            summary.setKeywords(value(values, "keywords", clear));
            summary.setComments(value(values, "description", clear));
            summary.setRevNumber(value(values, "revision", clear));
            summary.setApplicationName(value(values, "application", clear));
            summary.setTemplate(value(values, "template", clear));
            setDate(summary, "created", value(values, "created", clear));
            setDate(summary, "modified", value(values, "modified", clear));
            setDate(summary, "lastPrinted", value(values, "lastPrinted", clear));
            String duration = value(values, "totalTime", clear);
            String currentDuration = summary.getEditTime() == 0 ? "" : String.valueOf(summary.getEditTime() / 600000000L);
            if (!duration.equals(currentDuration)) {
                if (duration.isEmpty()) summary.removeEditTime();
                else {
                    long minutes = Long.parseLong(duration);
                    if (minutes < 0) throw new IllegalArgumentException("总编辑时长不能为负数");
                    summary.setEditTime(Math.multiplyExact(minutes, 600000000L));
                }
            }
            docSummary.setCategory(value(values, "category", clear));
            docSummary.setCompany(value(values, "company", clear));
            docSummary.setManager(value(values, "manager", clear));
            docSummary.setContentStatus(value(values, "contentStatus", clear));
            if (clear) {
                if (docSummary.getCustomProperties() != null) docSummary.removeCustomProperties();
            } else {
                CustomProperties custom = docSummary.getCustomProperties();
                if (custom == null) custom = new CustomProperties();
                String base = value(values, "hyperlinkBase", false);
                if (base.isEmpty()) custom.remove(LINKBASE);
                else custom.put(LINKBASE, base);
                Set<String> requested = new HashSet<>();
                for (Map.Entry<String, String> entry : values.entrySet()) {
                    if (!entry.getKey().startsWith("custom:")) continue;
                    String name = new String(Base64.getUrlDecoder().decode(entry.getKey().substring(7)), StandardCharsets.UTF_8);
                    if (LINKBASE.equals(name)) continue;
                    requested.add(name);
                    if (!Objects.equals(string(custom.get(name)), entry.getValue())) custom.put(name, entry.getValue());
                }
                for (String name : new ArrayList<>(custom.nameSet())) {
                    if (!LINKBASE.equals(name) && !requested.contains(name)) custom.remove(name);
                }
                if (custom.isEmpty()) {
                    if (docSummary.getCustomProperties() != null) docSummary.removeCustomProperties();
                }
                else docSummary.setCustomProperties(custom);
            }

            replace(root, SummaryInformation.DEFAULT_STREAM_NAME, summary);
            replace(root, DocumentSummaryInformation.DEFAULT_STREAM_NAME, docSummary);
            try (OutputStream target = Files.newOutputStream(output, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
                fs.writeFilesystem(target);
            }
        }
    }

    private static SummaryInformation summary(DirectoryNode root) throws Exception {
        if (!root.hasEntry(SummaryInformation.DEFAULT_STREAM_NAME)) return null;
        try (DocumentInputStream stream = root.createDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME)) {
            PropertySet properties = PropertySetFactory.create(stream);
            return properties instanceof SummaryInformation ? (SummaryInformation) properties : null;
        }
    }

    private static DocumentSummaryInformation documentSummary(DirectoryNode root) throws Exception {
        if (!root.hasEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME)) return null;
        try (DocumentInputStream stream = root.createDocumentInputStream(DocumentSummaryInformation.DEFAULT_STREAM_NAME)) {
            PropertySet properties = PropertySetFactory.create(stream);
            return properties instanceof DocumentSummaryInformation ? (DocumentSummaryInformation) properties : null;
        }
    }

    private static void replace(DirectoryNode root, String name, PropertySet properties) throws Exception {
        if (root.hasEntry(name)) {
            Entry old = root.getEntry(name);
            old.delete();
        }
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        properties.write(bytes);
        root.createDocument(name, new ByteArrayInputStream(bytes.toByteArray()));
    }

    private static String value(Map<String, String> values, String key, boolean clear) {
        // HPSF cannot serialize null string properties. An empty string is the
        // correct cleared value and keeps the property stream structurally valid.
        return clear ? "" : values.getOrDefault(key, "");
    }

    private static String string(Object value) { return value == null ? "" : String.valueOf(value); }

    private static String time(Date value) {
        return value == null ? "" : TIME.format(value.toInstant().atZone(ZoneId.systemDefault()));
    }

    private static void setDate(SummaryInformation summary, String key, String value) {
        Date current;
        switch (key) {
            case "created": current = summary.getCreateDateTime(); break;
            case "modified": current = summary.getLastSaveDateTime(); break;
            case "lastPrinted": current = summary.getLastPrinted(); break;
            default: throw new IllegalArgumentException(key);
        }
        if (value.equals(time(current))) return;
        Date date = value.isEmpty() ? null : Date.from(LocalDateTime.parse(value, TIME).atZone(ZoneId.systemDefault()).toInstant());
        switch (key) {
            case "created": if (date == null) summary.removeCreateDateTime(); else summary.setCreateDateTime(date); break;
            case "modified": if (date == null) summary.removeLastSaveDateTime(); else summary.setLastSaveDateTime(date); break;
            case "lastPrinted": if (date == null) summary.removeLastPrinted(); else summary.setLastPrinted(date); break;
            default: throw new IllegalArgumentException(key);
        }
    }

    private static String customKey(String name) {
        return "custom:" + Base64.getUrlEncoder().withoutPadding().encodeToString(name.getBytes(StandardCharsets.UTF_8));
    }

    private static Map<String, String> readInput() throws IOException {
        Map<String, String> values = new HashMap<>();
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(System.in, StandardCharsets.UTF_8))) {
            String line;
            while ((line = reader.readLine()) != null) {
                int tab = line.indexOf('\t');
                if (tab < 1) continue;
                values.put(line.substring(0, tab), new String(Base64.getDecoder().decode(line.substring(tab + 1)), StandardCharsets.UTF_8));
            }
        }
        return values;
    }

    private static void emit(String key, String value) {
        System.out.println(key + "\t" + Base64.getEncoder().encodeToString(value.getBytes(StandardCharsets.UTF_8)));
    }
}
