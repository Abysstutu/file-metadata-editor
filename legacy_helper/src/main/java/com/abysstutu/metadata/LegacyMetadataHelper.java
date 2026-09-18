package com.abysstutu.metadata;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;

import org.apache.poi.hpsf.DocumentSummaryInformation;
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
        "category", "company", "manager", "hyperlinkBase", "contentStatus"
    );

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
            values.put("created", string(summary == null ? null : summary.getCreateDateTime()));
            values.put("modified", string(summary == null ? null : summary.getLastSaveDateTime()));
            values.put("lastPrinted", string(summary == null ? null : summary.getLastPrinted()));
            values.put("totalTime", string(summary == null ? null : summary.getEditTime()));
            values.put("application", string(summary == null ? null : summary.getApplicationName()));
            values.put("category", string(docSummary == null ? null : docSummary.getCategory()));
            values.put("company", string(docSummary == null ? null : docSummary.getCompany()));
            values.put("manager", string(docSummary == null ? null : docSummary.getManager()));
            // HyperlinkBase is not exposed by Apache POI's current HPSF API.
            values.put("hyperlinkBase", "");
            values.put("contentStatus", string(docSummary == null ? null : docSummary.getContentStatus()));
            for (String key : KEYS) emit(key, values.getOrDefault(key, ""));
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
            docSummary.setCategory(value(values, "category", clear));
            docSummary.setCompany(value(values, "company", clear));
            docSummary.setManager(value(values, "manager", clear));
            docSummary.setContentStatus(value(values, "contentStatus", clear));

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
