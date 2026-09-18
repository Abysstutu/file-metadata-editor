# Legacy Office metadata helper

This Maven project wraps Apache POI HPSF/POIFS so the Python desktop app can read and rewrite the standard property streams inside `.doc`, `.xls`, and `.ppt` files without converting their document content.

Build with Java 11+ and Maven:

```bash
mvn package
```

Copy `target/legacy-metadata-helper-1.0.0.jar` to `legacy_helper/legacy-metadata-helper.jar` before packaging the Python application.

The helper uses a line-oriented, Base64-encoded protocol over standard input/output to preserve Chinese and multiline metadata safely.
