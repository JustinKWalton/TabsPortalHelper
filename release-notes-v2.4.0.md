**Schema change: Subject column removed, new Issue column added, Element column repurposed.**

### What''s new
- BSI[4] Issue column added to the canonical TABS schema.
- Subject column has been dropped - the helper no longer writes the PDF /Subj field. The TABSportal Bluebeam profile already removed the Subject column display in v2.3.
- Element column (BSI[2]) now holds what used to be Subject content.
- Issue column (BSI[4]) now holds what used to be Element content.

### API changes (breaking)
The /clipboard/bluebeam-markup endpoint accepts new JSON keys: note, status, element, location, issue, contents, rcHtml.

Removed keys: subject, itemCategory.

### Required FlutterFlow update
The sendMarkupToBluebeam.dart custom action must be updated to send the new payload shape. Without that update, the new Element and Issue columns will paste empty.

### Fixed
- TrayApp.cs source rewritten with \u escapes for all special characters, eliminating the cp1252 mojibake that had repeatedly corrupted em-dashes in tray notifications and dialog titles.
