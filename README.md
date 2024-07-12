# BibTeX2Word
BibTeX2Word is a VBA script that allows you to import BibTeX citations into Microsoft Word's bibliography. This script extracts citation data from BibTeX entries and converts it into the appropriate XML format for Word, enabling seamless integration of your references.

## Features
- Supports multiple citation types including Journal Articles, Books, Book Sections, Conference Proceedings, and Reports.
- Automatically generates a unique citation tag.
- Adds citations to Word's bibliography.
- Inserts citations at the current cursor position in the document.

## Installation
1. Open Microsoft Word.
2. Press `Alt + F11` to open the VBA editor.
3. Go to `Insert` -> `Module` to create a new module.
4. Copy and paste the VBA script from this repository into the module.
5. Close the VBA editor.

## Usage
Bind one of the following functions to a shortcut key:

1. `TransformIntoCitation`: Transform selected text into a citation
2. `PasteIntoCitation`: Transform clipboard content into a citation

`PasteIntoCitation` requires `Microsoft Forms 2.0 Object Library`. If you cannot find the tool in the reference list, import FM20.DLL file from the system32 directory.

## License
This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program. If not, see https://www.gnu.org/licenses/.

## Contributing
Contributions are welcome! Please open an issue or submit a pull request on GitHub.
