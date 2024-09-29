# ExcelIO

ExcelIO is a Swift framework for reading and writing Excel `.xlsx` files. It allows developers to integrate Excel file manipulation into their Swift applications, with support for handling nested objects, arrays, and decoding directly into Swift models.

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Getting Started](#getting-started)
  - [Reading XLSX Files](#reading-xlsx-files)
  - [Writing XLSX Files](#writing-xlsx-files)
  - [Decoding XLSX Files into Swift Models](#decoding-xlsx-files-into-swift-models)
- [Handling Complex Data Structures](#handling-complex-data-structures)
- [Styles and Formatting](#styles-and-formatting)
- [API Reference](#api-reference)
  - [XLSXWorkbook](#xlsxworkbook)
  - [XLSXWorksheet](#xlsxworksheet)
  - [XLSXCell](#xlsxcell)
  - [XLSXReader](#xlsxreader)
  - [XLSXWriter](#xlsxwriter)
  - [XLSXDecoder](#xlsxdecoder)
- [Examples](#examples)
  - [Creating a Simple Workbook](#example-1-creating-a-simple-workbook)
  - [Parsing an Existing XLSX File](#example-2-parsing-an-existing-xlsx-file)
  - [Decoding XLSX to Swift Object](#example-3-decoding-xlsx-to-swift-object)
- [Contributing](#contributing)
- [License](#license)

## Features

- **Read XLSX Files**: Extract data from Excel files, including multiple sheets, shared strings, and cell values.
- **Write XLSX Files**: Create Excel files with multiple sheets and cells, supporting various data types.
- **Decode XLSX Files into Swift Models**: Seamlessly decode an `.xlsx` file into Swift `Decodable` objects.
- **Handle Nested Structures**: Manage complex data structures with nested objects and arrays.
- **Pure Swift Implementation**: Uses native Swift libraries with `ZIPFoundation` for handling zip archives.

## Installation

### Swift Package Manager

You can integrate ExcelIO into your project using Swift Package Manager.

1. **Open your project in Xcode**.
2. Go to **`File > Swift Packages > Add Package Dependency...`**.
3. Enter the **repository URL**:

   ```
   https://github.com/klaudiusivan/ExcelIO.git
   ```

4. Select the latest version and **add it to your project**.

## Getting Started

Import the framework in your Swift file:

```swift
import ExcelIO
```

### Reading XLSX Files

To read an `.xlsx` file and convert its contents into a `XLSXWorkbook` object:

```swift
import ExcelIO

do {
    let fileURL = URL(fileURLWithPath: "/path/to/your/file.xlsx")
    let workbook = try XLSXReader.read(from: fileURL)
    
    // Iterate over sheets
    for sheet in workbook.sheets {
        print("Sheet Name: \(sheet.name)")
        
        // Iterate over cells
        for cell in sheet.cells.values {
            print("Cell \(cell.reference): \(cell.value)")
        }
    }
} catch {
    print("Error reading XLSX: \(error)")
}
```

### Writing XLSX Files

To write data to an `.xlsx` file:

```swift
import ExcelIO

// Create a workbook
let workbook = XLSXWorkbook()

// Create a sheet and add cells
let sheet = XLSXWorksheet(name: "Sheet1")
sheet.setCell(at: "A1", value: "Name")
sheet.setCell(at: "B1", value: "Email")
sheet.setCell(at: "A2", value: "John Doe")
sheet.setCell(at: "B2", value: "john.doe@example.com")

// Add the sheet to the workbook
workbook.addSheet(sheet)

do {
    let fileURL = URL(fileURLWithPath: "/path/to/save/workbook.xlsx")
    try XLSXWriter.save(workbook: workbook, to: fileURL)
    print("Workbook saved to \(fileURL.path)")
} catch {
    print("Error writing XLSX: \(error)")
}
```

### Decoding XLSX Files into Swift Models

ExcelIO supports decoding `.xlsx` files directly into Swift models that conform to `Decodable`.

**Example with a Swift Model:**

```swift
struct Person: Decodable {
    var name: String
    var age: Int
    var email: String
}

let fileURL = URL(fileURLWithPath: "/path/to/your/file.xlsx")

do {
    // Decode the XLSX file into the Person model
    let person = try XLSXReader.read(Person.self, from: fileURL)
    print("Name: \(person.name)")
    print("Age: \(person.age)")
    print("Email: \(person.email)")
} catch {
    print("Error decoding XLSX: \(error)")
}
```

## Handling Complex Data Structures

ExcelIO can handle nested objects and arrays within your data models, supporting conversion between Swift objects and Excel files.

**Example with Nested Structures:**

```swift
struct Contact: Decodable {
    var name: String
    var phone: String
    var email: String
}

struct AddressBook: Decodable {
    var contacts: [Contact]
}

let fileURL = URL(fileURLWithPath: "/path/to/your/addressbook.xlsx")

do {
    let addressBook = try XLSXReader.read(AddressBook.self, from: fileURL)
    print("Contacts: \(addressBook.contacts)")
} catch {
    print("Error decoding XLSX: \(error)")
}
```

## API Reference

### XLSXWorkbook

Represents an Excel workbook containing multiple sheets.

- **Properties**:
  - `sheets`: The array of `XLSXWorksheet` objects in the workbook.
  - `sharedStrings`: The array of shared strings for the workbook.
- **Methods**:
  - `addSheet(_ sheet: XLSXWorksheet)`: Adds a sheet to the workbook.
  - `toJSON() -> [String: Any]`: Converts the workbook to a JSON-compatible dictionary.

### XLSXWorksheet

Represents a worksheet within a workbook.

- **Properties**:
  - `name`: The name of the worksheet.
  - `cells`: A dictionary mapping cell references to `XLSXCell` objects.
- **Methods**:
  - `setCell(at reference: String, value: String)`: Sets the value of a cell at the given reference.
  - `getCell(at reference: String) -> XLSXCell?`: Retrieves the cell at the given reference.
  - `toJSON() -> [[String: Any]]`: Converts the worksheet into a JSON-compatible array.

### XLSXCell

Represents a single cell in a worksheet.

- **Properties**:
  - `reference`: The cell reference (e.g., "A1").
  - `value`: The cell's value.
  - `type`: The cell's data type.

### XLSXReader

Handles reading and parsing XLSX files.

- **Methods**:
  - `read(from fileURL: URL) throws -> XLSXWorkbook`: Reads an XLSX file and returns a workbook object.
  - `read<T: Decodable>(_ type: T.Type, from fileURL: URL) throws -> T`: Reads an XLSX file and decodes it into a `Decodable` Swift object.

### XLSXWriter

Handles writing data to XLSX files.

- **Methods**:
  - `save(workbook: XLSXWorkbook, to fileURL: URL) throws`: Writes the workbook data to an XLSX file.

### XLSXDecoder

Handles decoding an XLSX workbook into a `Decodable` object.

- **Initialization**:
  - `init(workbookDict: [String: Any])`: Creates an `XLSXDecoder` with a workbook dictionary.
- **Usage**:
  - Decodes the workbook's data into Swift types, supporting all typical types (`String`, `Int`, `Double`, `Bool`, `Date`, `Data`, and nested containers).

## Examples

### Example 1: Creating a Simple Workbook

```swift
import ExcelIO

let workbook = XLSXWorkbook()
let sheet = XLSXWorksheet(name: "Test Sheet")

sheet.setCell(at: "A1", value: "Header1")
sheet.setCell(at: "B1", value: "Header2")
sheet.setCell(at: "A2", value: "Value1")
sheet.setCell(at: "B2", value: "Value2")

workbook.addSheet(sheet)

do {
    let fileURL = URL(fileURLWithPath: "/path/to/save/testWorkbook.xlsx")
    try XLSXWriter.save(workbook: workbook, to: fileURL)
    print("Workbook saved to \(fileURL.path)")
} catch {
    print("Error saving workbook: \(error)")
}
```

### Example 2: Parsing an Existing XLSX File

```swift
import ExcelIO

do {
    let fileURL = URL(fileURLWithPath: "/path/to/existing/file.xlsx")
    let workbook = try XLSXReader.read(from: fileURL)
    
    for sheet in workbook.sheets {
        print("Sheet: \(sheet.name)")
        for cell in sheet.cells.values {
            print("Cell \(cell.reference): \(cell.value)")
        }
    }
} catch {
    print("Error reading XLSX file: \(error)")
}
```

### Example 3

: Decoding XLSX to Swift Object

```swift
import ExcelIO

struct Employee: Decodable {
    var name: String
    var position: String
    var salary: Double
}

let fileURL = URL(fileURLWithPath: "/path/to/your/employees.xlsx")

do {
    let employee = try XLSXReader.read(Employee.self, from: fileURL)
    print("Name: \(employee.name)")
    print("Position: \(employee.position)")
    print("Salary: \(employee.salary)")
} catch {
    print("Error decoding XLSX: \(error)")
}
```

## Contributing

Contributions are welcome! To contribute to ExcelIO:

1. **Fork the repository** on GitHub.
2. **Create a new branch** for your feature or bug fix.
3. **Commit your changes** with clear messages.
4. **Push your branch** to your forked repository.
5. **Submit a pull request** detailing your changes.

## License

ExcelIO is released under the MIT License.
```
