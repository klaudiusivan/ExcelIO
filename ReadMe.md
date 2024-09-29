# ExcelIO

ExcelIO is a Swift framework for reading and writing Excel XLSX files. It allows developers to integrate Excel file manipulation into their Swift applications easily, with support for handling nested objects, arrays, and basic cell styling.

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Getting Started](#getting-started)
  - [Reading XLSX Files](#reading-xlsx-files)
  - [Writing XLSX Files](#writing-xlsx-files)
- [Handling Complex Data Structures](#handling-complex-data-structures)
- [Styles and Formatting](#styles-and-formatting)
- [API Reference](#api-reference)
  - [XLSXWorkbook](#xlsxworkbook)
  - [XLSXWorksheet](#xlsxworksheet)
  - [XLSXCell](#xlsxcell)
  - [XLSXReader](#xlsxreader)
  - [XLSXWriter](#xlsxwriter)
- [Examples](#examples)
- [Contributing](#contributing)
- [License](#license)

## Features

- **Read XLSX Files**: Extract data from Excel files, including multiple sheets and cell values.
- **Write XLSX Files**: Create Excel files with multiple sheets, cells, and data types.
- **Encode/Decode `Encodable` Objects**: Convert Swift `Encodable` objects to Excel files and back to JSON format.
- **Handle Complex Data Structures**: Support for nested objects and arrays within data models.
- **Styles and Formatting**: Basic support for cell styles and date formatting.
- **Pure Swift Implementation**: No external dependencies; built using native Swift frameworks.

## Installation

### Swift Package Manager

ExcelIO can be added to your project using Swift Package Manager.

1. **Open your project in Xcode**.
2. Go to **`File > Swift Packages > Add Package Dependency...`**.
3. Enter the **repository URL**:

   ```
   https://github.com/klaudiusivan/ExcelIO.git
   ```

4. Select the latest version and **add it to your project**.

## Getting Started

This section provides an overview of the basic usage of ExcelIO for reading and writing Excel files.

### Reading XLSX Files

To read an XLSX file and convert it to JSON:

```swift
import ExcelIO

do {
    let fileURL = URL(fileURLWithPath: "/path/to/your/file.xlsx")
    let jsonData = try XLSXReader.readJSON(from: fileURL)
    if let jsonString = String(data: jsonData, encoding: .utf8) {
        print("JSON Output:")
        print(jsonString)
    }
} catch {
    print("Error reading XLSX: \(error)")
}
```

### Writing XLSX Files

#### From an `Encodable` Object

You can write any `Encodable` Swift object to an XLSX file:

```swift
import ExcelIO

struct Person: Encodable {
    var name: String
    var email: String
}

let person = Person(name: "John Doe", email: "john.doe@example.com")

do {
    let fileURL = URL(fileURLWithPath: "/path/to/save/person.xlsx")
    try XLSXWriter.save(encodableObject: person, to: fileURL, sheetName: "Person")
    print("Workbook saved to \(fileURL.path)")
} catch {
    print("Error writing XLSX: \(error)")
}
```

#### From an Array of `Encodable` Objects

```swift
let people = [
    Person(name: "John Doe", email: "john.doe@example.com"),
    Person(name: "Jane Smith", email: "jane.smith@example.com")
]

do {
    let fileURL = URL(fileURLWithPath: "/path/to/save/people.xlsx")
    try XLSXWriter.save(encodableArray: people, to: fileURL, sheetName: "People")
    print("Workbook saved to \(fileURL.path)")
} catch {
    print("Error writing XLSX: \(error)")
}
```

## Handling Complex Data Structures

ExcelIO can handle nested objects and arrays within your data models.

**Example with Nested Structures:**

```swift
struct CV: Encodable {
    var fullName: String
    var contactInformation: ContactInformation
    var experiences: [Experience]
}

struct ContactInformation: Encodable {
    var email: String
    var phone: String
}

struct Experience: Encodable {
    var jobTitle: String
    var companyName: String
}

let cv = CV(
    fullName: "John Doe",
    contactInformation: ContactInformation(email: "john.doe@example.com", phone: "123-456-7890"),
    experiences: [
        Experience(jobTitle: "Software Engineer", companyName: "Tech Corp"),
        Experience(jobTitle: "Senior Developer", companyName: "Innovate Ltd")
    ]
)

do {
    let fileURL = URL(fileURLWithPath: "/path/to/save/CV.xlsx")
    try XLSXWriter.save(encodableObject: cv, to: fileURL, sheetName: "CV")
    print("CV saved to \(fileURL.path)")
} catch {
    print("Error writing XLSX: \(error)")
}
```

## Styles and Formatting

ExcelIO supports basic cell styling and date formatting.

**Example with Dates:**

```swift
struct Event: Encodable {
    var name: String
    var date: Date
}

let event = Event(name: "Conference", date: Date())

do {
    let fileURL = URL(fileURLWithPath: "/path/to/save/event.xlsx")
    try XLSXWriter.save(encodableObject: event, to: fileURL, sheetName: "Event")
    print("Event saved to \(fileURL.path)")
} catch {
    print("Error writing XLSX: \(error)")
}
```

Date values are automatically formatted as dates in Excel.

## API Reference

### XLSXWorkbook

Represents an Excel workbook containing multiple sheets.

- **Methods**:
  - `addSheet(from dictionary: [String: Any], sheetName: String)`
  - `addSheet(from array: [[String: Any]], sheetName: String)`
  - `toJSON() -> [String: Any]`

### XLSXWorksheet

Represents a worksheet within a workbook.

- **Methods**:
  - `setCell(at reference: String, value: Any)`
  - `populate(with dictionary: [String: Any], parentWorkbook: XLSXWorkbook)`
  - `populate(with array: [[String: Any]], parentWorkbook: XLSXWorkbook)`

### XLSXCell

Represents a single cell in a worksheet.

### XLSXReader

Handles reading XLSX files.

- **Methods**:
  - `read(from fileURL: URL) throws -> XLSXWorkbook`
  - `readJSON(from fileURL: URL) throws -> Data`

### XLSXWriter

Handles writing XLSX files.

- **Methods**:
  - `save<T: Encodable>(encodableObject: T, to fileURL: URL, sheetName: String = "Sheet1") throws`
  - `save<T: Encodable>(encodableArray: [T], to fileURL: URL, sheetName: String = "Sheet1") throws`

## Examples

### Example 1: Reading an XLSX File

```swift
do {
    let fileURL = URL(fileURLWithPath: "/path/to/your/data.xlsx")
    let workbook = try XLSXReader.read(from: fileURL)
    for sheet in workbook.sheets {
        print("Sheet: \(sheet.name)")
        for cell in sheet.cells.values {
            print("\(cell.reference): \(cell.value)")
        }
    }
} catch {
    print("Error reading XLSX: \(error)")
}
```

### Example 2: Writing Custom Data to XLSX

```swift
let workbook = XLSXWorkbook()
let sheet = XLSXWorksheet(name: "Custom Data")
sheet.setCell(at: "A1", value: "Header1")
sheet.setCell(at: "B1", value: "Header2")
sheet.setCell(at: "A2", value: "Value1")
sheet.setCell(at: "B2", value: "Value2")
workbook.addSheet(sheet)

do {
    let fileURL = URL(fileURLWithPath: "/path/to/save/custom_data.xlsx")
    try XLSXWriter.save(workbook: workbook, to: fileURL)
    print("Workbook saved to \(fileURL.path)")
} catch {
    print("Error writing XLSX: \(error)")
}
```

## Contributing

Contributions are welcome! If you'd like to contribute to **ExcelIO**, please follow these steps:

1. **Fork the repository** on GitHub.
2. **Create a new branch** for your feature or bug fix.
3. **Commit your changes** with clear and descriptive messages.
4. **Push your branch** to your forked repository.
5. **Submit a pull request** detailing your changes.

## License

ExcelIO is released under the MIT License.
