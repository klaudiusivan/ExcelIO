//
//  FileIOTests.swift
//  ExcelIO
//
//  Created by Klaudius Ivan on 9/29/24.
//


import XCTest
import ExcelIO

final class FileIOTests: XCTestCase {
    
    func testWritingAndSavingWorkbook() throws {
        let workbook = XLSXWorkbook()
        let sheet = XLSXWorksheet(name: "Sheet1")
        sheet.setCell(at: "A1", value: "Title")
        sheet.setCell(at: "B1", value: "Description")
        sheet.setCell(at: "A2", value: "Swift")
        sheet.setCell(at: "B2", value: "Programming Language")
        workbook.addSheet(sheet)
        
        let fileURL = FileManager.default.temporaryDirectory.appendingPathComponent("WorkbookSaveTest.xlsx")
        
        // Check if file exists and remove it
        if FileManager.default.fileExists(atPath: fileURL.path) {
            try FileManager.default.removeItem(at: fileURL)
        }
        
        try XLSXWriter.save(workbook: workbook, to: fileURL)

        XCTAssertTrue(FileManager.default.fileExists(atPath: fileURL.path), "Workbook should be saved to file system")
    }
    
    func testReadingExistingXLSX() throws {
        // Step 1: Create a mock workbook
        let workbook = XLSXWorkbook()
        let sheet = XLSXWorksheet(name: "Mock Sheet")
        sheet.setCell(at: "A1", value: "Hello")
        sheet.setCell(at: "B1", value: "World")
        workbook.addSheet(sheet)
        
        // Step 2: Save the workbook as "Sample.xlsx" in a temporary directory
        let fileURL = FileManager.default.temporaryDirectory.appendingPathComponent("Sample.xlsx")
        
        // Check if file exists and remove it to avoid conflicts
        if FileManager.default.fileExists(atPath: fileURL.path) {
            try FileManager.default.removeItem(at: fileURL)
        }
        
        // Save the mock file
        try XLSXWriter.save(workbook: workbook, to: fileURL)
        
        // Step 3: Read the mock file using XLSXReader
        let readWorkbook = try XLSXReader.read(from: fileURL)
        
        // Step 4: Perform the test assertions
        XCTAssertTrue(readWorkbook.sheets.count > 0, "Workbook should contain at least one sheet")
        
        if let firstSheet = readWorkbook.sheets.first {
            XCTAssertEqual(firstSheet.name, "Mock Sheet", "First sheet should be named 'Mock Sheet'")
            XCTAssertEqual(firstSheet.getCell(at: "A1")?.value, "Hello", "Cell A1 should contain 'Hello'")
            XCTAssertEqual(firstSheet.getCell(at: "B1")?.value, "World", "Cell B1 should contain 'World'")
        }
    }


}
