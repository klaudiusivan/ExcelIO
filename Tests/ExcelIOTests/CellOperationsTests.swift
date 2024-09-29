//
//  CellOperationsTests.swift
//  ExcelIO
//
//  Created by Klaudius Ivan on 9/29/24.
//


import XCTest
import ExcelIO

final class CellOperationsTests: XCTestCase {
    
    func testCellOperations() throws {
        let sheet = XLSXWorksheet(name: "Data Sheet")
        sheet.setCell(at: "A1", value: "Hello")
        sheet.setCell(at: "B2", value: "World")
        sheet.setCell(at: "C3", value: "123")

        XCTAssertEqual(sheet.getCell(at: "A1")?.value, "Hello", "Cell A1 should contain 'Hello'")
        XCTAssertEqual(sheet.getCell(at: "B2")?.value, "World", "Cell B2 should contain 'World'")
        XCTAssertEqual(sheet.getCell(at: "C3")?.value, "123", "Cell C3 should contain '123'")
        XCTAssertNil(sheet.getCell(at: "D4"), "Cell D4 should be nil as it is not set")
    }
    
    func testCellUpdate() throws {
        let sheet = XLSXWorksheet(name: "Data Sheet")
        sheet.setCell(at: "A1", value: "Initial Value")
        sheet.setCell(at: "A1", value: "Updated Value")

        XCTAssertEqual(sheet.getCell(at: "A1")?.value, "Updated Value", "Cell A1 should be updated to 'Updated Value'")
    }
}
