//
//  WorkbookTests.swift
//  ExcelIO
//
//  Created by Klaudius Ivan on 9/29/24.
//


import XCTest
import ExcelIO

final class WorkbookTests: XCTestCase {
    
    func testWorkbookAndWorksheetCreation() throws {
        let workbook = XLSXWorkbook()
        let sheet = XLSXWorksheet(name: "Test Sheet")
        workbook.addSheet(sheet)
        
        XCTAssertEqual(workbook.sheets.count, 1, "Workbook should contain one sheet")
        XCTAssertEqual(workbook.sheets.first?.name, "Test Sheet", "Sheet name should be 'Test Sheet'")
    }
    
    func testAddingMultipleSheets() throws {
        let workbook = XLSXWorkbook()
        let sheet1 = XLSXWorksheet(name: "Sheet1")
        let sheet2 = XLSXWorksheet(name: "Sheet2")
        
        workbook.addSheet(sheet1)
        workbook.addSheet(sheet2)
        
        XCTAssertEqual(workbook.sheets.count, 2, "Workbook should contain two sheets")
        XCTAssertEqual(workbook.sheets[1].name, "Sheet2", "Second sheet name should be 'Sheet2'")
    }
}
