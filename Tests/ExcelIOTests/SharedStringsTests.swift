//
//  SharedStringsTests.swift
//  ExcelIO
//
//  Created by Klaudius Ivan on 9/29/24.
//


import XCTest
import ExcelIO

final class SharedStringsTests: XCTestCase {
    
    func testSharedStringsHandling() throws {
        let workbook = XLSXWorkbook()
        let index1 = workbook.addSharedString("Shared String 1")
        let index2 = workbook.addSharedString("Shared String 2")
        let duplicateIndex = workbook.addSharedString("Shared String 1") // Duplicate

        XCTAssertEqual(workbook.sharedStrings.count, 2, "Shared strings should not contain duplicates")
        XCTAssertEqual(index1, 0, "Index of 'Shared String 1' should be 0")
        XCTAssertEqual(index2, 1, "Index of 'Shared String 2' should be 1")
        XCTAssertEqual(duplicateIndex, index1, "Duplicate 'Shared String 1' should return the same index as the original")
    }
}
