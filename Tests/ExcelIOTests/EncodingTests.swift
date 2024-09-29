//
//  EncodingTests.swift
//  ExcelIO
//
//  Created by Klaudius Ivan on 9/29/24.
//


import XCTest
import ExcelIO

final class EncodingTests: XCTestCase {
    
    func testSimpleObjectEncoding() throws {
        struct SimpleObject: Encodable {
            var name: String
            var age: Int
            var email: String
        }

        let simpleObject = SimpleObject(name: "Alice", age: 25, email: "alice@example.com")
        let encoder = XLSXEncoder()
        let workbook = try encoder.encode(simpleObject)

        let fileURL = FileManager.default.temporaryDirectory.appendingPathComponent("SimpleObject.xlsx")
        
        // Check if file exists and remove it
        if FileManager.default.fileExists(atPath: fileURL.path) {
            try FileManager.default.removeItem(at: fileURL)
        }
        
        try XLSXWriter.save(workbook: workbook, to: fileURL)

        XCTAssertTrue(FileManager.default.fileExists(atPath: fileURL.path))
    }
    
    func testNestedObjectEncoding() throws {
        struct Address: Encodable {
            var street: String
            var city: String
            var zip: String
        }

        struct Person: Encodable {
            var name: String
            var address: Address
            var hobbies: [String]
        }

        let person = Person(
            name: "Bob",
            address: Address(street: "123 Elm St", city: "Somewhere", zip: "12345"),
            hobbies: ["Reading", "Gaming", "Traveling"]
        )
        let encoder = XLSXEncoder()
        let workbook = try encoder.encode(person)

        let fileURL = FileManager.default.temporaryDirectory.appendingPathComponent("NestedObject.xlsx")
        
        // Check if file exists and remove it
        if FileManager.default.fileExists(atPath: fileURL.path) {
            try FileManager.default.removeItem(at: fileURL)
        }
        
        try XLSXWriter.save(workbook: workbook, to: fileURL)

        XCTAssertTrue(FileManager.default.fileExists(atPath: fileURL.path))
    }

}
