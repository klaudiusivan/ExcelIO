//
//  XLSXEncoder.swift
//  ExcelIO
//
//  Created by Klaudius Ivan on 9/29/24.
//

import Foundation

public class XLSXEncoder {
    
    public init() {}
    
    public func encode<T: Encodable>(_ value: T) throws -> XLSXWorkbook {
        let workbook = XLSXWorkbook()
        try encodeToWorkbook(value, into: workbook, sheetName: "Main Sheet")
        return workbook
    }
    
    // Function to encode an encodable value into a specific workbook and sheet
    private func encodeToWorkbook<T: Encodable>(_ value: T, into workbook: XLSXWorkbook, sheetName: String) throws {
        let encoder = InternalEncoder(codingPath: [], userInfo: [:])
        try value.encode(to: encoder)
        
        var mainSheetData: [String: Any] = [:]
        
        // Separate nested arrays/objects into their own sheets
        for (key, nestedValue) in encoder.storage {
            if let nestedArray = nestedValue as? [Encodable] {
                let nestedSheet = try createSheet(fromArray: nestedArray, sheetName: key)
                workbook.addSheet(nestedSheet)
            } else if let nestedObject = nestedValue as? Encodable {
                let nestedSheet = try createSheet(fromObject: nestedObject, sheetName: key)
                workbook.addSheet(nestedSheet)
            } else {
                // Simple values go to the main sheet
                mainSheetData[key] = nestedValue
            }
        }
        
        // Create the main sheet with simple properties
        let mainSheet = createMainSheet(from: mainSheetData, sheetName: sheetName)
        workbook.addSheet(mainSheet)
    }
    
    // Function to create the main sheet from storage
    private func createMainSheet(from storage: [String: Any], sheetName: String) -> XLSXWorksheet {
        let sheet = XLSXWorksheet(name: sheetName)
        
        // Set headers and data in rows 1 and 2
        var columnIndex = 0
        for (key, value) in storage {
            let headerCell = "\(Character(UnicodeScalar(65 + columnIndex)!))1"
            sheet.setCell(at: headerCell, value: key)
            
            let valueCell = "\(Character(UnicodeScalar(65 + columnIndex)!))2"
            sheet.setCell(at: valueCell, value: "\(value)")
            
            columnIndex += 1
        }
        
        return sheet
    }
    
    // Function to create a sheet for an array of objects
    private func createSheet(fromArray array: [Encodable], sheetName: String) throws -> XLSXWorksheet {
        let sheet = XLSXWorksheet(name: sheetName)
        
        // Assume all objects in the array are of the same type
        guard let firstObject = array.first else { return sheet }
        
        // Convert first object to a dictionary to extract headers
        let encoder = InternalEncoder(codingPath: [], userInfo: [:])
        try firstObject.encode(to: encoder)
        let headers = encoder.storage.keys.sorted()
        
        // Set headers in row 1
        for (index, header) in headers.enumerated() {
            let headerCell = "\(Character(UnicodeScalar(65 + index)!))1"
            sheet.setCell(at: headerCell, value: header)
        }
        
        // Set values in subsequent rows
        for (rowIndex, object) in array.enumerated() {
            let row = rowIndex + 2 // Start at row 2
            
            let encoder = InternalEncoder(codingPath: [], userInfo: [:])
            try object.encode(to: encoder)
            
            for (colIndex, header) in headers.enumerated() {
                let valueCell = "\(Character(UnicodeScalar(65 + colIndex)!))\(row)"
                if let value = encoder.storage[header] {
                    sheet.setCell(at: valueCell, value: "\(value)")
                }
            }
        }
        
        return sheet
    }
    
    // Function to create a sheet for a nested object
    private func createSheet(fromObject object: Encodable, sheetName: String) throws -> XLSXWorksheet {
        let sheet = XLSXWorksheet(name: sheetName)
        
        let encoder = InternalEncoder(codingPath: [], userInfo: [:])
        try object.encode(to: encoder)
        let headers = encoder.storage.keys.sorted()
        
        // Set headers in row 1
        for (index, header) in headers.enumerated() {
            let headerCell = "\(Character(UnicodeScalar(65 + index)!))1"
            sheet.setCell(at: headerCell, value: header)
        }
        
        // Set values in row 2
        for (colIndex, header) in headers.enumerated() {
            let valueCell = "\(Character(UnicodeScalar(65 + colIndex)!))2"
            if let value = encoder.storage[header] {
                sheet.setCell(at: valueCell, value: "\(value)")
            }
        }
        
        return sheet
    }
}


// Internal Encoder
fileprivate class InternalEncoder: Encoder {
    var codingPath: [CodingKey]
    var userInfo: [CodingUserInfoKey: Any]
    var storage: [String: Any] = [:]
    
    init(codingPath: [CodingKey], userInfo: [CodingUserInfoKey: Any]) {
        self.codingPath = codingPath
        self.userInfo = userInfo
    }
    
    func container<Key>(keyedBy type: Key.Type) -> KeyedEncodingContainer<Key> {
        let container = KeyedContainer<Key>(encoder: self)
        return KeyedEncodingContainer(container)
    }
    
    func unkeyedContainer() -> UnkeyedEncodingContainer {
        return UnkeyedContainer(encoder: self)
    }
    
    func singleValueContainer() -> SingleValueEncodingContainer {
        return SingleValueContainer(encoder: self)
    }
}

// Keyed Encoding Container
fileprivate struct KeyedContainer<Key: CodingKey>: KeyedEncodingContainerProtocol {
    var encoder: InternalEncoder
    var codingPath: [CodingKey] { encoder.codingPath }
    
    mutating func encodeNil(forKey key: Key) throws {
        encoder.storage[key.stringValue] = nil
    }
    
    mutating func encode<T: Encodable>(_ value: T, forKey key: Key) throws {
        encoder.storage[key.stringValue] = value
    }
    
    mutating func nestedContainer<NestedKey>(keyedBy type: NestedKey.Type, forKey key: Key) -> KeyedEncodingContainer<NestedKey> {
        fatalError("Nested containers are not supported yet.")
    }
    
    mutating func nestedUnkeyedContainer(forKey key: Key) -> UnkeyedEncodingContainer {
        fatalError("Unkeyed nested containers are not supported yet.")
    }
    
    mutating func superEncoder() -> Encoder {
        return encoder
    }
    
    mutating func superEncoder(forKey key: Key) -> Encoder {
        return encoder
    }
}

// Unkeyed Encoding Container
fileprivate struct UnkeyedContainer: UnkeyedEncodingContainer {
    var encoder: InternalEncoder
    var codingPath: [CodingKey] { encoder.codingPath }
    var count: Int = 0
    
    mutating func encodeNil() throws {
        encoder.storage["\(count)"] = nil
        count += 1
    }
    
    mutating func encode<T: Encodable>(_ value: T) throws {
        encoder.storage["\(count)"] = value
        count += 1
    }
    
    mutating func nestedContainer<NestedKey>(keyedBy type: NestedKey.Type) -> KeyedEncodingContainer<NestedKey> {
        fatalError("Unkeyed nested containers are not supported yet.")
    }
    
    mutating func nestedUnkeyedContainer() -> UnkeyedEncodingContainer {
        fatalError("Unkeyed nested containers are not supported yet.")
    }
    
    mutating func superEncoder() -> Encoder {
        return encoder
    }
}

// Single Value Encoding Container
fileprivate struct SingleValueContainer: SingleValueEncodingContainer {
    var encoder: InternalEncoder
    var codingPath: [CodingKey] { encoder.codingPath }
    
    mutating func encodeNil() throws {
        encoder.storage[codingPath.last?.stringValue ?? ""] = nil
    }
    
    mutating func encode<T: Encodable>(_ value: T) throws {
        encoder.storage[codingPath.last?.stringValue ?? ""] = value
    }
}
