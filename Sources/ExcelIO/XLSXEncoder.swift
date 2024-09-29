//
//  XLSXEncoder.swift
//  ExcelIO
//
//  Created by Klaudius Ivan on 9/29/24.
//

import Foundation

public class XLSXEncoder {
    
    public init() {}
    
    // Function to encode any encodable value into a workbook
    public func encode<T: Encodable>(_ value: T) throws -> XLSXWorkbook {
        let workbook = XLSXWorkbook()
        try encodeToWorkbook(value, into: workbook, sheetName: "Main Sheet")
        return workbook
    }
    
    // Function to encode an encodable value into a specific workbook and sheet
    private func encodeToWorkbook<T: Encodable>(_ value: T, into workbook: XLSXWorkbook, sheetName: String) throws {
        let encoder = InternalEncoder(codingPath: [], userInfo: [:])
        try value.encode(to: encoder)
        
        // Convert the container's stored properties into a sheet
        let sheet = createSheet(from: encoder.storage, sheetName: sheetName)
        workbook.addSheet(sheet)
    }
    
    // Function to create a sheet from a storage dictionary
    private func createSheet(from storage: [String: Any], sheetName: String) -> XLSXWorksheet {
        let sheet = XLSXWorksheet(name: sheetName)
        
        var row = 1
        for (index, key) in storage.keys.sorted().enumerated() {
            // Set headers in row 1
            let headerCell = "\(Character(UnicodeScalar(65 + index)!))\(row)"
            sheet.setCell(at: headerCell, value: key)
            
            // Set values in row 2
            let valueCell = "\(Character(UnicodeScalar(65 + index)!))\(row + 1)"
            if let value = storage[key] {
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
