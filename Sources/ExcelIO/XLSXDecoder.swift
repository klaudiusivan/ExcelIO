//
//  XLSXDecoder.swift
//  ExcelIO
//
//  Created by Klaudius Ivan on 9/29/24.
//


import Foundation

public class XLSXDecoder: Decoder {
    private let workbookDict: [String: Any]
    
    public var codingPath: [CodingKey] = []
    public var userInfo: [CodingUserInfoKey: Any] = [:]
    
    init(workbookDict: [String: Any]) {
        self.workbookDict = workbookDict
    }
    
    public func container<Key: CodingKey>(keyedBy type: Key.Type) throws -> KeyedDecodingContainer<Key> {
        let container = XLSXKeyedDecodingContainer<Key>(workbookDict: workbookDict, codingPath: codingPath)
        return KeyedDecodingContainer(container)
    }
    
    public func unkeyedContainer() throws -> UnkeyedDecodingContainer {
        fatalError("Unkeyed container decoding is not supported yet")
    }
    
    public func singleValueContainer() throws -> SingleValueDecodingContainer {
        fatalError("Single value container decoding is not supported yet")
    }
}

struct XLSXKeyedDecodingContainer<Key: CodingKey>: KeyedDecodingContainerProtocol {
    let workbookDict: [String: Any]
    var codingPath: [CodingKey] = []

    var allKeys: [Key] {
        return workbookDict.keys.compactMap { Key(stringValue: $0) }
    }

    func contains(_ key: Key) -> Bool {
        return workbookDict[key.stringValue] != nil
    }

    func decodeNil(forKey key: Key) throws -> Bool {
        return workbookDict[key.stringValue] == nil
    }

    func decode(_ type: String.Type, forKey key: Key) throws -> String {
        guard let value = workbookDict[key.stringValue] as? String else {
            throw DecodingError.typeMismatch(String.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a string"))
        }
        return value
    }

    func decode(_ type: Int.Type, forKey key: Key) throws -> Int {
        guard let value = workbookDict[key.stringValue] else {
            throw DecodingError.keyNotFound(key, DecodingError.Context(codingPath: codingPath, debugDescription: "No value associated with key \(key.stringValue)"))
        }
        if let intValue = value as? Int {
            return intValue
        } else if let stringValue = value as? String, let intValue = Int(stringValue) {
            return intValue
        }
        throw DecodingError.typeMismatch(Int.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected an integer"))
    }

    func decode(_ type: Double.Type, forKey key: Key) throws -> Double {
        guard let value = workbookDict[key.stringValue] else {
            throw DecodingError.keyNotFound(key, DecodingError.Context(codingPath: codingPath, debugDescription: "No value associated with key \(key.stringValue)"))
        }
        if let doubleValue = value as? Double {
            return doubleValue
        } else if let stringValue = value as? String, let doubleValue = Double(stringValue) {
            return doubleValue
        }
        throw DecodingError.typeMismatch(Double.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a double"))
    }
    
    func decode(_ type: Float.Type, forKey key: Key) throws -> Float {
        guard let value = workbookDict[key.stringValue] else {
            throw DecodingError.keyNotFound(key, DecodingError.Context(codingPath: codingPath, debugDescription: "No value associated with key \(key.stringValue)"))
        }
        if let floatValue = value as? Float {
            return floatValue
        } else if let stringValue = value as? String, let floatValue = Float(stringValue) {
            return floatValue
        }
        throw DecodingError.typeMismatch(Float.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a float"))
    }

    func decode(_ type: Bool.Type, forKey key: Key) throws -> Bool {
        guard let value = workbookDict[key.stringValue] else {
            throw DecodingError.keyNotFound(key, DecodingError.Context(codingPath: codingPath, debugDescription: "No value associated with key \(key.stringValue)"))
        }
        if let boolValue = value as? Bool {
            return boolValue
        } else if let stringValue = value as? String {
            switch stringValue.lowercased() {
            case "true", "yes", "1":
                return true
            case "false", "no", "0":
                return false
            default:
                break
            }
        }
        throw DecodingError.typeMismatch(Bool.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a boolean"))
    }
    
    func decode(_ type: Date.Type, forKey key: Key) throws -> Date {
        guard let value = workbookDict[key.stringValue] as? String else {
            throw DecodingError.keyNotFound(key, DecodingError.Context(codingPath: codingPath, debugDescription: "No value associated with key \(key.stringValue)"))
        }
        
        // Attempt parsing with ISO8601 format first
        let isoFormatter = ISO8601DateFormatter()
        if let date = isoFormatter.date(from: value) {
            return date
        }
        
        // Attempt parsing with a standard date formatter as a fallback
        let formatter = DateFormatter()
        formatter.dateFormat = "yyyy-MM-dd'T'HH:mm:ssZ" // ISO8601
        if let date = formatter.date(from: value) {
            return date
        }
        
        throw DecodingError.typeMismatch(Date.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a date in ISO8601 format or 'yyyy-MM-dd'T'HH:mm:ssZ'"))
    }
    
    func decode(_ type: Data.Type, forKey key: Key) throws -> Data {
        guard let value = workbookDict[key.stringValue] as? String else {
            throw DecodingError.keyNotFound(key, DecodingError.Context(codingPath: codingPath, debugDescription: "No value associated with key \(key.stringValue)"))
        }
        
        // Assuming the data is base64 encoded
        if let data = Data(base64Encoded: value) {
            return data
        }
        
        throw DecodingError.typeMismatch(Data.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a base64 encoded string"))
    }

    // Generic decoding for all Decodable types
    func decode<T: Decodable>(_ type: T.Type, forKey key: Key) throws -> T {
        guard let value = workbookDict[key.stringValue] else {
            throw DecodingError.keyNotFound(key, DecodingError.Context(codingPath: codingPath, debugDescription: "No value associated with key \(key.stringValue)"))
        }
        
        // Handle primitive Decodable types
        let decoder = XLSXSingleValueDecoder(value: value, codingPath: codingPath)
        return try T(from: decoder)
    }
    
    // For nested keyed containers
    func nestedContainer<NestedKey: CodingKey>(keyedBy type: NestedKey.Type, forKey key: Key) throws -> KeyedDecodingContainer<NestedKey> {
        guard let nestedDict = workbookDict[key.stringValue] as? [String: Any] else {
            throw DecodingError.typeMismatch([String: Any].self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a nested dictionary"))
        }
        let container = XLSXKeyedDecodingContainer<NestedKey>(workbookDict: nestedDict, codingPath: codingPath + [key])
        return KeyedDecodingContainer(container)
    }

    // For nested unkeyed containers (arrays)
    func nestedUnkeyedContainer(forKey key: Key) throws -> UnkeyedDecodingContainer {
        guard let nestedArray = workbookDict[key.stringValue] as? [Any] else {
            throw DecodingError.typeMismatch([Any].self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a nested array"))
        }
        return XLSXUnkeyedDecodingContainer(nestedArray: nestedArray, codingPath: codingPath + [key])
    }

    // For super decoding (for inheritance)
    func superDecoder() throws -> Decoder {
        return XLSXDecoder(workbookDict: workbookDict)
    }

    func superDecoder(forKey key: Key) throws -> Decoder {
        guard let nestedDict = workbookDict[key.stringValue] as? [String: Any] else {
            throw DecodingError.keyNotFound(key, DecodingError.Context(codingPath: codingPath, debugDescription: "No value associated with key \(key.stringValue)"))
        }
        return XLSXDecoder(workbookDict: nestedDict)
    }
}


struct XLSXUnkeyedDecodingContainer: UnkeyedDecodingContainer {
    let nestedArray: [Any]
    var currentIndex: Int = 0
    var codingPath: [CodingKey] = []
    
    var count: Int? {
        return nestedArray.count
    }
    
    var isAtEnd: Bool {
        return currentIndex >= (count ?? 0)
    }
    
    mutating func decodeNil() throws -> Bool {
        guard !isAtEnd else {
            throw DecodingError.valueNotFound(Any?.self, DecodingError.Context(codingPath: codingPath, debugDescription: "No more elements in array"))
        }
        if nestedArray[currentIndex] is NSNull {
            currentIndex += 1
            return true
        }
        return false
    }
    
    mutating func decode(_ type: String.Type) throws -> String {
        guard !isAtEnd else { throw DecodingError.valueNotFound(String.self, DecodingError.Context(codingPath: codingPath, debugDescription: "No more elements in array")) }
        if let stringValue = nestedArray[currentIndex] as? String {
            currentIndex += 1
            return stringValue
        } else {
            throw DecodingError.typeMismatch(String.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a string"))
        }
    }
    
    mutating func decode(_ type: Int.Type) throws -> Int {
        guard !isAtEnd else { throw DecodingError.valueNotFound(Int.self, DecodingError.Context(codingPath: codingPath, debugDescription: "No more elements in array")) }
        let value = nestedArray[currentIndex]
        currentIndex += 1
        
        if let intValue = value as? Int {
            return intValue
        } else if let stringValue = value as? String, let intValue = Int(stringValue) {
            return intValue
        }
        throw DecodingError.typeMismatch(Int.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected an integer"))
    }
    
    mutating func decode(_ type: Double.Type) throws -> Double {
        guard !isAtEnd else { throw DecodingError.valueNotFound(Double.self, DecodingError.Context(codingPath: codingPath, debugDescription: "No more elements in array")) }
        let value = nestedArray[currentIndex]
        currentIndex += 1
        
        if let doubleValue = value as? Double {
            return doubleValue
        } else if let stringValue = value as? String, let doubleValue = Double(stringValue) {
            return doubleValue
        }
        throw DecodingError.typeMismatch(Double.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a double"))
    }
    
    mutating func decode(_ type: Bool.Type) throws -> Bool {
        guard !isAtEnd else { throw DecodingError.valueNotFound(Bool.self, DecodingError.Context(codingPath: codingPath, debugDescription: "No more elements in array")) }
        let value = nestedArray[currentIndex]
        currentIndex += 1
        
        if let boolValue = value as? Bool {
            return boolValue
        } else if let stringValue = value as? String {
            if stringValue.lowercased() == "true" {
                return true
            } else if stringValue.lowercased() == "false" {
                return false
            }
        }
        throw DecodingError.typeMismatch(Bool.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a boolean"))
    }
    
    mutating func decode(_ type: Date.Type) throws -> Date {
        guard !isAtEnd else { throw DecodingError.valueNotFound(Date.self, DecodingError.Context(codingPath: codingPath, debugDescription: "No more elements in array")) }
        let value = nestedArray[currentIndex]
        currentIndex += 1
        
        if let dateString = value as? String {
            let formatter = ISO8601DateFormatter()
            if let date = formatter.date(from: dateString) {
                return date
            }
        }
        throw DecodingError.typeMismatch(Date.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a date in ISO8601 format"))
    }
    
    // Additional decoding methods for types such as Float, etc.
    
    mutating func decode<T>(_ type: T.Type) throws -> T where T: Decodable {
        guard !isAtEnd else { throw DecodingError.valueNotFound(T.self, DecodingError.Context(codingPath: codingPath, debugDescription: "No more elements in array")) }
        let value = nestedArray[currentIndex]
        currentIndex += 1
        
        // Handle Decodable types
        let decoder = XLSXSingleValueDecoder(value: value, codingPath: codingPath)
        return try T(from: decoder)
    }

    mutating func nestedContainer<NestedKey: CodingKey>(keyedBy type: NestedKey.Type) throws -> KeyedDecodingContainer<NestedKey> {
        guard !isAtEnd else { throw DecodingError.valueNotFound([String: Any].self, DecodingError.Context(codingPath: codingPath, debugDescription: "No more elements in array")) }
        guard let nestedDict = nestedArray[currentIndex] as? [String: Any] else {
            throw DecodingError.typeMismatch([String: Any].self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a nested dictionary"))
        }
        currentIndex += 1
        let container = XLSXKeyedDecodingContainer<NestedKey>(workbookDict: nestedDict, codingPath: codingPath)
        return KeyedDecodingContainer(container)
    }
    
    mutating func nestedUnkeyedContainer() throws -> UnkeyedDecodingContainer {
        guard !isAtEnd else { throw DecodingError.valueNotFound([Any].self, DecodingError.Context(codingPath: codingPath, debugDescription: "No more elements in array")) }
        guard let nestedArray = nestedArray[currentIndex] as? [Any] else {
            throw DecodingError.typeMismatch([Any].self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a nested array"))
        }
        currentIndex += 1
        return XLSXUnkeyedDecodingContainer(nestedArray: nestedArray, codingPath: codingPath)
    }
    
    mutating func superDecoder() throws -> Decoder {
        return XLSXDecoder(workbookDict: ["super": nestedArray[currentIndex]])
    }
}

struct XLSXSingleValueDecoder: Decoder {
    let value: Any
    var codingPath: [CodingKey] = []
    var userInfo: [CodingUserInfoKey: Any] = [:]

    func container<Key>(keyedBy: Key.Type) throws -> KeyedDecodingContainer<Key> {
        guard let dictionary = value as? [String: Any] else {
            throw DecodingError.typeMismatch([String: Any].self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a dictionary"))
        }
        let container = XLSXKeyedDecodingContainer<Key>(workbookDict: dictionary, codingPath: codingPath)
        return KeyedDecodingContainer(container)
    }

    func unkeyedContainer() throws -> UnkeyedDecodingContainer {
        guard let array = value as? [Any] else {
            throw DecodingError.typeMismatch([Any].self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected an array"))
        }
        return XLSXUnkeyedDecodingContainer(nestedArray: array, codingPath: codingPath)
    }

    func singleValueContainer() throws -> SingleValueDecodingContainer {
        return self
    }
}

extension XLSXSingleValueDecoder: SingleValueDecodingContainer {
    func decodeNil() -> Bool {
        return value is NSNull
    }

    func decode(_ type: Bool.Type) throws -> Bool {
        guard let boolValue = value as? Bool else {
            throw DecodingError.typeMismatch(Bool.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a Bool"))
        }
        return boolValue
    }

    func decode(_ type: String.Type) throws -> String {
        guard let stringValue = value as? String else {
            throw DecodingError.typeMismatch(String.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a String"))
        }
        return stringValue
    }

    func decode(_ type: Double.Type) throws -> Double {
        if let doubleValue = value as? Double {
            return doubleValue
        } else if let stringValue = value as? String, let doubleValue = Double(stringValue) {
            return doubleValue
        }
        throw DecodingError.typeMismatch(Double.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected a Double"))
    }

    func decode(_ type: Int.Type) throws -> Int {
        if let intValue = value as? Int {
            return intValue
        } else if let stringValue = value as? String, let intValue = Int(stringValue) {
            return intValue
        }
        throw DecodingError.typeMismatch(Int.self, DecodingError.Context(codingPath: codingPath, debugDescription: "Expected an Int"))
    }

    // Implement decode methods for other types as needed

    func decode<T>(_ type: T.Type) throws -> T where T: Decodable {
        return try T(from: self)
    }
}
