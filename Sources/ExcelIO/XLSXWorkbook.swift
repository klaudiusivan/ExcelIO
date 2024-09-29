//
//  XLSXWorkbook.swift
//  CV Maker
//
//  Created by Klaudius Ivan on 9/29/24.
//

import Foundation

public class XLSXWorkbook {
    public var sheets: [XLSXWorksheet] = []
    public var sharedStrings: [String] = []

    public init() {}

    public func addSheet(_ sheet: XLSXWorksheet) {
        sheets.append(sheet)
    }
    
    public func toJSON() -> [String: Any] {
        var workbookDict: [String: Any] = [:]
        for sheet in sheets {
            workbookDict[sheet.name] = sheet.toJSON()
        }
        return workbookDict
    }
    
    /// Add a shared string if it's not already present
    public func addSharedString(_ string: String) -> Int {
        // Check if the string already exists
        if let index = sharedStrings.firstIndex(of: string) {
            return index
        } else {
            // If not, add it and return its index
            sharedStrings.append(string)
            return sharedStrings.count - 1
        }
    }
    
    /// Retrieve the index of a shared string
    public func sharedStringIndex(for string: String) -> Int? {
        return sharedStrings.firstIndex(of: string)
    }
}
