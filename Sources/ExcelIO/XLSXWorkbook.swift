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
}
