//
//  XLSXWorksheet.swift
//  CV Maker
//
//  Created by Klaudius Ivan on 9/29/24.
//

import Foundation

public class XLSXWorksheet {
    public var name: String
    public var cells: [String: XLSXCell] = [:]
    
    public init(name: String) {
        self.name = name
    }
    
    public func setCell(at reference: String, value: String) {
        let cell = XLSXCell(reference: reference, value: value)
        cells[reference] = cell
    }
    
    public func getCell(at reference: String) -> XLSXCell? {
        return cells[reference]
    }
}

