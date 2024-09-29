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
    
    public func toJSON() -> [[String: Any]] {
        let rows = groupCellsByRow(cells)
        var jsonRows: [[String: Any]] = []
        for row in rows.sorted(by: { $0.key < $1.key }) {
            var rowDict: [String: Any] = [:]
            for cell in row.value {
                let columnName = getColumnName(from: cell.reference)
                rowDict[columnName] = cell.value
            }
            jsonRows.append(rowDict)
        }
        return jsonRows
    }
    
    private func groupCellsByRow(_ cells: [String: XLSXCell]) -> [Int: [XLSXCell]] {
        var rows: [Int: [XLSXCell]] = [:]
        for cell in cells.values {
            if let rowIndex = getRowIndex(from: cell.reference) {
                if rows[rowIndex] == nil {
                    rows[rowIndex] = []
                }
                rows[rowIndex]?.append(cell)
            }
        }
        return rows
    }
    
    private func getRowIndex(from reference: String) -> Int? {
        let digits = reference.replacingOccurrences(of: "[^0-9]", with: "", options: .regularExpression)
        return Int(digits)
    }
    
    private func getColumnName(from reference: String) -> String {
        let letters = reference.replacingOccurrences(of: "[^A-Z]", with: "", options: .regularExpression)
        return letters
    }
}


