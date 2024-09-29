//
//  XLSXWriter.swift
//  CV Maker
//
//  Created by Klaudius Ivan on 9/29/24.
//

import Foundation
internal import ZIPFoundation

public class XLSXWriter {
    public static func save(workbook: XLSXWorkbook, to fileURL: URL) throws {
        let tempDirectoryURL = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: tempDirectoryURL, withIntermediateDirectories: true, attributes: nil)
        
        defer {
            try? FileManager.default.removeItem(at: tempDirectoryURL)
        }
        
        let xlURL = tempDirectoryURL.appendingPathComponent("xl")
        try FileManager.default.createDirectory(at: xlURL, withIntermediateDirectories: true, attributes: nil)
        
        // Write sharedStrings.xml
        let sharedStringsURL = xlURL.appendingPathComponent("sharedStrings.xml")
        try writeSharedStrings(workbook.sharedStrings, to: sharedStringsURL)
        
        // Write workbook.xml
        let workbookXMLURL = xlURL.appendingPathComponent("workbook.xml")
        try writeWorkbookXML(workbook, to: workbookXMLURL)
        
        // Write sheets
        for (index, sheet) in workbook.sheets.enumerated() {
            let sheetURL = xlURL.appendingPathComponent("worksheets/sheet\(index + 1).xml")
            try writeWorksheetXML(sheet, to: sheetURL)
        }
        
        // Create zip archive
        try zipFolder(at: tempDirectoryURL, to: fileURL)
    }
    
    private static func writeSharedStrings(_ sharedStrings: [String], to url: URL) throws {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        """
        for string in sharedStrings {
            xml += "<si><t>\(escapeXML(string))</t></si>"
        }
        xml += "</sst>"
        
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }
    
    private static func writeWorkbookXML(_ workbook: XLSXWorkbook, to url: URL) throws {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheets>
        """
        for (index, sheet) in workbook.sheets.enumerated() {
            xml += "<sheet name=\"\(sheet.name)\" sheetId=\"\(index + 1)\"/>"
        }
        xml += "</sheets></workbook>"
        
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }
    
    private static func writeWorksheetXML(_ sheet: XLSXWorksheet, to url: URL) throws {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
        """
        for cell in sheet.cells.values {
            xml += "<c r=\"\(cell.reference)\"><v>\(escapeXML(cell.value))</v></c>"
        }
        xml += "</sheetData></worksheet>"
        
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }
    
    private static func zipFolder(at sourceURL: URL, to destinationURL: URL) throws {
        let fileManager = FileManager()
        try fileManager.zipItem(at: sourceURL, to: destinationURL)
    }
}
