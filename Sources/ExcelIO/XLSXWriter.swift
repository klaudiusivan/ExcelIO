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
        print("Temporary Directory:", tempDirectoryURL.path) // Debug
        
        try FileManager.default.createDirectory(at: tempDirectoryURL, withIntermediateDirectories: true, attributes: nil)
        
        defer {
            try? FileManager.default.removeItem(at: tempDirectoryURL)
        }
        
        let xlURL = tempDirectoryURL.appendingPathComponent("xl")
        try FileManager.default.createDirectory(at: xlURL, withIntermediateDirectories: true, attributes: nil)
        print("Created xl directory at:", xlURL.path) // Debug
        
        // Create worksheets subdirectory
        let worksheetsURL = xlURL.appendingPathComponent("worksheets")
        try FileManager.default.createDirectory(at: worksheetsURL, withIntermediateDirectories: true, attributes: nil)
        print("Created worksheets directory at:", worksheetsURL.path) // Debug
        
        // Write workbook.xml
        let workbookXMLURL = xlURL.appendingPathComponent("workbook.xml")
        print("Writing workbook.xml to:", workbookXMLURL.path) // Debug
        try writeWorkbookXML(workbook, to: workbookXMLURL)
        
        // Write sharedStrings.xml if any shared strings exist
        if !workbook.sharedStrings.isEmpty {
            let sharedStringsURL = xlURL.appendingPathComponent("sharedStrings.xml")
            print("Writing sharedStrings.xml to:", sharedStringsURL.path) // Debug
            try writeSharedStrings(workbook.sharedStrings, to: sharedStringsURL)
        }
        
        // Write sheets
        for (index, sheet) in workbook.sheets.enumerated() {
            let sheetURL = worksheetsURL.appendingPathComponent("sheet\(index + 1).xml")
            print("Writing sheet\(index + 1).xml to:", sheetURL.path) // Debug
            try writeWorksheetXML(sheet, to: sheetURL)
        }
        
        // Write the necessary [Content_Types].xml
        let contentTypesURL = tempDirectoryURL.appendingPathComponent("[Content_Types].xml")
        print("Writing [Content_Types].xml to:", contentTypesURL.path) // Debug
        try writeContentTypesXML(workbook, to: contentTypesURL)
        
        // Write relationships (not covered fully here but needed for a complete file)
        let relsURL = tempDirectoryURL.appendingPathComponent("_rels")
        try FileManager.default.createDirectory(at: relsURL, withIntermediateDirectories: true, attributes: nil)
        let relationshipsURL = relsURL.appendingPathComponent(".rels")
        print("Writing .rels to:", relationshipsURL.path) // Debug
        try writeRelationshipsXML(to: relationshipsURL)
        
        // Create zip archive
        print("Creating zip archive at:", fileURL.path) // Debug
        try zipFolder(at: tempDirectoryURL, to: fileURL)
    }
    
    // Function to write workbook.xml
    private static func writeWorkbookXML(_ workbook: XLSXWorkbook, to url: URL) throws {
        do {
            var xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheets>
            """
            for (index, sheet) in workbook.sheets.enumerated() {
                xml += "<sheet name=\"\(sheet.name)\" sheetId=\"\(index + 1)\" r:id=\"rId\(index + 1)\"/>"
            }
            xml += "</sheets></workbook>"
            
            try xml.write(to: url, atomically: true, encoding: .utf8)
        } catch {
            print("Error writing workbook.xml:", error) // Debug
            throw error
        }
    }

    // Function to write sharedStrings.xml (if used)
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
    
    // Function to write worksheet XML
    private static func writeWorksheetXML(_ sheet: XLSXWorksheet, to url: URL) throws {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
        """
        for cell in sheet.cells.values {
            xml += "<c r=\"\(cell.reference)\" t=\"str\"><v>\(escapeXML(cell.value))</v></c>"
        }
        xml += "</sheetData></worksheet>"
        
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }
    
    // Function to write [Content_Types].xml
    private static func writeContentTypesXML(_ workbook: XLSXWorkbook, to url: URL) throws {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="xml" ContentType="application/xml"/>
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        """
        if !workbook.sharedStrings.isEmpty {
            xml += """
            <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
            """
        }
        for (index, _) in workbook.sheets.enumerated() {
            xml += """
            <Override PartName="/xl/worksheets/sheet\(index + 1).xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            """
        }
        xml += "</Types>"
        
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }
    
    // Function to write relationships XML
    private static func writeRelationshipsXML(to url: URL) throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>
        """
        
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }
    
    private static func zipFolder(at sourceURL: URL, to destinationURL: URL) throws {
        let fileManager = FileManager()
        try fileManager.zipItem(at: sourceURL, to: destinationURL)
    }
    
    // Utility function to escape XML special characters
    private static func escapeXML(_ string: String) -> String {
        var escapedString = string
        escapedString = escapedString.replacingOccurrences(of: "&", with: "&amp;")
        escapedString = escapedString.replacingOccurrences(of: "<", with: "&lt;")
        escapedString = escapedString.replacingOccurrences(of: ">", with: "&gt;")
        escapedString = escapedString.replacingOccurrences(of: "\"", with: "&quot;")
        escapedString = escapedString.replacingOccurrences(of: "'", with: "&apos;")
        return escapedString
    }
}
