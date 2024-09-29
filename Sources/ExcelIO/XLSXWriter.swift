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
        
        try createWorkbookXML(workbook: workbook, at: tempDirectoryURL)
        try createSharedStringsXML(workbook: workbook, at: tempDirectoryURL)
        try createWorksheetXMLs(workbook: workbook, at: tempDirectoryURL)
        try createContentTypesXML(at: tempDirectoryURL)
        try createRelationships(workbook: workbook, at: tempDirectoryURL)
        
        try zipFolder(at: tempDirectoryURL, to: fileURL)
    }
    
    private static func createWorkbookXML(workbook: XLSXWorkbook, at url: URL) throws {
        let xlURL = url.appendingPathComponent("xl")
        try FileManager.default.createDirectory(at: xlURL, withIntermediateDirectories: true, attributes: nil)
        
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
        xml += "<sheets>\n"
        for (index, sheet) in workbook.sheets.enumerated() {
            xml += "<sheet name=\"\(sheet.name)\" sheetId=\"\(index + 1)\" r:id=\"rId\(index + 1)\"/>\n"
        }
        xml += "</sheets>\n"
        xml += "</workbook>"
        
        let fileURL = xlURL.appendingPathComponent("workbook.xml")
        try xml.write(to: fileURL, atomically: true, encoding: .utf8)
    }
    
    private static func createSharedStringsXML(workbook: XLSXWorkbook, at url: URL) throws {
        let xlURL = url.appendingPathComponent("xl")
        let sharedStringsURL = xlURL.appendingPathComponent("sharedStrings.xml")
        
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"\(workbook.sharedStrings.count)\" uniqueCount=\"\(workbook.sharedStrings.count)\">\n"
        for string in workbook.sharedStrings {
            xml += "<si><t>\(escapeXML(string))</t></si>\n"
        }
        xml += "</sst>"
        
        try xml.write(to: sharedStringsURL, atomically: true, encoding: .utf8)
    }
    
    private static func createWorksheetXMLs(workbook: XLSXWorkbook, at url: URL) throws {
        let xlURL = url.appendingPathComponent("xl")
        let worksheetsURL = xlURL.appendingPathComponent("worksheets")
        try FileManager.default.createDirectory(at: worksheetsURL, withIntermediateDirectories: true, attributes: nil)
        
        for (index, sheet) in workbook.sheets.enumerated() {
            var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            xml += "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
            xml += "<sheetData>\n"
            let rows = groupCellsByRow(sheet.cells)
            for (rowIndex, cells) in rows {
                xml += "<row r=\"\(rowIndex)\">\n"
                for cell in cells {
                    let sstIndex = getSharedStringIndex(for: cell.value, in: &workbook.sharedStrings)
                    xml += "<c r=\"\(cell.reference)\" t=\"s\"><v>\(sstIndex)</v></c>\n"
                }
                xml += "</row>\n"
            }
            xml += "</sheetData>\n"
            xml += "</worksheet>"
            
            let fileURL = worksheetsURL.appendingPathComponent("sheet\(index + 1).xml")
            try xml.write(to: fileURL, atomically: true, encoding: .utf8)
        }
    }
    
    private static func createContentTypesXML(at url: URL) throws {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
        xml += "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
        xml += "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
        xml += "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n"
        xml += "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\n"
        for index in 1...3 {
            xml += "<Override PartName=\"/xl/worksheets/sheet\(index).xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n"
        }
        xml += "</Types>"
        
        let fileURL = url.appendingPathComponent("[Content_Types].xml")
        try xml.write(to: fileURL, atomically: true, encoding: .utf8)
    }
    
    private static func createRelationships(workbook: XLSXWorkbook, at url: URL) throws {
        // _rels/.rels
        let relsURL = url.appendingPathComponent("_rels")
        try FileManager.default.createDirectory(at: relsURL, withIntermediateDirectories: true, attributes: nil)
        let relsXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """
        let relsFileURL = relsURL.appendingPathComponent(".rels")
        try relsXML.write(to: relsFileURL, atomically: true, encoding: .utf8)
        
        // xl/_rels/workbook.xml.rels
        let xlURL = url.appendingPathComponent("xl")
        let xlRelsURL = xlURL.appendingPathComponent("_rels")
        try FileManager.default.createDirectory(at: xlRelsURL, withIntermediateDirectories: true, attributes: nil)
        var workbookRelsXML = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        workbookRelsXML += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
        workbookRelsXML += "<Relationship Id=\"rIdSharedStrings\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>\n"
        for (index, _) in workbook.sheets.enumerated() {
            workbookRelsXML += "<Relationship Id=\"rId\(index + 1)\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet\(index + 1).xml\"/>\n"
        }
        workbookRelsXML += "</Relationships>"
        let workbookRelsFileURL = xlRelsURL.appendingPathComponent("workbook.xml.rels")
        try workbookRelsXML.write(to: workbookRelsFileURL, atomically: true, encoding: .utf8)
    }
    
    private static func zipFolder(at sourceURL: URL, to destinationURL: URL) throws {
        let fileManager = FileManager()
        try fileManager.zipItem(at: sourceURL, to: destinationURL)
    }
    
    private static func groupCellsByRow(_ cells: [String: XLSXCell]) -> [Int: [XLSXCell]] {
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
    
    private static func getRowIndex(from reference: String) -> Int? {
        let digits = reference.replacingOccurrences(of: "[^0-9]", with: "", options: .regularExpression)
        return Int(digits)
    }
    
    private static func getSharedStringIndex(for value: String, in sharedStrings: inout [String]) -> Int {
        if let index = sharedStrings.firstIndex(of: value) {
            return index
        } else {
            sharedStrings.append(value)
            return sharedStrings.count - 1
        }
    }
    
    private static func escapeXML(_ string: String) -> String {
        var escaped = string.replacingOccurrences(of: "&", with: "&amp;")
        escaped = escaped.replacingOccurrences(of: "<", with: "&lt;")
        escaped = escaped.replacingOccurrences(of: ">", with: "&gt;")
        escaped = escaped.replacingOccurrences(of: "\"", with: "&quot;")
        escaped = escaped.replacingOccurrences(of: "'", with: "&apos;")
        return escaped
    }
}
