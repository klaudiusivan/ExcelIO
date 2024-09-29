//
//  XLSXReader.swift
//  CV Maker
//
//  Created by Klaudius Ivan on 9/29/24.
//


import Foundation
internal import ZIPFoundation

public class XLSXReader {
    public static func read(from fileURL: URL) throws -> XLSXWorkbook {
        let tempDirectoryURL = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.unzipItem(at: fileURL, to: tempDirectoryURL)
        
        defer {
            try? FileManager.default.removeItem(at: tempDirectoryURL)
        }
        
        let workbook = XLSXWorkbook()
        let xlURL = tempDirectoryURL.appendingPathComponent("xl")
        
        let sharedStringsURL = xlURL.appendingPathComponent("sharedStrings.xml")
        if FileManager.default.fileExists(atPath: sharedStringsURL.path) {
            let sharedStringsData = try Data(contentsOf: sharedStringsURL)
            workbook.sharedStrings = try parseSharedStrings(data: sharedStringsData)
        }
        
        let workbookXMLURL = xlURL.appendingPathComponent("workbook.xml")
        let sheetNames = try parseWorkbook(data: try Data(contentsOf: workbookXMLURL))
        
        let worksheetsURL = xlURL.appendingPathComponent("worksheets")
        for (index, sheetName) in sheetNames.enumerated() {
            let sheetURL = worksheetsURL.appendingPathComponent("sheet\(index + 1).xml")
            if FileManager.default.fileExists(atPath: sheetURL.path) {
                let sheetData = try Data(contentsOf: sheetURL)
                let sheet = XLSXWorksheet(name: sheetName)
                try parseWorksheet(data: sheetData, into: sheet, sharedStrings: workbook.sharedStrings)
                workbook.addSheet(sheet)
            }
        }
        
        return workbook
    }
    
    private static func parseSharedStrings(data: Data) throws -> [String] {
        var sharedStrings: [String] = []
        let parser = XMLParser(data: data)
        let delegate = SharedStringsParserDelegate()
        parser.delegate = delegate
        parser.parse()
        sharedStrings = delegate.sharedStrings
        return sharedStrings
    }
    
    private static func parseWorkbook(data: Data) throws -> [String] {
        var sheetNames: [String] = []
        let parser = XMLParser(data: data)
        let delegate = WorkbookParserDelegate()
        parser.delegate = delegate
        parser.parse()
        sheetNames = delegate.sheetNames
        return sheetNames
    }
    
    private static func parseWorksheet(data: Data, into sheet: XLSXWorksheet, sharedStrings: [String]) throws {
        let parser = XMLParser(data: data)
        let delegate = WorksheetParserDelegate(sheet: sheet, sharedStrings: sharedStrings)
        parser.delegate = delegate
        parser.parse()
    }
}

// XML Parser Delegates
class SharedStringsParserDelegate: NSObject, XMLParserDelegate {
    var sharedStrings: [String] = []
    private var currentString: String = ""
    private var isTElement = false
    
    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName qName: String?, attributes attributeDict: [String : String] = [:]) {
        if elementName == "t" {
            isTElement = true
            currentString = ""
        }
    }
    
    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if isTElement {
            currentString += string
        }
    }
    
    func parser(_ parser: XMLParser, didEndElement elementName: String, namespaceURI: String?, qualifiedName qName: String?) {
        if elementName == "t" {
            isTElement = false
            sharedStrings.append(currentString)
        }
    }
}

class WorkbookParserDelegate: NSObject, XMLParserDelegate {
    var sheetNames: [String] = []
    
    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName qName: String?, attributes attributeDict: [String : String] = [:]) {
        if elementName == "sheet", let name = attributeDict["name"] {
            sheetNames.append(name)
        }
    }
}

class WorksheetParserDelegate: NSObject, XMLParserDelegate {
    let sheet: XLSXWorksheet
    let sharedStrings: [String]
    private var currentCellReference: String = ""
    private var currentValue: String = ""
    private var isVElement = false
    private var isSSType = false
    
    init(sheet: XLSXWorksheet, sharedStrings: [String]) {
        self.sheet = sheet
        self.sharedStrings = sharedStrings
    }
    
    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName qName: String?, attributes attributeDict: [String : String] = [:]) {
        if elementName == "c" {
            currentCellReference = attributeDict["r"] ?? ""
            isSSType = attributeDict["t"] == "s"
        } else if elementName == "v" {
            isVElement = true
            currentValue = ""
        }
    }
    
    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if isVElement {
            currentValue += string
        }
    }
    
    func parser(_ parser: XMLParser, didEndElement elementName: String, namespaceURI: String?, qualifiedName qName: String?) {
        if elementName == "v" {
            isVElement = false
        } else if elementName == "c" {
            let value: String
            if isSSType, let index = Int(currentValue) {
                value = sharedStrings[index]
            } else {
                value = currentValue
            }
            sheet.setCell(at: currentCellReference, value: value)
        }
    }
}
