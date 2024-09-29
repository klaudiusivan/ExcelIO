//
//  WorkbookSheetsParser.swift
//  ExcelIO
//
//  Created by Klaudius Ivan on 9/29/24.
//


import Foundation

class WorkbookSheetsParser: NSObject, XMLParserDelegate {
    var sheetURLs: [URL] = []
    private var baseDirectory: URL
    private var currentSheetName: String?
    
    init(baseDirectory: URL) {
        self.baseDirectory = baseDirectory
    }
    
    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName qName: String?, attributes attributeDict: [String: String]) {
        if elementName == "sheet" {
            if let sheetName = attributeDict["name"] {
                currentSheetName = sheetName
            }
        }
    }
    
    func parser(_ parser: XMLParser, didEndElement elementName: String, namespaceURI: String?, qualifiedName qName: String?) {
        if elementName == "sheet" {
            if let name = currentSheetName {
                let sheetPath = "worksheets/\(name).xml"
                let sheetURL = baseDirectory.appendingPathComponent(sheetPath)
                sheetURLs.append(sheetURL)
            }
            currentSheetName = nil
        }
    }
}
