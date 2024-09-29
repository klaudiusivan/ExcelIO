//
//  SharedStringsParser.swift
//  ExcelIO
//
//  Created by Klaudius Ivan on 9/29/24.
//


import Foundation

class SharedStringsParser: NSObject, XMLParserDelegate {
    var sharedStrings: [String] = []
    private var currentString: String = ""
    private var inTextElement = false
    
    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName qName: String?, attributes attributeDict: [String: String]) {
        if elementName == "si" {
            currentString = ""
        } else if elementName == "t" {
            inTextElement = true
        }
    }
    
    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if inTextElement {
            currentString += string
        }
    }
    
    func parser(_ parser: XMLParser, didEndElement elementName: String, namespaceURI: String?, qualifiedName qName: String?) {
        if elementName == "si" {
            sharedStrings.append(currentString)
        } else if elementName == "t" {
            inTextElement = false
        }
    }
}
