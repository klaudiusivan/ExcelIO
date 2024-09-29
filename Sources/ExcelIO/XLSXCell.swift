//
//  XLSXCell.swift
//  CV Maker
//
//  Created by Klaudius Ivan on 9/29/24.
//

import Foundation

public class XLSXCell {
    public var reference: String
    public var value: String
    public var type: String = "s" // Default to shared string
    
    public init(reference: String, value: String) {
        self.reference = reference
        self.value = value
    }
}

