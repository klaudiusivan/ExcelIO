//
//  Utilities.swift
//  CV Maker
//
//  Created by Klaudius Ivan on 9/29/24.
//

import Foundation

func escapeXML(_ string: String) -> String {
    var escaped = string.replacingOccurrences(of: "&", with: "&amp;")
    escaped = escaped.replacingOccurrences(of: "<", with: "&lt;")
    escaped = escaped.replacingOccurrences(of: ">", with: "&gt;")
    escaped = escaped.replacingOccurrences(of: "\"", with: "&quot;")
    escaped = escaped.replacingOccurrences(of: "'", with: "&apos;")
    return escaped
}
