// swift-tools-version:5.0

import PackageDescription

let package = Package(
    name: "ExcelIO",
    platforms: [
        .macOS(.v10_15), .iOS(.v13) // Specify supported platforms
    ],
    products: [
        .library(
            name: "ExcelIO",
            targets: ["ExcelIO"]),
    ],
    dependencies: [
        .package(url: "https://github.com/weichsel/ZIPFoundation.git", from: "0.9.12")
    ],
    targets: [
        .target(
            name: "ExcelIO",
            dependencies: [] // Add dependencies here if needed
        ),
        .testTarget(
            name: "ExcelIOTests",
            dependencies: ["ExcelIO"]),
    ]
)
