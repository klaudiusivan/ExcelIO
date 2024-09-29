// swift-tools-version:5.7

import PackageDescription

let package = Package(
    name: "ExcelIO",
    platforms: [
        .macOS(.v10_15), .iOS(.v13)
    ],
    products: [
        .library(
            name: "ExcelIO",
            targets: ["ExcelIO"]),
    ],
    dependencies: [
        .package(url: "https://github.com/weichsel/ZIPFoundation.git", exact: "0.9.19")
    ],
    targets: [
        .target(
            name: "ExcelIO",
            dependencies: ["ZIPFoundation"]
        ),
        .testTarget(
            name: "ExcelIOTests",
            dependencies: ["ExcelIO"]),
    ]
)
