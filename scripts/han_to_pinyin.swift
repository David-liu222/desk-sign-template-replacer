import Foundation

while let raw = readLine() {
    let value = NSMutableString(string: raw)
    CFStringTransform(value, nil, kCFStringTransformToLatin, false)
    CFStringTransform(value, nil, kCFStringTransformStripCombiningMarks, false)
    print(value)
}
