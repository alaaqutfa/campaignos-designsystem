// ExtendScript – Export ALL visible unlocked items (per-layer reversed index)
// #target illustrator

// CONFIGURATION
var pattern = "img_";
var useGeometricBounds = false;   // false = visibleBounds, true = geometricBounds
var logFile = new File("~/Desktop/debug_log.txt");
logFile.open("w");

function log(msg) {
    logFile.writeln(msg);
}

// Helper: pad number with leading zeros
function zeroPad(num, len) {
    var s = String(num);
    while (s.length < len) s = "0" + s;
    return s;
}

// Helper: safe get property
function safeGet(obj, prop) {
    try {
        return obj[prop];
    } catch (e) {
        return null;
    }
}

// Helper: escape CSV field
function escapeCSVField(field) {
    if (typeof field === "string" && (field.indexOf(",") > -1 || field.indexOf('"') > -1)) {
        return '"' + field.replace(/"/g, '""') + '"';
    }
    return field;
}

// ================== Name parsing ==================
var KNOWN_TYPES = ["banner", "bannerLogo", "bannerShopName"];

function parseItemName(name) {
    if (!name || name.indexOf(pattern) !== 0) {
        return { imageName: name || "", type: "image" };
    }
    var rest = name.substring(pattern.length); // after "img_"
    var parts = rest.split("_");
    if (parts.length === 0) {
        return { imageName: rest, type: "image" };
    }
    var lastPart = parts[parts.length - 1];
    // Check if last part is a known type
    if (KNOWN_TYPES && typeof KNOWN_TYPES.indexOf === "function" && KNOWN_TYPES.indexOf(lastPart) !== -1) {
        var type = lastPart;
        var imageName = parts.slice(0, -1).join("_");
        return { imageName: imageName, type: type };
    } else {
        return { imageName: rest, type: "image" };
    }
}

// ================== Color helpers ==================
function colorToString(color) {
    if (!color) return "";
    if (color.typename === "RGBColor") {
        var r = Math.round(color.red * 255);
        var g = Math.round(color.green * 255);
        var b = Math.round(color.blue * 255);
        return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
    } else if (color.typename === "CMYKColor") {
        return "CMYK(" + color.cyan + "," + color.magenta + "," + color.yellow + "," + color.black + ")";
    } else if (color.typename === "GrayColor") {
        var gray = Math.round(color.gray * 100);
        return "Gray(" + gray + "%)";
    } else {
        return color.typename;
    }
}

// Get fill color from any page item (recursively for groups)
function getItemFillColor(item) {
    if (!item) return "";
    try {
        // Direct fillColor
        if (item.fillColor) {
            return colorToString(item.fillColor);
        }
        // If it's a group, look into its children
        if (item.typename === "GroupItem" && item.pageItems && item.pageItems.length > 0) {
            for (var i = 0; i < item.pageItems.length; i++) {
                var col = getItemFillColor(item.pageItems[i]);
                if (col) return col;
            }
        }
        // If it's a compound path, it may have fillColor at path level
        if (item.typename === "CompoundPathItem" && item.pathItems && item.pathItems.length > 0) {
            for (var i = 0; i < item.pathItems.length; i++) {
                var col = getItemFillColor(item.pathItems[i]);
                if (col) return col;
            }
        }
    } catch (e) {
        log("Error in getItemFillColor: " + e.message);
    }
    return "";
}

// Get text color from a TextFrame
function getTextColor(textFrame) {
    if (!textFrame || textFrame.typename !== "TextFrame") return "";
    try {
        var charAttr = safeGet(textFrame.textRange, "characterAttributes");
        if (charAttr && charAttr.color) {
            return colorToString(charAttr.color);
        }
        if (textFrame.fillColor) {
            return colorToString(textFrame.fillColor);
        }
    } catch (e) {
        log("Error in getTextColor: " + e.message);
    }
    return "";
}

// ================== Collection per layer ==================
function collectItems() {
    var doc = app.activeDocument;
    var layerItems = {};
    var layerOrder = [];

    for (var i = 0; i < doc.layers.length; i++) {
        var layer = doc.layers[i];
        if (layer.hidden || layer.locked) continue;
        var itemsInLayer = [];
        collectItemsRecursive(layer, itemsInLayer, layer.name);
        if (itemsInLayer.length > 0) {
            layerItems[layer.name] = itemsInLayer;
            layerOrder.push(layer.name);
        }
    }

    return { layerItems: layerItems, layerOrder: layerOrder };
}

function collectItemsRecursive(parent, itemsArray, layerName) {
    if (!parent || parent.hidden || parent.locked) return;

    // GroupItem: add if name matches, but do not descend
    if (parent.typename === "GroupItem") {
        if (parent.name && parent.name.indexOf(pattern) !== -1) {
            itemsArray.push({ item: parent, layerName: layerName });
        }
        return;
    }

    // Layer: iterate its children
    if (parent.typename === "Layer") {
        for (var i = 0; i < parent.pageItems.length; i++) {
            collectItemsRecursive(parent.pageItems[i], itemsArray, parent.name);
        }
        return;
    }

    // All other items: add if name matches
    if (parent.name && parent.name.indexOf(pattern) !== -1) {
        itemsArray.push({ item: parent, layerName: layerName });
    }
}

// ================== Main ==================
function main() {
    if (!app.documents.length) {
        alert("لا يوجد مستند مفتوح.");
        return;
    }

    var doc = app.activeDocument;
    var activeArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
    var rect = activeArtboard.artboardRect;
    var artboardLeft = rect[0];
    var artboardTop = rect[1];
    var artboardRight = rect[2];
    var artboardBottom = rect[3];
    var artboardWidth = artboardRight - artboardLeft;
    var artboardHeight = artboardTop - artboardBottom;

    var collected = collectItems();
    var layerItems = collected.layerItems;
    var layerOrder = collected.layerOrder;

    if (layerOrder.length === 0) {
        alert("لم يتم العثور على أي عناصر مرئية وغير مقفلة.");
        return;
    }

    // CSV headers
    var headers = [
        "layout", "layout_type", "element_type", "color",
        "image", "index", "anchor", "ArtBoard_Width", "ArtBoard_Height",
        "Img_Width", "Img_Height", "Img_X", "Img_y", "ratio", "width_pct",
        "height_pct", "x_offset_pct", "y_offset_pct", "font_size_pct"
    ];
    var csvRows = [headers.join(",")];

    var errors = [];
    var totalItems = 0;

    for (var l = 0; l < layerOrder.length; l++) {
        var layerName = layerOrder[l];
        var items = layerItems[layerName];
        // Reverse order: last item becomes index 1
        for (var idx = items.length - 1; idx >= 0; idx--) {
            var itemObj = items[idx];
            var item = itemObj.item;
            var indexInLayer = items.length - idx;  // 1-based reversed index

            log("Processing: " + item.name + " (type: " + item.typename + ", layer: " + layerName + ")");

            try {
                // Parse name
                var parsed = parseItemName(item.name);
                var imageName = parsed.imageName;
                var elementTypeFromName = parsed.type;
                log("  parsed: imageName=" + imageName + ", type=" + elementTypeFromName);

                // Get bounds
                var bounds;
                if (useGeometricBounds) {
                    bounds = safeGet(item, "geometricBounds");
                } else {
                    bounds = safeGet(item, "visibleBounds");
                }
                if (!bounds || bounds.length < 4) {
                    throw new Error("Bounds not available");
                }

                var left = bounds[0];
                var top = bounds[1];
                var right = bounds[2];
                var bottom = bounds[3];
                var imgWidth = Math.abs(right - left);
                var imgHeight = Math.abs(top - bottom);
                var imgX = left - artboardLeft;
                var imgY = artboardTop - top;

                var ratio = (artboardHeight !== 0) ? artboardWidth / artboardHeight : 0;
                var widthPct = (imgWidth / artboardWidth) * 100;
                var heightPct = (imgHeight / artboardHeight) * 100;
                var xOffsetPct = (imgX / artboardWidth) * 100;
                var yOffsetPct = (imgY / artboardHeight) * 100;

                var anchor = "top-left";

                // Color and font size
                var colorValue = "";
                var fontSizePct = "";

                if (elementTypeFromName === "banner") {
                    colorValue = getItemFillColor(item);
                    log("  banner color: " + colorValue);
                } else if (elementTypeFromName === "bannerShopName") {
                    if (item.typename === "TextFrame") {
                        colorValue = getTextColor(item);
                        try {
                            var charAttr = safeGet(item.textRange, "characterAttributes");
                            if (charAttr && charAttr.size) {
                                fontSizePct = (charAttr.size / artboardHeight) * 100;
                            }
                        } catch (e) { /* ignore */ }
                    }
                    log("  bannerShopName color: " + colorValue + ", fontSizePct: " + fontSizePct);
                } else {
                    // "image" or "bannerLogo" – no color extraction
                    if (item.typename === "TextFrame") {
                        try {
                            var charAttr = safeGet(item.textRange, "characterAttributes");
                            if (charAttr && charAttr.size) {
                                fontSizePct = (charAttr.size / artboardHeight) * 100;
                            }
                        } catch (e) { /* ignore */ }
                    }
                }

                // If text frame and font size not set yet
                if (fontSizePct === "" && item.typename === "TextFrame") {
                    try {
                        var charAttr = safeGet(item.textRange, "characterAttributes");
                        if (charAttr && charAttr.size) {
                            fontSizePct = (charAttr.size / artboardHeight) * 100;
                        }
                    } catch (e) { /* ignore */ }
                }

                // Round numeric values to 2 decimals
                function round2(val) {
                    if (typeof val !== "number") return val;
                    return Math.round(val * 100) / 100;
                }

                var rowValues = [
                    activeArtboard.name,
                    layerName,
                    elementTypeFromName,
                    colorValue,
                    imageName,
                    indexInLayer,                     // integer
                    anchor,
                    round2(artboardWidth),
                    round2(artboardHeight),
                    round2(imgWidth),
                    round2(imgHeight),
                    round2(imgX),
                    round2(imgY),
                    round2(ratio),
                    round2(widthPct),
                    round2(heightPct),
                    round2(xOffsetPct),
                    round2(yOffsetPct),
                    round2(fontSizePct)
                ];

                // Build CSV row
                var escapedRow = "";
                for (var f = 0; f < rowValues.length; f++) {
                    var field = rowValues[f];
                    if (typeof field === "string" && (field.indexOf(",") > -1 || field.indexOf('"') > -1)) {
                        field = '"' + field.replace(/"/g, '""') + '"';
                    }
                    if (f > 0) escapedRow += ",";
                    escapedRow += field;
                }
                csvRows.push(escapedRow);
                totalItems++;

            } catch (e) {
                errors.push({
                    index: indexInLayer,
                    name: item.name || "غير مسموى",
                    type: safeGet(item, "typename") || "Unknown",
                    error: e.message,
                    layer: layerName
                });
                log("ERROR: " + e.message);
            }
        }
    }

    // Save CSV
    var desktop = Folder.desktop;
    var csvFile = new File(desktop + "/artboard_export.csv");
    csvFile.open("w");
    csvFile.encoding = "UTF-8";
    csvFile.write(csvRows.join("\n"));
    csvFile.close();

    // Save errors if any
    if (errors.length > 0) {
        var errorFile = new File(desktop + "/artboard_export_errors.txt");
        errorFile.open("w");
        errorFile.encoding = "UTF-8";
        errorFile.write("Errors encountered during export:\n\n");
        for (var e = 0; e < errors.length; e++) {
            errorFile.write("Layer: " + errors[e].layer + " - Item #" + errors[e].index +
                " - Name: " + errors[e].name + " - Type: " + errors[e].type +
                " - Error: " + errors[e].error + "\n");
        }
        errorFile.close();
    }

    // Final alert
    alert("تم التصدير إلى:\n" + csvFile.fsName +
        "\nعدد العناصر الإجمالي: " + totalItems +
        (errors.length ? "\nأخطاء: " + errors.length + " (انظر ملف الأخطاء)" : ""));

    logFile.close();
}

// Run
main();