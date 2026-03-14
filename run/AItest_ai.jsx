var art_offset = 250;
var currentDoc = null;
var currentCutMap = {};
var currentArtboard = {};
var artboardXY = [];
var end = 1;
var mqxSize = 0;
var newArtboardSize = 0;

function normalizeFilePath(path) {
    return String(path || "").replace(/\\/g, "/");
}

function normalizeWinPath(path) {
    return String(path || "").replace(/\//g, "\\");
}

function parseRunArgs(data) {
    var s = String(data || "");
    var p1 = s.indexOf(";");
    var p2 = s.indexOf(";", p1 + 1);

    if (p1 < 0) {
        return { cx: (+s || 0), filePath: "", outDir: "" };
    }
    if (p2 < 0) {
        return {
            cx: (+(s.substring(0, p1)) || 0),
            filePath: s.substring(p1 + 1),
            outDir: ""
        };
    }
    return {
        cx: (+(s.substring(0, p1)) || 0),
        filePath: s.substring(p1 + 1, p2),
        outDir: s.substring(p2 + 1)
    };
}

function parseSizeToken(token) {
    var m = String(token || "").match(/(\d+(\.\d+)?)\s*[\*xX×]\s*(\d+(\.\d+)?)/);
    if (!m) return null;
    return [ +m[1], +m[3] ];
}

function mm2pt(mm, needRound) {
    var pt = UnitValue(mm, "mm").as("pt");
    if (needRound) return Math.round(pt);
    return pt;
}

function esIndexOf(searchElement) {
    if (!searchElement) return -1;
    var name = String(searchElement).toUpperCase();
    var array = [
        "PANTONE 802 C",
        "PANTONE 802C",
        "CUT",
        "DIE",
        "DIELINE",
        "刀",
        "模切"
    ];
    for (var i = 0; i < array.length; i++) {
        if (name.indexOf(array[i]) !== -1) {
            return i;
        }
    }
    return -1;
}

function isDieLineStroke(pathItem) {
    if (!pathItem || !pathItem.stroked || !pathItem.strokeColor) return false;

    if (pathItem.strokeColor.typename == "SpotColor") {
        var spot = pathItem.strokeColor.spot;
        return !!spot && esIndexOf(spot.name) != -1;
    }

    if (pathItem.strokeColor.typename == "RGBColor") {
        var c = pathItem.strokeColor;
        return c.green >= 200 && c.red <= 80 && c.blue <= 80;
    }

    if (pathItem.strokeColor.typename == "CMYKColor") {
        var k = pathItem.strokeColor;
        var isGreenLike = k.cyan >= 50 && k.yellow >= 50 && k.magenta <= 35 && k.black <= 35;
        var isMagentaLike = k.magenta >= 70 && k.cyan <= 30 && k.yellow <= 30 && k.black <= 30;
        return isGreenLike || isMagentaLike;
    }

    return false;
}

function isNoFillPath(pathItem) {
    if (!pathItem) return false;
    if (!pathItem.filled) return true;
    if (!pathItem.fillColor) return true;
    return pathItem.fillColor.typename == "NoColor";
}

function getSizeMatchInfo(pathItem, mmL, mmS, tolMm) {
    if (!pathItem) {
        return { ok: false, flip: false, score: 999999 };
    }
    var tol = mm2pt((tolMm === undefined ? 3 : tolMm), false);
    var pw = Math.abs(pathItem.width);
    var ph = Math.abs(pathItem.height);
    var w1 = mm2pt(mmL, false);
    var h1 = mm2pt(mmS, false);
    var w2 = mm2pt(mmS, false);
    var h2 = mm2pt(mmL, false);

    var d1 = Math.abs(pw - w1) + Math.abs(ph - h1);
    var d2 = Math.abs(pw - w2) + Math.abs(ph - h2);

    if (d1 <= tol * 2) {
        return { ok: true, flip: false, score: d1 };
    }
    if (d2 <= tol * 2) {
        return { ok: true, flip: true, score: d2 };
    }
    return { ok: false, flip: (d2 < d1), score: Math.min(d1, d2) };
}

function isLikelyDieLineByGeometry(pathItem, mmL, mmS) {
    if (!pathItem || !pathItem.stroked) return false;
    if (pathItem.guides || pathItem.clipping) return false;
    if (!pathItem.closed) return false;
    if (!isNoFillPath(pathItem)) return false;
    return getSizeMatchInfo(pathItem, mmL, mmS, 6).ok;
}

function resizeItemToSize(item, targetW, targetH) {
    if (!item) return;
    var bw = Math.abs(item.width);
    var bh = Math.abs(item.height);
    if (bw <= 0.01 || bh <= 0.01) return;
    var sx = targetW * 100.0 / bw;
    var sy = targetH * 100.0 / bh;
    if (Math.abs(sx - 100.0) < 0.2 && Math.abs(sy - 100.0) < 0.2) return;
    var lineScale = (Math.abs(sx) + Math.abs(sy)) / 2.0;
    item.resize(sx, sy, true, true, true, true, lineScale, Transformation.CENTER);
}

function doAction() {
    function createAction(str) {
        var f = new File("~/ScriptAction.aia");
        f.open("w");
        f.write(str);
        f.close();
        app.loadAction(f);
        f.remove();
    }
    var actionString = [
        "/version 3",
        "/name [ 15",
        "\t4a5358e5afbce585a5e58aa8e4bd9c",
        "]",
        "/isOpen 1",
        "/actionCount 1",
        "/action-1 {",
        "\t/name [ 18",
        "\t\te8aebee7bdaee5afb9e9bd90e68f8fe8beb9",
        "\t]",
        "\t/keyIndex 0",
        "\t/colorIndex 0",
        "\t/isOpen 0",
        "\t/eventCount 1",
        "\t/event-1 {",
        "\t\t/useRulersIn1stQuadrant 0",
        "\t\t/internalName (ai_plugin_setStroke)",
        "\t\t/localizedName [ 12",
        "\t\t\te8aebee7bdaee68f8fe8beb9",
        "\t\t]",
        "\t\t/isOpen 0",
        "\t\t/isOn 1",
        "\t\t/hasDialog 0",
        "\t\t/parameterCount 1",
        "\t\t/parameter-1 {",
        "\t\t\t/key 1634494318",
        "\t\t\t/showInPalette -1",
        "\t\t\t/type (enumerated)",
        "\t\t\t/name [ 6",
        "\t\t\t\te5b185e4b8ad",
        "\t\t\t]",
        "\t\t\t/value 0",
        "\t\t}",
        "\t}",
        "}"
    ].join("\n");
    createAction(actionString);
    actionString = null;
    app.doScript("设置对齐描边", "JSX导入动作", false);
    app.unloadAction("JSX导入动作", "");
}

function findAllDieLines() {
    if (!currentDoc) {
        return "500;findAllDieLine未识别到文档";
    }
    var items = currentDoc.pathItems;
    app.executeMenuCommand("deselectall");
    while (currentDoc.selection.length > 0) {
        currentDoc.selection[0].selected = false;
    }
    for (var j = 0; j < items.length; j++) {
        var p = items[j];
        if (isDieLineStroke(p)) {
            var strokeDashes = p.strokeDashes;
            if (!strokeDashes || strokeDashes.length == 0) {
                p.selected = true;
            }
        }
    }
    doAction();
    return "200;刀线对齐描边居中设置完成。";
}

function startForDoc(data) {
    currentDoc = null;
    currentCutMap = {};
    var file = normalizeFilePath(data);
    if (!file) {
        return "500;ERR;文件路径为空";
    }
    var fileO = new File(decodeURI(file));
    if (!fileO.exists) {
        return "500;ERR;文件不存在:" + file;
    }
    currentDoc = app.open(fileO);
    return "200;文件打开成功";
}

function getMapCount() {
    var count = 0;
    for (var key in currentCutMap) {
        if (currentCutMap.hasOwnProperty(key)) {
            count++;
        }
    }
    return count;
}

function addNewDoc(data) {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档";
    }
    if (!getMapCount()) {
        return "2;未找到刀线;未找到符合尺寸的刀线";
    }
    var w = +data.split(";")[0];
    var h = +data.split(";")[1];
    var mmCx = +data.split(";")[2];
    var removeId = +data.split(";")[5] || 0;
    var docName = data.split(";")[6] || "";
    w = mm2pt(w + mmCx * 2, false);
    h = mm2pt(h + mmCx * 2, false);
    var y0 = 2000;
    var max_loop = Math.ceil(getMapCount() * (h + offset) / 4000);
    var each_loop = Math.ceil(getMapCount() / max_loop);
    for (var i = 0; i < getMapCount(); i++) {
        if (i % each_loop == 0) {
            y0 = 2000;
        }
        var start_1 = mm2pt(newArtboardSize) + max_right + Math.floor(i / each_loop) * (w + offset) * 2;
        var newArtboards = currentDoc.artboards.add([start_1, y0, start_1 + w, y0 - h]);
        var newArtboardsDie = currentDoc.artboards.add([start_1 + w + offset, y0, start_1 + w + offset + w, y0 - h]);
        if (docName != "") {
            newArtboards.name = docName;
            newArtboardsDie.name = docName;
            if (!currentArtboard[docName]) {
                currentArtboard[docName] = [];
            }
            $.writeln(currentDoc.artboards.length);
            currentArtboard[docName].push(currentDoc.artboards.length - 1);
            currentArtboard[docName].push(currentDoc.artboards.length);
        }
        y0 = y0 - h - offset;
    }
    currentDoc.artboards.remove(removeId);
    return "200;新建文档成功";
}

function findBounds(item) {
    var minX = Infinity;
    var minY = Infinity;
    var maxX = -Infinity;
    var maxY = -Infinity;

    function processItem(node) {
        if (node.typename == "TextFrame") {
            return processItem(node.createOutline());
        } else if (("pageItems" in node) == false) {
            var bounds = node.geometricBounds;
            minX = Math.min(minX, bounds[0]);
            minY = Math.min(minY, bounds[3]);
            maxX = Math.max(maxX, bounds[2]);
            maxY = Math.max(maxY, bounds[1]);
        } else if (!node.clipped) {
            var items = node.pageItems;
            for (var i = 0; i < items.length; i++) {
                processItem(items[i]);
            }
        } else {
            var clippedBounds = node.pageItems[0].geometricBounds;
            minX = Math.min(minX, clippedBounds[0]);
            minY = Math.min(minY, clippedBounds[3]);
            maxX = Math.max(maxX, clippedBounds[2]);
            maxY = Math.max(maxY, clippedBounds[1]);
        }
    }

    processItem(item);
    return [minX, minY, maxX, maxY];
}

function artboardInPathItems(data) {
    var bounds = data.geometricBounds;
    if (bounds[0] > artboardXY[0] && bounds[1] < artboardXY[1] && bounds[2] < artboardXY[2] && bounds[3] > artboardXY[3]) {
        return true;
    }
    return false;
}

function saveAs(data) {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档";
    }
    var file = normalizeFilePath(data);
    var saveFile = new File(decodeURI(file));
    var pdfSaveOpts = new PDFSaveOptions();
    pdfSaveOpts.compatibility = PDFCompatibility.ACROBAT6;
    pdfSaveOpts.acrobatLayers = true;
    pdfSaveOpts.viewAfterSaving = false;
    currentDoc.saveAs(saveFile, pdfSaveOpts);
    currentDoc.close(SaveOptions.DONOTSAVECHANGES);
    currentDoc = null;
    return "200;保存成功";
}

function getArtLines() {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档";
    }

    function pathInArt(path) {
        if (path.parent.typename == "GroupItem" && path.parent.clipped) {
            return null;
        }
        var bounds = path.geometricBounds;
        var tol = mm2pt(8, false);
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        var bestArt = null;
        var bestScore = 99999999;
        for (var art in currentCutMap) {
            var rect = currentCutMap[art].rect;
            if (bounds[0] >= (rect[0] - tol) && bounds[1] <= (rect[1] + tol) && bounds[2] <= (rect[2] + tol) && bounds[3] >= (rect[3] - tol)) {
                return art;
            }

            var nearTol = mm2pt(20, false);
            var dx = 0;
            var dy = 0;
            if (centerX < (rect[0] - nearTol)) dx = (rect[0] - nearTol) - centerX;
            else if (centerX > (rect[2] + nearTol)) dx = centerX - (rect[2] + nearTol);

            if (centerY > (rect[1] + nearTol)) dy = centerY - (rect[1] + nearTol);
            else if (centerY < (rect[3] - nearTol)) dy = (rect[3] - nearTol) - centerY;

            var score = dx + dy;
            if (score < bestScore) {
                bestScore = score;
                bestArt = art;
            }
        }

        if (bestArt && bestScore <= mm2pt(20, false)) {
            $.writeln("PATH_NEAR_ART: " + bestArt + " | score=" + bestScore + " | name=" + path.name);
            return bestArt;
        }

        path.selected = true;
        return null;
    }

    var items = currentDoc.pathItems;
    var currentLineCount = 0;
    var usedPathMap = {};
    var fallbackByArt = {};

    function getPathKey(pathItem) {
        try {
            return String(pathItem.uuid);
        } catch (e) {
        }
        try {
            return String(pathItem.name) + "|" + String(pathItem.width) + "|" + String(pathItem.height) + "|" + String(pathItem.geometricBounds);
        } catch (e2) {
        }
        return String(pathItem);
    }

    function getArtLineCount(artName) {
        var n = 0;
        for (var lineName in currentCutMap[artName].lines) {
            n++;
        }
        return n;
    }

    function addDetectedLine(artName, pathItem, sizeInfo, reason) {
        var key = getPathKey(pathItem);
        if (usedPathMap[key]) {
            return false;
        }
        usedPathMap[key] = true;

        currentLineCount++;
        var bounds = pathItem.geometricBounds;
        pathItem.name = "mqx_" + currentLineCount;
        currentCutMap[artName].lines[pathItem.name] = {
            index: currentLineCount,
            width: pathItem.width,
            height: pathItem.height,
            flip: !!sizeInfo.flip,
            itemsToGroup: [],
            itemsIndexToGroup: [],
            bounds: {
                minX: bounds[0] - mm2pt(currentCutMap[artName].mmCx, false),
                minY: bounds[3] - mm2pt(currentCutMap[artName].mmCx, false),
                maxX: bounds[2] + mm2pt(currentCutMap[artName].mmCx, false),
                maxY: bounds[1] + mm2pt(currentCutMap[artName].mmCx, false)
            }
        };

        if (reason) {
            $.writeln("USE_DIELINE[" + reason + "]: " + artName + " | w=" + pathItem.width + " h=" + pathItem.height + " | name=" + pathItem.name);
        }
        return true;
    }

    function rememberFallbackCandidate(artName, pathItem, sizeInfo, colorMatched) {
        var score = sizeInfo.score;
        if (colorMatched) score -= mm2pt(2, false);
        if (!pathItem.stroked) score += mm2pt(30, false);
        if (!pathItem.closed) score += mm2pt(20, false);
        if (!isNoFillPath(pathItem)) score += mm2pt(12, false);
        if (pathItem.guides || pathItem.clipping) score += mm2pt(50, false);

        var old = fallbackByArt[artName];
        if (!old || score < old.score) {
            fallbackByArt[artName] = {
                pathItem: pathItem,
                sizeInfo: sizeInfo,
                score: score
            };
        }
    }

    function canForceUseCandidate(artName, cand) {
        if (!cand || !cand.pathItem) return false;
        if (!cand.pathItem.stroked) return false;

        var targetW = mm2pt(currentCutMap[artName].mmL, false);
        var targetH = mm2pt(currentCutMap[artName].mmS, false);
        var limit = Math.max(mm2pt(15, false), (targetW + targetH) * 0.12);

        return cand.score <= limit;
    }

    for (var j = 0; j < items.length; j++) {
        var p = items[j];
        var art = pathInArt(p);
        if (art) {
            var mmL = currentCutMap[art].mmL;
            var mmS = currentCutMap[art].mmS;
            var sizeInfoStrict = getSizeMatchInfo(p, mmL, mmS, 3);
            var sizeInfoRelaxed = getSizeMatchInfo(p, mmL, mmS, 8);
            var colorMatched = isDieLineStroke(p);
            var matchedByColor = colorMatched;
            var matchedByGeom = false;

            rememberFallbackCandidate(art, p, sizeInfoRelaxed, colorMatched);

            if (!matchedByColor) {
                matchedByGeom = isLikelyDieLineByGeometry(p, mmL, mmS);
                if (matchedByGeom) {
                    sizeInfoRelaxed = getSizeMatchInfo(p, mmL, mmS, 6);
                    $.writeln("FALLBACK_DIELINE: " + art + " | w=" + p.width + " h=" + p.height + " | name=" + p.name);
                }
            }

            if (matchedByColor || matchedByGeom) {
                if (p.parent.typename == "CompoundPathItem") {
                    while (currentDoc.selection.length > 0) {
                        currentDoc.selection[0].selected = false;
                    }
                    p.selected = true;
                    app.executeMenuCommand("noCompoundPath");
                }
                addDetectedLine(art, p, (sizeInfoStrict.ok ? sizeInfoStrict : sizeInfoRelaxed), (matchedByColor ? "color" : "geometry"));
            }
        }
    }

    for (var artFallback in currentCutMap) {
        if (getArtLineCount(artFallback) <= 0 && fallbackByArt[artFallback]) {
            if (canForceUseCandidate(artFallback, fallbackByArt[artFallback])) {
                addDetectedLine(artFallback, fallbackByArt[artFallback].pathItem, fallbackByArt[artFallback].sizeInfo, "best_candidate");
            }
        }
    }

    for (var artName in currentCutMap) {
        currentCutMap[artName].kinds = Math.max(1, getArtLineCount(artName));
    }

    if (currentLineCount <= 0) {
        throw new Error("500;未识别到刀线");
    }

    var textFramesArray = [];
    for (var k = 0; k < currentDoc.textFrames.length; k++) {
        textFramesArray.push(currentDoc.textFrames[k]);
    }
    for (var n = 0; n < textFramesArray.length; n++) {
        textFramesArray[n].createOutline();
    }
}

function normalizeLineGroupsToFilename() {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档";
    }
    for (var artName in currentCutMap) {
        var artData = currentCutMap[artName];
        for (var lineName in artData.lines) {
            var cutData = artData.lines[lineName];
            var targetW = mm2pt(artData.mmL, false);
            var targetH = mm2pt(artData.mmS, false);
            if (cutData.flip) {
                var tmp = targetW;
                targetW = targetH;
                targetH = tmp;
            }
            try {
                resizeItemToSize(currentDoc.groupItems.getByName("group_" + lineName), targetW, targetH);
            } catch (eGroup) {
            }
            var lineObj = currentDoc.pathItems.getByName(lineName);
            resizeItemToSize(lineObj, targetW, targetH);
            var b2 = lineObj.geometricBounds;
            cutData.width = Math.abs(lineObj.width);
            cutData.height = Math.abs(lineObj.height);
            cutData.bounds = {
                minX: b2[0] - mm2pt(artData.mmCx, false),
                minY: b2[3] - mm2pt(artData.mmCx, false),
                maxX: b2[2] + mm2pt(artData.mmCx, false),
                maxY: b2[1] + mm2pt(artData.mmCx, false)
            };
        }
    }
    return "200;OK";
}

function addNewArt(cx) {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档";
    }
    var x0 = -2000;
    var y0 = 2000;
    var max_height = 0;
    var total_count = 0;
    for (var art in currentCutMap) {
        for (var line in currentCutMap[art].lines) {
            if (max_height < currentCutMap[art].lines[line].height) {
                max_height = currentCutMap[art].lines[line].height;
            }
            total_count++;
        }
    }
    var max_loop = Math.ceil((max_height + mm2pt(2 * cx, false)) * total_count / 4000);
    var each_loop = Math.ceil(total_count / max_loop);
    var i = 0;
    var x1 = 0;
    var maxx = 0;
    for (var artName in currentCutMap) {
        for (var lineName in currentCutMap[artName].lines) {
            var _w = currentCutMap[artName].lines[lineName].width + mm2pt(cx * 2, false);
            var _h = currentCutMap[artName].lines[lineName].height + mm2pt(cx * 2, false);
            if (y0 - _h < -8000) {
                x0 = maxx + art_offset;
                y0 = 2000;
            }
            i++;
            x1 = x0 + _w + art_offset;
            if (maxx < x1 + _w) {
                maxx = x1 + _w;
            }
            var artPrint = currentDoc.artboards.add([x0, y0, x0 + _w, y0 - _h]);
            var artLines = currentDoc.artboards.add([x1, y0, x1 + _w, y0 - _h]);
            artPrint.name = "art_1_" + lineName;
            artLines.name = "art_2_" + lineName;
            y0 -= _h + art_offset;
        }
    }
}

function renameAllItems() {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档";
    }

    function judgeInCutMap(p, lines, ib) {
        var pName = (p && p.name !== undefined && p.name !== null) ? String(p.name) : "";
        for (var name in lines) {
            if (pName == name) return;
            var isLeft = Math.round(ib[0] - lines[name].bounds.minX);
            if (isLeft < -1) continue;
            var isTop = Math.round(ib[1] - lines[name].bounds.minY);
            if (isTop < -1) continue;
            var isRight = Math.round(lines[name].bounds.maxX - ib[2]);
            if (isRight < -1) continue;
            var isBottom = Math.round(lines[name].bounds.maxY - ib[3]);
            if (isBottom < -1) continue;
            if (pName.indexOf("_g_") !== -1) return;
            p.name = name + "_g_" + (lines[name].itemsToGroup.length);
            lines[name].itemsToGroup.push(String(p.name));
            return;
        }
    }

    var totalCount = currentDoc.pageItems.length;
    for (var i = 0; i < totalCount; i++) {
        var p = currentDoc.pageItems[i];
        if (p.parent.typename == "GroupItem" || p.parent.typename == "CompoundPathItem") {
            continue;
        }
        var ib = findBounds(p);
        for (var art in currentCutMap) {
            judgeInCutMap(p, currentCutMap[art].lines, ib);
        }
    }
}

function moveLineToNewArt(data) {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档";
    }
    for (var art in currentCutMap) {
        for (var line in currentCutMap[art].lines) {
            var mqx = currentDoc.pathItems.getByName(line);
            var pos = currentDoc.artboards.getByName("art_2_" + line).artboardRect;
            var bounds = mqx.geometricBounds;
            if (currentCutMap[art].lines[line].flip > 3) {
                mqx.rotate(90, true, true, true, true, Transformation.CENTER);
            }
            var globX = (pos[0] + pos[2]) / 2;
            var globY = (pos[1] + pos[3]) / 2;
            var curX = (bounds[0] + bounds[2]) / 2;
            var curY = (bounds[1] + bounds[3]) / 2;
            mqx.translate(globX - curX, globY - curY);
        }
    }
}

function groupAllItemByLine() {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档";
    }
    for (var art in currentCutMap) {
        for (var line in currentCutMap[art].lines) {
            app.executeMenuCommand("deselectall");
            $.writeln("强制清空前选中的数量：" + currentDoc.selection.length);
            while (currentDoc.selection.length > 0) {
                currentDoc.selection[0].selected = false;
            }
            $.writeln("强制清空后选中的数量：" + currentDoc.selection.length);
            var cutData = currentCutMap[art].lines[line];
            $.writeln("当前画板名称：" + art);
            $.writeln("当前刀线名称：" + line);
            var selectedItems = [];
            for (var i = cutData.itemsToGroup.length - 1; i >= 0; i--) {
                $.writeln("当前元素名称：" + cutData.itemsToGroup[i]);
                $.writeln("当前：" + cutData.itemsToGroup.length);
                var p = currentDoc.pageItems.getByName(cutData.itemsToGroup[i]);
                if (isDieLineStroke(p)) {
                    var mqx = currentDoc.pathItems.getByName(line);
                    var pos = currentDoc.artboards.getByName("art_2_" + line).artboardRect;
                    var bounds = mqx.geometricBounds;
                    if (currentCutMap[art].lines[line].flip > 3) {
                        p.rotate(90, true, true, true, true, Transformation.CENTER);
                    }
                    var globX = (pos[0] + pos[2]) / 2;
                    var globY = (pos[1] + pos[3]) / 2;
                    var curX = (bounds[0] + bounds[2]) / 2;
                    var curY = (bounds[1] + bounds[3]) / 2;
                    p.translate(globX - curX, globY - curY);
                    continue;
                }
                selectedItems.push(p);
            }
            currentDoc.selection = selectedItems;
            app.executeMenuCommand("group");
            var currentSelection = currentDoc.selection;
            $.writeln("重命名编组时候选中的元素数量：" + currentSelection.length);
            for (var s = 0; s < currentSelection.length; s++) {
                if (!currentSelection[s].name) {
                    var group = currentSelection[s];
                    group.name = "group_" + line;
                    group.selected = false;
                    break;
                }
            }
            app.executeMenuCommand("deselectall");
        }
    }
}

function moveAllItemByLine() {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档";
    }
    for (var art in currentCutMap) {
        for (var line in currentCutMap[art].lines) {
            var cutData = currentCutMap[art].lines[line];
            var group = currentDoc.groupItems.getByName("group_" + line);
            if (cutData.flip > 3) {
                group.rotate(90, true, true, true, true, Transformation.CENTER);
            }
            var pos = currentDoc.artboards.getByName("art_1_" + line).artboardRect;
            var bounds = findBounds(group);
            var globX = (pos[0] + pos[2]) / 2;
            var globY = (pos[1] + pos[3]) / 2;
            var curX = (bounds[0] + bounds[2]) / 2;
            var curY = (bounds[1] + bounds[3]) / 2;
            group.translate(globX - curX, globY - curY);
        }
    }
}

function ungroup(obj) {
    var elements = [];
    var items = obj.pageItems;
    for (var i = 0; i < items.length; i += 1) {
        elements.push(items[i]);
    }
    if (elements.length < 1) {
        obj.remove();
        return;
    }
    for (var j = 0; j < elements.length; j += 1) {
        try {
            if (elements[j].parent.typename != "Layer") {
                elements[j].moveBefore(obj);
            }
            if (elements[j].typename == "GroupItem" && !elements[j].clipped && elements[j].blendingMode == BlendModes.NORMAL && elements[j].opacity == 100) {
                ungroup(elements[j]);
            }
        } catch (e) {
            $.writeln(e);
        }
    }
}

function main(data) {
    try {
        currentDoc = null;
        currentCutMap = {};
        currentArtboard = {};
        end = 1;

        var args = parseRunArgs(data);
        var openRet = startForDoc(args.filePath);
        if (String(openRet).indexOf("200;") !== 0 || !currentDoc) {
            return openRet || "1;未找到文档;未找到打开的文档";
        }

        var sum = 0;
        var cx = +args.cx;
        var filePath = normalizeWinPath(args.filePath);
        var result = filePath.split("\\");
        var endPath = [];
        for (var i = 0; i < result.length; i++) {
            if (i == (result.length - 1)) {
                continue;
            }
            endPath.push(result[i]);
        }
        var downPath = args.outDir ? normalizeWinPath(args.outDir) : endPath.join("\\");
        var baseName = result[result.length - 1];
        var fileMetaName = baseName.substring(0, baseName.lastIndexOf("."));
        var fileMetaParts = fileMetaName.split("^");
        var fileDims = fileMetaParts.length > 6 ? parseSizeToken(fileMetaParts[6]) : null;
        var fileKinds = +(fileMetaParts[2] || 0);
        if (!(fileKinds > 0)) fileKinds = 1;
        if (!fileDims) {
            throw new Error("文件名尺寸格式错误:" + fileMetaName);
        }

        var artboardsLength = currentDoc.artboards.length;
        if (artboardsLength == 1) {
            currentArtboard = currentDoc.artboards[0];
            currentCutMap[fileMetaName] = {
                index: 0,
                mmL: +fileDims[0],
                mmS: +fileDims[1],
                mmCx: cx,
                kinds: fileKinds,
                rect: currentArtboard.artboardRect,
                lines: {}
            };
            sum = currentCutMap[fileMetaName].kinds * 2;
        } else {
            for (var k = 0; k < artboardsLength; k++) {
                currentArtboard = currentDoc.artboards[k];
                var artName = currentArtboard.name;
                var artParts = String(artName).split("^");
                var artDims = artParts.length > 6 ? parseSizeToken(artParts[6]) : null;
                var artKinds = +(artParts[2] || 0);
                if (!artDims) artDims = fileDims;
                if (!(artKinds > 0)) artKinds = 1;
                if (!artName || artName.indexOf("^") < 0) artName = fileMetaName + "_ART" + (k + 1);
                currentCutMap[artName] = {
                    index: k,
                    mmL: +artDims[0],
                    mmS: +artDims[1],
                    mmCx: cx,
                    kinds: artKinds,
                    rect: currentArtboard.artboardRect,
                    lines: {}
                };
                sum += currentCutMap[artName].kinds * 2;
            }
        }

        for (var layerIdx = currentDoc.layers.length - 1; layerIdx >= 0; layerIdx--) {
            if (currentDoc.groupItems.length) {
                ungroup(currentDoc.layers[layerIdx]);
            }
        }

        var pageItems = currentDoc.pageItems;
        for (var m = 0; m < pageItems.length; m++) {
            try {
                var item = pageItems[m];
                if (item && item.name !== undefined) {
                    item.name = "1";
                }
            } catch (e1) {
            }
        }

        findAllDieLines();
        getArtLines();
        renameAllItems();
        groupAllItemByLine();
        normalizeLineGroupsToFilename();
        addNewArt(cx);
        moveLineToNewArt(cx);
        moveAllItemByLine();

        var remove_ids = [];
        if (artboardsLength > 1) {
            for (var art in currentCutMap) {
                remove_ids.push(currentCutMap[art].index);
            }
            for (var ridx = remove_ids.length - 1; ridx >= 0; ridx--) {
                currentDoc.artboards.remove(remove_ids[ridx]);
            }
        } else {
            currentDoc.artboards.remove(0);
        }

        if (sum !== currentDoc.artboards.length) {
            throw new Error("警告：文档中存在多余的画板，请检查。");
        }

        for (var artName2 in currentCutMap) {
            var kinds = currentCutMap[artName2].kinds;
            var newFile = downPath + "\\" + artName2 + ".pdf";
            var saveFile = new File(decodeURI(newFile));
            var pdfSaveOpts = new PDFSaveOptions();
            pdfSaveOpts.compatibility = PDFCompatibility.ACROBAT6;
            pdfSaveOpts.acrobatLayers = true;
            pdfSaveOpts.viewAfterSaving = false;
            pdfSaveOpts.saveMultipleArtboards = true;
            pdfSaveOpts.cropToArtboard = true;
            pdfSaveOpts.preserveEditability = false;
            pdfSaveOpts.generateThumbnails = true;
            pdfSaveOpts.embedFont = true;
            pdfSaveOpts.embedImages = true;
            pdfSaveOpts.artboardRange = end + "-" + (kinds * 2 + end - 1);
            end = end + kinds * 2;
            currentDoc.saveAs(saveFile, pdfSaveOpts);
        }

        try {
            currentDoc.close(SaveOptions.DONOTSAVECHANGES);
        } catch (eClose) {
        }
        currentDoc = null;
        return "200;OK";
    } catch (error) {
        try {
            if (currentDoc) {
                currentDoc.close(SaveOptions.DONOTSAVECHANGES);
            }
        } catch (e2) {
        }
        currentDoc = null;
        return "500;ERR;" + error + ((error && error.line) ? (" @line=" + error.line) : "");
    }
}
