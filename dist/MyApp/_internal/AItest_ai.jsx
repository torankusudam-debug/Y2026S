var art_offset = 250;
var currentDoc = null;
var currentCutMap = {};
var currentArtboard = {}; // 用于存画板拆分后的画板索引
var artboardXY = []; //左上右下
var end = 1;
var mqxSize = 0;
var newArtboardSize = 0;

function normalizeFilePath(path) {
    return String(path || "").replace(/\\/g, "/");
}

function normalizeWinPath(path) {
    return String(path || "").replace(/\//g, "\\");
}

function safeOutputName(name) {
    var s = String(name || "");
    s = s.replace(/[\\\/:\*\?"<>\|]/g, "_");
    s = s.replace(/[\r\n\t]/g, " ");
    s = s.replace(/\s+/g, " ");
    s = s.replace(/[\.\s]+$/g, "");
    if (!s) s = "output";
    if (s.length > 120) s = s.substring(0, 120);
    return s;
}

function parseRunArgs(data) {
    var s = String(data || "");
    var p1 = s.indexOf(";");
    var p2 = s.indexOf(";", p1 + 1);
    if (p1 < 0) return { cx: (+s || 0), filePath: "", outDir: "" };
    if (p2 < 0) {
        return { cx: (+(s.substring(0, p1)) || 0), filePath: s.substring(p1 + 1), outDir: "" };
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
    var pt = UnitValue(mm, "mm").as("pt")
    if (needRound) return Math.round(pt);
    return pt;
}

function esIndexOf(searchElement) {
    if (!searchElement) return -1;
    searchElement = String(searchElement).toUpperCase();
    var array = ["PANTONE 802 C", "PANTONE 802C", "CUT", "DIE", "DIELINE", "刀", "模切"];
    for (var i = 0; i < array.length; i++)
        if (searchElement.indexOf(array[i]) !== -1)
            return i; // 找到，返回索引 (>= 0)
    return -1; // 未找到，返回 -1
}

function isDieLineStroke(pathItem) {
    if (!pathItem || !pathItem.stroked || !pathItem.strokeColor) return false;
    if (pathItem.strokeColor.typename == "SpotColor") {
        return !!pathItem.strokeColor.spot && esIndexOf(pathItem.strokeColor.spot.name) != -1;
    }
    if (pathItem.strokeColor.typename == "RGBColor") {
        var c = pathItem.strokeColor;
        return c.green >= 200 && c.red <= 80 && c.blue <= 80;
    }
    if (pathItem.strokeColor.typename == "CMYKColor") {
        var k = pathItem.strokeColor;
        return (k.cyan >= 50 && k.yellow >= 50 && k.magenta <= 35 && k.black <= 35) ||
               (k.magenta >= 70 && k.cyan <= 30 && k.yellow <= 30 && k.black <= 30);
    }
    return false;
}

function resizeItemToSize(item, targetW, targetH) {
    if (!item) return;
    var bw = Math.abs(item.width), bh = Math.abs(item.height);
    if (bw <= 0.01 || bh <= 0.01) return;
    var sx = targetW * 100.0 / bw;
    var sy = targetH * 100.0 / bh;
    if (Math.abs(sx - 100.0) < 0.2 && Math.abs(sy - 100.0) < 0.2) return;
    item.resize(sx, sy, true, true, true, true, (Math.abs(sx) + Math.abs(sy)) / 2.0, Transformation.CENTER);
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
        "	4a5358e5afbce585a5e58aa8e4bd9c",
        "]",
        "/isOpen 1",
        "/actionCount 1",
        "/action-1 {",
        "	/name [ 18",
        "		e8aebee7bdaee5afb9e9bd90e68f8fe8beb9",
        "	]",
        "	/keyIndex 0",
        "	/colorIndex 0",
        "	/isOpen 0",
        "	/eventCount 1",
        "	/event-1 {",
        "		/useRulersIn1stQuadrant 0",
        "		/internalName (ai_plugin_setStroke)",
        "		/localizedName [ 12",
        "			e8aebee7bdaee68f8fe8beb9",
        "		]",
        "		/isOpen 0",
        "		/isOn 1",
        "		/hasDialog 0",
        "		/parameterCount 1",
        "		/parameter-1 {",
        "			/key 1634494318",
        "			/showInPalette -1",
        "			/type (enumerated)",
        "			/name [ 6",
        "				e5b185e4b8ad",
        "			]",
        "			/value 0",
        "		}",
        "	}",
        "}"
    ].join("\n");
    createAction(actionString);
    var actionString = null;
    app.doScript("设置对齐描边", "JSX导入动作", false);
    app.unloadAction("JSX导入动作", "");
}

function findAllDieLines() {
    if (!currentDoc) {
        return "500;findAllDieLine未识别到文档";
    }
    //元素
    var items = currentDoc.pathItems;
    app.executeMenuCommand("deselectall");
    while (currentDoc.selection.length > 0) {
        currentDoc.selection[0].selected = false;
    }
    // 遍历所有元素，查看第一层是否有刀线的元素
    for (var j = 0; j < items.length; j++) {
        var p = items[j];
        //判断是否是刀线
        if (isDieLineStroke(p)) {
            if (!p.strokeDashes || p.strokeDashes.length == 0) {
                p.selected = true;
            }
            var t = p.textRange;
        }
    }
    doAction();
    return "200;刀线对齐描边居中设置完成。"
}

function startForDoc(data) {
    currentDoc = null;
    currentCutMap = {};
    var file = normalizeFilePath(data);
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
        return "1;未找到文档;未找到打开的文档"
    }
    if (!getMapCount()) {
        return "2;未找到刀线;未找到符合尺寸的刀线"
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

    function processItem(item) {
        if (item.typename == "TextFrame") {
            return processItem(item.createOutline())
        } else if ("pageItems" in item == false) {
            var bounds = item.geometricBounds;
            minX = Math.min(minX, bounds[0]);
            minY = Math.min(minY, bounds[3]);
            maxX = Math.max(maxX, bounds[2]);
            maxY = Math.max(maxY, bounds[1]);
        } else if (!item.clipped) {
            var items = item.pageItems;
            for (var i = 0; i < items.length; i++) {
                processItem(items[i]);
            }
        } else {
            var bounds = item.pageItems[0].geometricBounds;
            minX = Math.min(minX, bounds[0]);
            minY = Math.min(minY, bounds[3]);
            maxX = Math.max(maxX, bounds[2]);
            maxY = Math.max(maxY, bounds[1]);
        }
    }

    processItem(item);
    return [minX, minY, maxX, maxY];
}

/**
 * 判断对象是不是在这个画板里面
 */
function artboardInPathItems(data) {
    var bounds = data.geometricBounds; //获取左x，上y，右x，下y的坐标(获取图像的边界)
    if (bounds[0] > artboardXY[0] && bounds[1] < artboardXY[1] && bounds[2] < artboardXY[2] && bounds[3] > artboardXY[3]) {
        return true;
    } else {
        return false;
    }
}


function saveAs(data) {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档"
    }
    var file = normalizeFilePath(data);
    var saveFile = new File(decodeURI(file));
    try {
        if (saveFile.parent && !saveFile.parent.exists) saveFile.parent.create();
    } catch (eCreate) {}
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
        return "1;未找到文档;未找到打开的文档"
    }

    function pathInArt(path) {
        if (path.parent.typename == "GroupItem" && path.parent.clipped) {
            return null
        }
        var bounds = path.geometricBounds; //获取左x，上y，右x，下y的坐标(获取图像的边界)
        for (var art in currentCutMap) {
            var rect = currentCutMap[art].rect;
            if (bounds[0] >= rect[0] && bounds[1] <= rect[1] && bounds[2] <= rect[2] && bounds[3] >= rect[3]) {
                return art;
            }
        }
        path.selected = true
    }
    var items = currentDoc.pathItems; //获取这个图层下的所有对象
    var currentLineCount = 0;

    for (var j = 0; j < items.length; j++) {
        var p = items[j];
        var art = pathInArt(p);
        if (art) {
            var mmL = currentCutMap[art].mmL;
            var mmS = currentCutMap[art].mmS;
            var mmCx = currentCutMap[art].mmCx;
            if (isDieLineStroke(p)) {
                if(p.parent.typename =="CompoundPathItem"){
                    while (currentDoc.selection.length > 0) {
                        currentDoc.selection[0].selected = false;
                    }
                    p.selected = true;
                    app.executeMenuCommand('noCompoundPath');
                }
                var pw = p.width;
                var ph = p.height;
                var isSameFlip1 = Math.abs(pw - mm2pt(mmL, false)) < mm2pt(8, false) && Math.abs(ph - mm2pt(mmS, false)) < mm2pt(8, false);
                var isSameFlip2 = Math.abs(pw - mm2pt(mmS, false)) < mm2pt(8, false) && Math.abs(ph - mm2pt(mmL, false)) < mm2pt(8, false);
                var flip = isSameFlip2 && !isSameFlip1;
                currentLineCount++;
                var bounds = p.geometricBounds;
                p.name = 'mqx_' + currentLineCount;
                currentCutMap[art]['lines'][p.name] = {
                    index: currentLineCount,
                    width: pw,
                    height: ph,
                    flip: flip,
                    itemsToGroup: [],
                    itemsIndexToGroup: [],
                    bounds: {
                        minX: bounds[0] - mm2pt(mmCx, false),
                        minY: bounds[3] - mm2pt(mmCx, false),
                        maxX: bounds[2] + mm2pt(mmCx, false),
                        maxY: bounds[1] + mm2pt(mmCx, false),
                    }
                };
               
            }
        }
    }
    for (var art in currentCutMap) {
        var lc = 0;
        for (var ln in currentCutMap[art]['lines']) lc++;
        currentCutMap[art].kinds = Math.max(1, lc);
    }
    if (currentLineCount <= 0) {
        throw new Error("500;未识别到刀线");
    }
    //转曲使用副本数组进行处理，实时访问文本类的数据时候length的长度会实时更新所以采用副本处理
    var textFramesArray = [];
    for (var k = 0; k < currentDoc.textFrames.length; k++) {
        textFramesArray.push(currentDoc.textFrames[k]);
    }
    for (var k = 0; k < textFramesArray.length; k++) {
        textFramesArray[k].createOutline();
    }
}

function addNewArt(cx) {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档"
    }
    var x0 = -2000;
    var y0 = 2000;
    var max_height = 0;
    var total_count = 0;
    for (var art in currentCutMap) {
        for (var line in currentCutMap[art]['lines']) {
            if (max_height < currentCutMap[art]['lines'][line].height)
                max_height = currentCutMap[art]['lines'][line].height;
            total_count++;
        }
    }
    var max_loop = Math.ceil((max_height + mm2pt(2 * cx, false)) * total_count / 4000);
    var each_loop = Math.ceil(total_count / max_loop);
    var i = 0;
    var x1 = 0;
    var maxx = 0;
    for (var art in currentCutMap) {
        for (var line in currentCutMap[art]['lines']) {
            var _w = currentCutMap[art]['lines'][line].width + mm2pt(cx * 2, false);
            var _h = currentCutMap[art]['lines'][line].height + mm2pt(cx * 2, false);
            // if (i % each_loop == 0) {
            //     x0 = maxx + _w + art_offset;
            //     y0 = 6000;
            // }
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
            artPrint.name = "art_1_" + line;
            artLines.name = "art_2_" + line;
            y0 -= _h + art_offset;
        }
    }
}

function renameAllItems() {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档"
    }

    function judgeInCutMap(p, lines) {
        for (var name in lines) {
            if (p.name == name) return;
            var isLeft = Math.round(ib[0] - lines[name].bounds.minX);
            if (isLeft < -1) continue;
            var isTop = Math.round(ib[1] - lines[name].bounds.minY);
            if (isTop < -1) continue;
            var isRight = Math.round(lines[name].bounds.maxX - ib[2]);
            if (isRight < -1) continue;
            var isBottom = Math.round(lines[name].bounds.maxY - ib[3]);
            if (isBottom < -1) continue;
            if (p.name.indexOf("_g_") !== -1) return;
            p.name = name + "_g_" + (lines[name].itemsToGroup.length);
            lines[name].itemsToGroup.push(p.name);
            return
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
            judgeInCutMap(p, currentCutMap[art].lines);
        }
    }
}



function moveLineToNewArt(data) {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档"
    }
    for (var art in currentCutMap) {
        for (var line in currentCutMap[art]['lines']) {
            var mqx = currentDoc.pathItems.getByName(line);
            var pos = currentDoc.artboards.getByName("art_2_" + line).artboardRect;
            var bounds = mqx.geometricBounds;
            if (currentCutMap[art]['lines'][line].flip > 3) {
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
        return "1;未找到文档;未找到打开的文档"
    }
    for (var art in currentCutMap) {
        for (var line in currentCutMap[art]['lines']) {
            app.executeMenuCommand("deselectall");
            $.writeln("强制清空前选中的数量：" + currentDoc.selection.length);
            while (currentDoc.selection.length > 0) {
                currentDoc.selection[0].selected = false;
            }
            $.writeln("强制清空后选中的数量：" + currentDoc.selection.length);
            var cutData = currentCutMap[art]['lines'][line];
            $.writeln("当前画板名称：" + art);
            $.writeln("当前刀线名称：" + line);
            var selectedItems = [];
            // function selectPluginBlendItems (group) {
                
            // }
            for (var i = cutData.itemsToGroup.length - 1; i >= 0; i--) {
                $.writeln("当前元素名称：" + cutData.itemsToGroup[i]);
                $.writeln("当前：" + cutData.itemsToGroup.length);
                var p = currentDoc.pageItems.getByName(cutData.itemsToGroup[i]);
                if (isDieLineStroke(p)) {
                    var mqx = currentDoc.pathItems.getByName(line);
                    var pos = currentDoc.artboards.getByName("art_2_" + line).artboardRect;
                    var bounds = mqx.geometricBounds;
                    if (currentCutMap[art]['lines'][line].flip > 3) {
                        p.rotate(90, true, true, true, true, Transformation.CENTER);
                    }
                    var globX = (pos[0] + pos[2]) / 2;
                    var globY = (pos[1] + pos[3]) / 2;
                    var curX = (bounds[0] + bounds[2]) / 2;
                    var curY = (bounds[1] + bounds[3]) / 2;
                    p.translate(globX - curX, globY - curY);
                    continue;
                }
                // if (p.typename == "PluginItem") {
                //     selectPluginBlendItems(p);
                // }
                selectedItems.push(p);
                // p.selected = true;
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
            };
            app.executeMenuCommand("deselectall");
        }
    }
}

function moveAllItemByLine() {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档"
    }
    for (var art in currentCutMap) {
        for (var line in currentCutMap[art]['lines']) {
            var cutData = currentCutMap[art]['lines'][line];
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
    var elements = new Array();
    var items = obj.pageItems;
    for (var i = 0; i < items.length; i += 1) {
        elements.push(items[i]);
    }
    if (elements.length < 1) {
        obj.remove();
        return;
    } else {
        for (var i = 0; i < elements.length; i += 1) {
            try {
                if (elements[i].parent.typename != "Layer") {
                    elements[i].moveBefore(obj);
                }
                if (elements[i].typename == "GroupItem" && !elements[i].clipped && elements[i].blendingMode == BlendModes.NORMAL && elements[i].opacity == 100 
                // && elements[i].appearanceAttributesCount == 1
                ) {
                    // $.writeln("当前元素名称：" + elements[i].name + elements[i].blendingMode);
                    var t = elements[i];
                    ungroup(elements[i]);
                }
            } catch (e) {
                $.writeln(e);
            }
        }
    }
}

function main(data) {
    try {
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
        var downFolder = new Folder(normalizeFilePath(downPath));
        if (!downFolder.exists) {
            try { downFolder.create(); } catch (eMk) {}
        }
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
            sum = fileKinds * 2;
        } else {
            for (var i = 0; i < artboardsLength; i++) {
                currentArtboard = currentDoc.artboards[i];
                var artName = currentArtboard.name;
                var artParts = String(artName).split("^");
                var artDims = artParts.length > 6 ? parseSizeToken(artParts[6]) : fileDims;
                var artKinds = +(artParts[2] || 0);
                if (!(artKinds > 0)) artKinds = fileKinds;
                if (!artName || artName.indexOf("^") < 0) artName = fileMetaName + "_ART" + (i + 1);
                currentCutMap[artName] = {
                    index: i,
                    mmL: +artDims[0],
                    mmS: +artDims[1],
                    mmCx: cx,
                    kinds: artKinds,
                    rect: currentArtboard.artboardRect,
                    lines: {}
                };
                sum += artKinds * 2;
            }
        }
        for (var i = currentDoc.layers.length - 1; i >= 0; i--) { //遍历图层
            if (currentDoc.groupItems.length) { //获取图层下的编组
                ungroup(currentDoc.layers[i]); //将所有编组的子对象都解组
            }
        }
        pageItems = currentDoc.pageItems;
        for (var i = 0; i < currentDoc.pageItems.length; i++) {
            try {
                var item = pageItems[i];
                
                // 检查是否可以设置 name 属性
                if (item.hasOwnProperty("name") || item.name !== undefined) {
                    item.name = "1";
                    count++;
                }
            } catch (e) {

            }
        }
        findAllDieLines();
        getArtLines();
        renameAllItems();
        groupAllItemByLine();

        sum = 0;
        for (var art2 in currentCutMap) {
            for (var line2 in currentCutMap[art2]['lines']) {
                var cutData = currentCutMap[art2]['lines'][line2];
                var targetW = mm2pt(currentCutMap[art2].mmL, false);
                var targetH = mm2pt(currentCutMap[art2].mmS, false);
                if (cutData.flip) {
                    var tmp = targetW; targetW = targetH; targetH = tmp;
                }
                try { resizeItemToSize(currentDoc.groupItems.getByName("group_" + line2), targetW, targetH); } catch (eResize1) {}
                try { resizeItemToSize(currentDoc.pathItems.getByName(line2), targetW, targetH); } catch (eResize2) {}
                var lineObj = currentDoc.pathItems.getByName(line2);
                var b2 = lineObj.geometricBounds;
                cutData.width = Math.abs(lineObj.width);
                cutData.height = Math.abs(lineObj.height);
                cutData.bounds = {
                    minX: b2[0] - mm2pt(currentCutMap[art2].mmCx, false),
                    minY: b2[3] - mm2pt(currentCutMap[art2].mmCx, false),
                    maxX: b2[2] + mm2pt(currentCutMap[art2].mmCx, false),
                    maxY: b2[1] + mm2pt(currentCutMap[art2].mmCx, false)
                };
            }
            sum += currentCutMap[art2].kinds * 2;
        }

        addNewArt(cx);
        moveLineToNewArt(cx) //复制刀线到画板中
        moveAllItemByLine();

        var remove_ids = []
        if (artboardsLength > 1) {
            for (var art in currentCutMap) {
                remove_ids.push(currentCutMap[art].index)
            }
            for (var i = remove_ids.length - 1; i >= 0; i--) {
                currentDoc.artboards.remove(remove_ids[i]);
            }
        } else {
            currentDoc.artboards.remove(0);
        }


        if (sum !== currentDoc.artboards.length) {
            throw new Error("警告：文档中存在多余的画板，请检查。");
        }

        for (var art in currentCutMap) {
            var kinds = currentCutMap[art].kinds;
            var newFile = normalizeFilePath(downPath + "\\" + safeOutputName(art) + ".pdf");
            var saveFile = new File(decodeURI(newFile));
            try {
                if (saveFile.parent && !saveFile.parent.exists) saveFile.parent.create();
            } catch (eCreate2) {}
            var pdfSaveOpts = new PDFSaveOptions();
            pdfSaveOpts.compatibility = PDFCompatibility.ACROBAT6;
            pdfSaveOpts.acrobatLayers = true;//保留图层
            pdfSaveOpts.viewAfterSaving = false;//不打开保存的pdf文件
            pdfSaveOpts.saveMultipleArtboards = false;//只保留选中的艺术板
            pdfSaveOpts.cropToArtboard = true;// 根据艺术板裁剪页面
            pdfSaveOpts.preserveEditability = false;//是否保留AI可编辑性
            pdfSaveOpts.generateThumbnails = true;
            pdfSaveOpts.embedFont = true;//嵌入字体
            pdfSaveOpts.embedImages = true;//嵌入图片
            pdfSaveOpts.artboardRange = end + "-" + (kinds * 2 + end - 1);
            end = end + kinds * 2;
            currentDoc.saveAs(saveFile, pdfSaveOpts);
        }

        try {
            currentDoc.close(SaveOptions.DONOTSAVECHANGES);
        } catch (eClose) {}
        currentDoc = null;
        return "200;OK";
    } catch (error) {
        try {
            if (currentDoc) currentDoc.close(SaveOptions.DONOTSAVECHANGES);
        } catch (e2) {}
        currentDoc = null;
        return "500;ERR;" + error + ((error && error.line) ? (" @line=" + error.line) : "");
    } finally {
    }
}