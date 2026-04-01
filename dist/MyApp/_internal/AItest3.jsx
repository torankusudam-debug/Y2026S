﻿var art_offset = 250;
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

function startForDoc(data) {
    currentDoc = null;
    currentCutMap = {};
    var file = normalizeFilePath(data);
    var fileO = new File(decodeURI(file));
    currentDoc = app.open(fileO);
    return "200;文件打开成功";
}


function mm2pt(mm, needRound) {
    var pt = UnitValue(mm, "mm").as("pt")
    if (needRound) return Math.round(pt);
    return pt;
}

function findBounds(item) {
    var minX = Infinity;
    var minY = Infinity;
    var maxX = -Infinity;
    var maxY = -Infinity;

    function processItem(item) {
        if (item.typename == "TextFrame") {
            var bounds = item.geometricBounds;
            minX = Math.min(minX, bounds[0]);
            minY = Math.min(minY, bounds[3]);
            maxX = Math.max(maxX, bounds[2]);
            maxY = Math.max(maxY, bounds[1]);
            minX = (minX + maxX) /2;
            maxX = minX;
            minY = (minY + maxY )/2;
            maxY = minY;
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

function esIndexOf(searchElement) {
    var array = ["PANTONE 802 C"];
    for (var i = 0; i < array.length; i++)
        if (array[i] === searchElement)
            return i; // 找到，返回索引 (>= 0)
    return -1; // 未找到，返回 -1
}

function deleteAllActions() {
    // 清除回与脚本冲突的动作
    try {
        app.unloadAction("移动编组", "");
    } catch (e) {
        $.writeln("警告: 删除移动编组动作失败")
    }
    try {
        app.unloadAction("JSX导入动作", "");
    } catch (e) {
        $.writeln("警告: 删除JSX导入动作失败")
    }
}

function doMoveItem(deltaX, deltaY) {
    function createAction(str) {
        var f = new File("~/MoveAction.aia");
        f.open("w");
        f.write(str);
        f.close();
        app.loadAction(f);
        f.remove();
    }
    var actionString = [
        "/version 3",
        "/name [ 12",
        "    e7a7bbe58aa8e7bc96e7bb84",
        "]",
        "/isOpen 1",
        "/actionCount 1",
        "/action-1 {",
        "    /name [ 12",
        "        e7a7bbe58aa8e58aa8e4bd9c",
        "    ]",
        "    /keyIndex 0",
        "    /colorIndex 0",
        "    /isOpen 0",
        "    /eventCount 1",
        "    /event-1 {",
        "        /useRulersIn1stQuadrant 0",
        "        /internalName (adobe_move)",
        "        /localizedName [ 6",
        "            e7a7bbe58aa8",
        "        ]",
        "        /isOpen 0",
        "        /isOn 1",
        "        /hasDialog 1",
        "        /showDialog 0",
        "        /parameterCount 3",
        "        /parameter-1 {",
        "            /key 1752136302",
        "            /showInPalette -1",
        "            /type (unit real)",
        "            /value " + deltaX ,
        "            /unit 592476268",
        "        }",
        "        /parameter-2 {",
        "            /key 1987339116",
        "            /showInPalette -1",
        "            /type (unit real)",
        "            /value " + deltaY ,
        "            /unit 592476268",
        "        }",
        "        /parameter-3 {",
        "            /key 1668247673",
        "            /showInPalette -1",
        "            /type (boolean)",
        "            /value 0",
        "        }",
        "    }",
        "}"
    ].join("\n");
    createAction(actionString);
    var actionString = null;
    app.doScript("移动动作", "移动编组", false);
    app.unloadAction("移动编组", "");    
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


// 选中元素 排除了无法选中的元素的干扰
function selectedItmes(obj) {
    //TODO 应回到原父元素下
    var originalParent = obj.parent;
    var tempLayer = currentDoc.layers.add();
    tempLayer.name = "PROCESSING";
    obj.move(tempLayer, ElementPlacement.PLACEATEND);
    activeDocument.activeLayer.hasSelectedArtwork = true;
    obj.move(originalParent, ElementPlacement.PLACEATEND)
    tempLayer.remove();
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
        if (p.stroked && p.strokeColor.typename == "SpotColor" && esIndexOf(p.strokeColor.spot.name) != -1) {
            if (p.strokeDashes.length == 0 || !p.strokeDashes) {
                p.selected = true;
            }
            var t = p.textRange;
        }
    }
    doAction();
    return "200;刀线对齐描边居中设置完成。"
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

    function findDieLinesGroupBounds(item) {
        var minX = Infinity;
        var minY = Infinity;
        var maxX = -Infinity;
        var maxY = -Infinity;

        function processItem(item) {
            if (item.typename == "TextFrame") {
                var bounds = item.geometricBounds;
                minX = Math.min(minX, bounds[0]);
                minY = Math.min(minY, bounds[3]);
                maxX = Math.max(maxX, bounds[2]);
                maxY = Math.max(maxY, bounds[1]);
                minX = (minX + maxX) /2;
                maxX = minX;
                minY = (minY + maxY )/2;
                maxY = minY;
            } else if ("pageItems" in item == false) {
                if (item.name !== "2") {
                    var bounds = item.geometricBounds;
                    minX = Math.min(minX, bounds[0]);
                    minY = Math.min(minY, bounds[3]);
                    maxX = Math.max(maxX, bounds[2]);
                    maxY = Math.max(maxY, bounds[1]);
                }
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
    //1.获取所有的刀线,且移到图层下
    var items = currentDoc.pathItems;
    var DieLineList = [];
    for (var i = 0; i < items.length; i++) {
        var p = items[i];
        if (p.stroked && p.strokeColor.typename == "SpotColor" && esIndexOf(p.strokeColor.spot.name) != -1) {
            //如果刀线上面有混合路径就去除混合路径
            if (p.parent.typename == "CompoundPathItem") {
                while (currentDoc.selection.length > 0) {
                    currentDoc.selection[0].selected = false;
                }
                p.selected = true;
                app.executeMenuCommand('noCompoundPath');
            }
            var originalLayer = p.layer;
            //移动到图层下
            p.move(originalLayer, ElementPlacement.PLACEATBEGINNING);
            DieLineList.push(p);
        }
    }
    //2.将刀线编组
    var group = [];
    for (var i = 0; i < DieLineList.length; i++) {
        var dieLine = DieLineList[i];
        if (dieLine.parent.name == "dieLine") {
            continue;
        }
        var newGroup = currentDoc.groupItems.add();
        dieLine.move(newGroup, ElementPlacement.PLACEATBEGINNING); 
        newGroup.name = "dieLine";
        var bounds = findBounds(dieLine);
        for (var j = 0; j < DieLineList.length; j++) {
            if (j != i) {
                var dieLine2 = DieLineList[j];
                var bounds2 = findBounds(dieLine2);// minX ,minY ,maxX ,maxY
                if (bounds[0] <= bounds2[0] && bounds[1] <= bounds2[1] && bounds[2] >= bounds2[2] && bounds[3] >= bounds2[3]) {
                    //刀线被完全包含的情况
                    dieLine2.move(newGroup, ElementPlacement.PLACEATBEGINNING);
                    dieLine2.name = "2";
                } else if (bounds[0] <= bounds2[0] && bounds[2] >= bounds2[2] && bounds2[0] == bounds2[2]) {
                    //一条线刀线竖着情况
                    if ((bounds[1] <= bounds2[1] && bounds[3] >= bounds2[1]) || (bounds[1] <= bounds2[3] && bounds[3] >= bounds2[3])) {
                        //1.有一端包含在刀线中
                        dieLine2.move(newGroup, ElementPlacement.PLACEATBEGINNING);
                        dieLine2.name = "2";
                    } else if (bounds[1] >= bounds2[1] && bounds[3] <= bounds2[3]) {
                        //2.跨最外层刀线
                        dieLine2.move(newGroup, ElementPlacement.PLACEATBEGINNING);
                        dieLine2.name = "2";
                    }
                } else if (bounds[1] <= bounds2[1] && bounds[3] >= bounds2[3] && bounds2[1] == bounds2[3]) {
                    //一条线刀线横着情况
                    if ((bounds[0] <= bounds2[0] && bounds[2] >= bounds2[0]) || (bounds[0] <= bounds2[2] && bounds[2] >= bounds2[2])) {
                        //1.有一端包含在刀线中
                        dieLine2.move(newGroup, ElementPlacement.PLACEATBEGINNING);
                        dieLine2.name = "2";
                    } else if (bounds[0] >= bounds2[0] && bounds[2] <= bounds2[2]) {
                        //2.跨最外层刀线
                        dieLine2.move(newGroup, ElementPlacement.PLACEATBEGINNING);
                        dieLine2.name = "2";
                    }
                }       
            }
        }
        group.push(newGroup);
    }

    //3.遍历所所有刀线的编组
    var currentLineCount = 0;
    for (var i = 0; i < group.length; i++) {
        var p = group[i];
        var art = pathInArt(p);
        if (art) {
            var currentArtboard = currentCutMap[art];
            var mmL = currentCutMap[art].mmL;
            var mmS = currentCutMap[art].mmS;
            var mmCx = currentCutMap[art].mmCx;
            var pBounds = findDieLinesGroupBounds(p);
            var pw = Math.abs(pBounds[2] - pBounds[0]);
            var ph = Math.abs(pBounds[3] - pBounds[1]);
            // 查看这个图像的边和要求的差距是否小于2pt 
            var isSameFlip1 = Math.abs(pw - mm2pt(mmL, false)) < 2 && Math.abs(ph - mm2pt(mmS, false)) < 2;
            var isSameFlip2 = Math.abs(pw - mm2pt(mmS, false)) < 2 && Math.abs(ph - mm2pt(mmL, false)) < 2;
            if (isSameFlip1 || isSameFlip2) {
                currentLineCount++;
                p.name = 'mqx_' + currentLineCount;
                currentCutMap[art]['lines'][p.name] = {
                        index: currentLineCount,
                        width: pw,
                        height: ph,
                        flip: isSameFlip2,
                        itemsToGroup: [],
                        itemsIndexToGroup: [],
                        bounds: {
                            minX: pBounds[0] - mm2pt(mmCx, false),
                            minY: pBounds[1] - mm2pt(mmCx, false),
                            maxX: pBounds[2] + mm2pt(mmCx, false),
                            maxY: pBounds[3] + mm2pt(mmCx, false),
                        }
                    };
            }
        }
    }
    var kindSum = 0;
    for (var art in currentCutMap) {
        kindSum += currentCutMap[art].kinds;
    }
    var strMsg = "符合尺寸刀线数:" + currentLineCount + ",款数要求：" + kindSum;
    if (currentLineCount != kindSum) {
        throw new Error("款数要求与刀线要求不符合，符合尺寸刀线数:" + currentLineCount + ",款数要求：" + kindSum);
    }
    $.writeln(strMsg);
    //转曲使用副本数组进行处理，实时访问文本类的数据时候length的长度会实时更新所以采用副本处理
    // var textFramesArray = [];
    // for (var k = 0; k < currentDoc.textFrames.length; k++) {
    //     textFramesArray.push(currentDoc.textFrames[k]);
    // }
    // for (var k = 0; k < textFramesArray.length; k++) {
    //     textFramesArray[k].createOutline();
    // }
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

// 根据元素所属的内容编组
// function renameAllItems() {
//     if (!currentDoc) {
//         return "1;未找到文档;未找到打开的文档"
//     }
//     function judgeInCutMap(p, lines) {
//         for (var name in lines) {
//             if (p.name == name) return;
//             var isLeft = Math.round(ib[0] - lines[name].bounds.minX);
//             if (isLeft < -1) continue;
//             var isTop = Math.round(ib[1] - lines[name].bounds.minY);
//             if (isTop < -1) continue;
//             var isRight = Math.round(lines[name].bounds.maxX - ib[2]);
//             if (isRight < -1) continue;
//             var isBottom = Math.round(lines[name].bounds.maxY - ib[3]);
//             if (isBottom < -1) continue;
//             if (p.name.indexOf("_g_") !== -1) return;
//             p.name = name + "_g_" + (lines[name].itemsToGroup.length);
//             lines[name].itemsToGroup.push(p.name);
//             return
//         }
//     }
//     var totalCount = currentDoc.pageItems.length;
//     for (var i = 0; i < totalCount; i++) {
//         var p = currentDoc.pageItems[i];
//         if (p.parent.typename == "Layer") {
//             var ib = findBounds(p);
//             for (var art in currentCutMap) {
//                 judgeInCutMap(p, currentCutMap[art].lines);
//             } 
//         }
//     }
// }

//通过名称判断这个编组是不是刀线的编组
function isValidMqxName(name) {
    var pattern = new RegExp("^mqx_\\d+$");
    return pattern.test(name);
}

//判断祖上是否已经分类
function parentIsGItme(item) {
    while(item.parent.typename !== "Layer") {
        if (item.parent.name.indexOf("_g_") !== -1) return true;
        item = item.parent;
    }
    return false;
}

function renameAllItems() {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档"
    }

    function ItemsBounds(item) {
        var minX = Infinity;
        var minY = Infinity;
        var maxX = -Infinity;
        var maxY = -Infinity;

        function processItem(item) {
            if (item.typename == "TextFrame") {
                var textBounds = item.geometricBounds;
                var textCenterX = (textBounds[0] + textBounds[2]) / 2;
                var textCenterY = (textBounds[1] + textBounds[3]) / 2;
                minX = Math.min(minX, textCenterX);
                minY = Math.min(minY, textCenterY);
                maxX = Math.max(maxX, textCenterX);
                maxY = Math.max(maxY, textCenterY);
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
        //排除刀线 && 排除刀线的编组 &&  排除剪切组中的元素 && 排除祖上是否已经归属
        if (!(p.stroked && p.strokeColor.typename == "SpotColor" && esIndexOf(p.strokeColor.spot.name) != -1) && !isValidMqxName(p.name) 
        && !p.parent.clipped && !parentIsGItme(p)
        ) {
            var ib = ItemsBounds(p);
            for (var art in currentCutMap) {
                judgeInCutMap(p, currentCutMap[art].lines);
            }
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
            //TODO 新建的编组需要移动到原来子元素的父元素下
            var newGroup = currentDoc.groupItems.add();
            //按理来说最终要编组的内容的父类是同一个
            var originalParent;
            for (var i = cutData.itemsToGroup.length - 1; i >= 0; i--) {
                var p = currentDoc.pageItems.getByName(cutData.itemsToGroup[i]);
                //排除祖上已经归属的元素
                if (parentIsGItme(p)) {
                    continue;
                }
                originalParent = p.parent;
                p.move(newGroup, ElementPlacement.PLACEATBEGINNING);
            }
            var currentSelection = currentDoc.selection;
            $.writeln("重命名编组时候选中的元素数量：" + currentSelection.length);
            newGroup.name = "group_" + line;
            newGroup.move(originalParent, ElementPlacement.PLACEATBEGINNING);
            app.executeMenuCommand("deselectall");
        }
    }
}

function moveLineToNewArt(data) {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档"
    }
    for (var art in currentCutMap) {
        for (var line in currentCutMap[art]['lines']) {
            var mqx = currentDoc.groupItems.getByName(line);
            var pos = currentDoc.artboards.getByName("art_2_" + line).artboardRect;
            var bounds = findBounds(mqx);
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
            // group.translate(globX - curX, globY - curY);
            // gorup.position = [globX, globY];
            currentDoc.selection = [];
            selectedItmes(group);
            // 计算移动偏移量，使编组中心与画板中心对齐
            var deltaX = (globX - curX).toFixed(10);
            var deltaY = (globY - curY).toFixed(10);
            $.writeln("移动偏移量：" + deltaX + ";" + deltaY);
            $.writeln("移动偏移量：" + group.name);
            doMoveItem(deltaX, deltaY);
            currentDoc.selection = [];
        }
    }
}

function TextItmesCreateOutline() {
    if (!currentDoc) {
        return "1;未找到文档;未找到打开的文档"
    }
    var textItems = currentDoc.textFrames;
    for (var i = textItems.length - 1; i >= 0; i--) {
        var text = textItems[i];
        text.createOutline();
    }
}

function main(data) {
    try {
        var args = parseRunArgs(data);
        var cx = args.cx;
        var filePath = args.filePath;
        if (!filePath) {
            return "500;ERR;未提供AI文件路径";
        }

        startForDoc(filePath);
        if (!currentDoc) {
            return "1;未找到文档;未找到打开的文档"
        }
        var sum = 0;
        var file = normalizeWinPath(filePath);
        var result = file.split("\\");
        var endPath = [];
        for (var i = 0; i < result.length; i++) {
            if (i == (result.length - 1)) {
                continue;
            }
            endPath.push(result[i]);
        }
        var downPath = args.outDir ? normalizeWinPath(args.outDir) : endPath.join("\\");
        var artboardsLength = currentDoc.artboards.length;
        if (artboardsLength == 1) {
            currentArtboard = currentDoc.artboards[0];
            var filename = result[result.length - 1]
            var failName = filename.substring(0, filename.lastIndexOf("."));
            var dimensions = failName.split("^")[6].split("x");
            currentCutMap[failName] = {
                index: i,
                mmL: +dimensions[0],
                mmS: +dimensions[1],
                mmCx: cx,
                kinds: +failName.split("^")[2],
                rect: currentArtboard.artboardRect,
                lines: {}
            };
            sum = +failName.split("^")[2] * 2;
        } else {
            for (var i = 0; i < artboardsLength; i++) {
                currentArtboard = currentDoc.artboards[i];
                var dimensions = currentArtboard.name.split("^")[6].split("x");
                currentCutMap[currentArtboard.name] = {
                    index: i,
                    mmL: +dimensions[0],
                    mmS: +dimensions[1],
                    mmCx: cx,
                    kinds: +currentArtboard.name.split("^")[2],
                    // kinds: 1,
                    rect: currentArtboard.artboardRect,
                    lines: {}
                };
                sum += +currentArtboard.name.split("^")[2] * 2;
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
        deleteAllActions();
        findAllDieLines();
        getArtLines();
        addNewArt(cx);
        renameAllItems();

        groupAllItemByLine();
        moveLineToNewArt(cx); //复制刀线到画板中
        moveAllItemByLine();
        TextItmesCreateOutline();

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
            var kinds = art.split("^")[2];
            var newFile = normalizeFilePath(downPath + "\\" + safeOutputName(art) + ".pdf");
            var saveFile = new File(decodeURI(newFile));
            try {
                if (saveFile.parent && !saveFile.parent.exists) saveFile.parent.create();
            } catch (eCreate) {}
            try {
                if (saveFile.exists) saveFile.remove();
            } catch (eRemove) {}
            var pdfSaveOpts = new PDFSaveOptions();
            pdfSaveOpts.compatibility = PDFCompatibility.ACROBAT6;
            pdfSaveOpts.acrobatLayers = true; //保留图层
            pdfSaveOpts.viewAfterSaving = false; //不打开保存的pdf文件
            pdfSaveOpts.saveMultipleArtboards = false; //只保留选中的艺术板
            pdfSaveOpts.cropToArtboard = true; // 根据艺术板裁剪页面
            pdfSaveOpts.preserveEditability = false; //是否保留AI可编辑性
            pdfSaveOpts.generateThumbnails = true;
            pdfSaveOpts.embedFont = true; //嵌入字体
            pdfSaveOpts.embedImages = true; //嵌入图片
            pdfSaveOpts.artboardRange = end + "-" + (kinds * 2 + end - 1);
            end = end + kinds * 2;
            currentDoc.saveAs(saveFile, pdfSaveOpts);
        }

        return "200;OK";
    } catch (error) {
        return "500;ERR;" + error + ((error && error.line) ? (" @line=" + error.line) : "");
    } finally {
        try {
            if (currentDoc) currentDoc.close(SaveOptions.DONOTSAVECHANGES);
        } catch (eClose) {}
        currentDoc = null;
    }
}

// main("2;D:\\我的数据\\桌面\\原稿路径\\温州起印包装有限公司^恋味居^1^拼版^80克铜版纸不干胶^印刷,不覆膜,模切成型^90x150^3340^铜版纸不干胶^3291842571536840087^SJ2026032508XM.ai");
// main("2;D:\\我的数据\\桌面\\原稿路径\\太阳人旗舰店^yflulily^1^打样^80克铜版纸不干胶^印刷,覆膜,模切^60x50^60^铜版纸不干胶^5049684613109886332^SJ20260202257E(覆哑膜10枚一张 不同款分开 一个品种排10个一张.ai");
