// =======================
// AItest_ai.jsx (FINAL++ ultra robust for multi-shapes)
// =======================

var art_offset = 250;
var currentDoc = null;
var currentCutMap = {};
var currentArtboard = {};
var end = 1;

if (typeof offset === "undefined") var offset = 0;
if (typeof max_right === "undefined") var max_right = 0;

// ✅ 严格模式：true=刀线数不足直接报错；false=尽量继续跑
var STRICT_MATCH = false;

// ✅ 画板归属 padding（mm）
var ART_PAD_MM = 20;

// ✅ 尺寸容差（mm）
var SIZE_TOL_MM = 6;

// ✅ 统一输出目录
var OUT_DIR = "D:\\test_data\\src";

function ensureIllustratorEnv() {
    if (typeof app === "undefined" || !app || typeof File === "undefined" || typeof UnitValue === "undefined") {
        throw new Error("此脚本是 Adobe Illustrator 的 ExtendScript(JSX)，必须在 Illustrator 内运行或通过 COM DoJavaScript 执行。不要用 node 直接执行。");
    }
}

function mm2pt(mm, needRound) {
    ensureIllustratorEnv();
    var pt = UnitValue(mm, "mm").as("pt");
    return needRound ? Math.round(pt) : pt;
}

function ensureDir(pathStr) {
    var f = new Folder(pathStr);
    if (!f.exists) f.create();
}

function safeFileName(name) {
    return String(name).replace(/[\\\/:\*\?"<>\|]/g, "_");
}

// ✅ getByName 安全版：找不到返回 null（不再炸 No such element）
function getByNameSafe(collection, name) {
    try {
        return collection.getByName(name);
    } catch (e) {
        return null;
    }
}

// -----------------------
// 颜色优先级：只是加分，不硬依赖
// 0 = 典型刀线；1 = 绿/Spot；2 = 其它
// -----------------------
function dielinePriority(sc) {
    if (!sc) return 2;

    if (sc.typename === "SpotColor" && sc.spot && sc.spot.name) {
        var n = ("" + sc.spot.name).toUpperCase();

        if (n.indexOf("CUT") !== -1) return 0;
        if (n.indexOf("DIE") !== -1) return 0;
        if (n.indexOf("DIELINE") !== -1) return 0;
        if (n.indexOf("CUTCONTOUR") !== -1) return 0;
        if (n.indexOf("CONTOUR") !== -1) return 0;
        if (n.indexOf("刀") !== -1) return 0;
        if (n.indexOf("模切") !== -1) return 0;

        if (n.indexOf("PANTONE 802") !== -1) return 0;
        if (n.indexOf("802") !== -1) return 0;

        return 1;
    }

    if (sc.typename === "RGBColor") {
        if (sc.green >= 200 && sc.red <= 80 && sc.blue <= 80) return 1;
        return 2;
    }

    if (sc.typename === "CMYKColor") {
        if (sc.magenta <= 15 && sc.cyan >= 40 && sc.yellow >= 40) return 1;
        return 2;
    }
    return 2;
}

function isGoodPathCandidate(p) {
    if (!p) return false;
    if (!p.stroked) return false;
    if (p.guides) return false;
    // 过滤太小线段
    if (p.width < mm2pt(5, false) || p.height < mm2pt(5, false)) return false;
    return true;
}

// bounds: [L, T, R, B]
function getArtByCenter(bounds) {
    var cx = (bounds[0] + bounds[2]) / 2;
    var cy = (bounds[1] + bounds[3]) / 2;
    var pad = mm2pt(ART_PAD_MM, false);

    for (var art in currentCutMap) {
        var rect = currentCutMap[art].rect;
        if (cx >= (rect[0] - pad) && cx <= (rect[2] + pad) && cy <= (rect[1] + pad) && cy >= (rect[3] - pad)) {
            return art;
        }
    }
    return null;
}

function sizeErrorPt(pw, ph, mmW, mmH) {
    var ew = mm2pt(mmW, false);
    var eh = mm2pt(mmH, false);
    return Math.abs(pw - ew) + Math.abs(ph - eh);
}

function bestMatchForArt(pw, ph, mmL, mmS, mmCx) {
    var t1w = mmL, t1h = mmS;
    var t2w = mmL + 2 * mmCx, t2h = mmS + 2 * mmCx;

    var e1 = sizeErrorPt(pw, ph, t1w, t1h);
    var e2 = sizeErrorPt(pw, ph, t1h, t1w);
    var e3 = sizeErrorPt(pw, ph, t2w, t2h);
    var e4 = sizeErrorPt(pw, ph, t2h, t2w);

    var best = { err: e1, flip: false, mode: "base" };
    if (e2 < best.err) best = { err: e2, flip: true, mode: "base" };
    if (e3 < best.err) best = { err: e3, flip: false, mode: "expand" };
    if (e4 < best.err) best = { err: e4, flip: true, mode: "expand" };
    return best;
}

// -----------------------
// 动作：描边对齐居中（保留你原逻辑）
// -----------------------
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
    actionString = null;
    app.doScript("设置对齐描边", "JSX导入动作", false);
    app.unloadAction("JSX导入动作", "");
}

function findAllDieLines() {
    if (!currentDoc) return "500;findAllDieLine未识别到文档";

    var items = currentDoc.pathItems;
    app.executeMenuCommand("deselectall");
    while (currentDoc.selection.length > 0) currentDoc.selection[0].selected = false;

    for (var j = 0; j < items.length; j++) {
        var p = items[j];
        if (!isGoodPathCandidate(p)) continue;

        var pr = dielinePriority(p.strokeColor);
        if (pr <= 1) p.selected = true;
    }

    doAction();
    return "200;刀线对齐描边居中设置完成。";
}

function startForDoc(path) {
    currentDoc = null;
    currentCutMap = {};
    var fileO = new File(decodeURI(path));
    currentDoc = app.open(fileO);
    return "200;文件打开成功";
}

function findBounds(item) {
    var minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;

    function processItem(it) {
        if (it.typename == "TextFrame") {
            return processItem(it.createOutline());
        } else if (("pageItems" in it) == false) {
            var b = it.geometricBounds;
            minX = Math.min(minX, b[0]);
            minY = Math.min(minY, b[3]);
            maxX = Math.max(maxX, b[2]);
            maxY = Math.max(maxY, b[1]);
        } else if (!it.clipped) {
            var items = it.pageItems;
            for (var i = 0; i < items.length; i++) processItem(items[i]);
        } else {
            var b2 = it.pageItems[0].geometricBounds;
            minX = Math.min(minX, b2[0]);
            minY = Math.min(minY, b2[3]);
            maxX = Math.max(maxX, b2[2]);
            maxY = Math.max(maxY, b2[1]);
        }
    }
    processItem(item);
    return [minX, minY, maxX, maxY];
}

// -----------------------
// ✅ 核心：更通用的刀线提取 + 多图形不崩
// - 尺寸为 0x0 时：按 priority + 面积 选
// - 尺寸正常时：按 priority + 尺寸误差 选
// - 刀线不足：按实际条数继续跑（默认不 throw）
// -----------------------
function getArtLines() {
    if (!currentDoc) return "1;未找到文档;未找到打开的文档";

    var items = currentDoc.pathItems;
    var tolPt = mm2pt(SIZE_TOL_MM, false);

    var candMap = {};
    for (var art in currentCutMap) candMap[art] = [];

    for (var j = 0; j < items.length; j++) {
        var p = items[j];
        if (!isGoodPathCandidate(p)) continue;

        var artKey = getArtByCenter(p.geometricBounds);
        if (!artKey) continue;

        var mmL = currentCutMap[artKey].mmL;
        var mmS = currentCutMap[artKey].mmS;
        var mmCx = currentCutMap[artKey].mmCx;

        var pw = p.width;
        var ph = p.height;
        var pr = dielinePriority(p.strokeColor);

        var sizeUnknown = (!mmL || !mmS || mmL <= 0 || mmS <= 0);

        if (sizeUnknown) {
            // ✅ 尺寸未知：用优先级 + 面积（更像外框）
            var area = pw * ph;
            var scoreU = pr * 1e15 - area; // score 越小越优
            candMap[artKey].push({
                p: p,
                score: scoreU,
                flip: false,
                err: 0,
                pr: pr,
                mode: "unknown",
                area: area
            });
        } else {
            // ✅ 尺寸已知：用优先级 + 尺寸误差
            var best = bestMatchForArt(pw, ph, mmL, mmS, mmCx);
            var score = pr * 1e9 + best.err;
            candMap[artKey].push({
                p: p,
                score: score,
                flip: best.flip,
                err: best.err,
                pr: pr,
                mode: best.mode
            });
        }
    }

    var globalIdx = 0;
    var totalPicked = 0;
    var totalNeedByName = 0;
    var forcedFallback = false;

    for (var art2 in currentCutMap) {
        var needByName = currentCutMap[art2].kinds;
        totalNeedByName += needByName;

        var mmL2 = currentCutMap[art2].mmL;
        var mmS2 = currentCutMap[art2].mmS;
        var sizeUnknown2 = (!mmL2 || !mmS2 || mmL2 <= 0 || mmS2 <= 0);

        var list = candMap[art2] || [];
        list.sort(function(a, b) { return a.score - b.score; });

        // 实际能选的最多数量
        var maxCanPick = list.length;
        var need = needByName;

        // 如果候选不足，按实际数量继续
        if (need > maxCanPick) {
            forcedFallback = true;
            need = maxCanPick;
        }

        var chosen = [];

        if (sizeUnknown2) {
            // 尺寸未知：直接按 score 取前 need 个
            for (var u = 0; u < list.length && chosen.length < need; u++) chosen.push(list[u]);
        } else {
            // 尺寸已知：优先拿 err<=tol 的，不够再按最接近补齐
            for (var i = 0; i < list.length && chosen.length < need; i++) {
                if (list[i].err <= tolPt) chosen.push(list[i]);
            }
            if (chosen.length < need) {
                forcedFallback = true;
                for (var k = 0; k < list.length && chosen.length < need; k++) {
                    var already = false;
                    for (var t = 0; t < chosen.length; t++) {
                        if (chosen[t].p === list[k].p) { already = true; break; }
                    }
                    if (!already) chosen.push(list[k]);
                }
            }
        }

        $.writeln("DEBUG art=" + art2 + " needByName=" + needByName + " pick=" + chosen.length + " cand=" + list.length);

        if (chosen.length < needByName && STRICT_MATCH) {
            throw new Error("500;刀线不符;刀线数不足, need=" + needByName + " pick=" + chosen.length);
        }

        // ✅ 写入 lines（只写入 chosen）
        for (var c = 0; c < chosen.length; c++) {
            globalIdx++;
            totalPicked++;

            var pp = chosen[c].p;

            if (pp.parent && pp.parent.typename == "CompoundPathItem") {
                try {
                    while (currentDoc.selection.length > 0) currentDoc.selection[0].selected = false;
                    pp.selected = true;
                    app.executeMenuCommand('noCompoundPath');
                } catch (e) {}
            }

            try {
                pp.name = "mqx_" + globalIdx;
            } catch (e2) {
                continue;
            }

            var bounds = pp.geometricBounds;
            var mmCx2 = currentCutMap[art2].mmCx;

            currentCutMap[art2].lines[pp.name] = {
                index: globalIdx,
                width: pp.width,
                height: pp.height,
                flip: chosen[c].flip,
                itemsToGroup: [],
                bounds: {
                    minX: bounds[0] - mm2pt(mmCx2, false),
                    minY: bounds[3] - mm2pt(mmCx2, false),
                    maxX: bounds[2] + mm2pt(mmCx2, false),
                    maxY: bounds[1] + mm2pt(mmCx2, false)
                }
            };
        }
    }

    $.writeln("DEBUG totalPicked=" + totalPicked + " totalNeedByName=" + totalNeedByName);

    // 转曲（尽量不炸）
    var tf = [];
    for (var x = 0; x < currentDoc.textFrames.length; x++) tf.push(currentDoc.textFrames[x]);
    for (var y = 0; y < tf.length; y++) { try { tf[y].createOutline(); } catch(e3) {} }

    if (forcedFallback) return "206;刀线不足或尺寸不可信，已按最优候选继续处理";
    return "200;刀线识别完成";
}

function addNewArt(cx) {
    if (!currentDoc) return "1;未找到文档;未找到打开的文档";

    var x0 = -2000, y0 = 2000;
    var max_height = 0, total_count = 0;

    for (var art in currentCutMap) {
        for (var line in currentCutMap[art].lines) {
            var h = currentCutMap[art].lines[line].height;
            if (max_height < h) max_height = h;
            total_count++;
        }
    }

    if (total_count <= 0) return "207;未创建新画板：未识别到任何刀线";

    var max_loop = Math.ceil((max_height + mm2pt(2 * cx, false)) * total_count / 4000);
    if (max_loop <= 0) max_loop = 1;
    var each_loop = Math.ceil(total_count / max_loop);
    if (!each_loop || each_loop <= 0) each_loop = 1;

    var i = 0;
    var x1 = 0;

    for (var art2 in currentCutMap) {
        for (var line2 in currentCutMap[art2].lines) {
            var _w = currentCutMap[art2].lines[line2].width + mm2pt(cx * 2, false);
            var _h = currentCutMap[art2].lines[line2].height + mm2pt(cx * 2, false);

            if (i % each_loop == 0) {
                x0 = x1 + _w + art_offset;
                y0 = 2000;
            }
            i++;

            // ✅ 垂直排列：先创建“上面的原图画板”，再创建“下面的轮廓画板”
            // y0 是顶边，y0-_h 是底边
            var topY = y0;
            var bottomY = y0 - _h;

            // 上：印刷(原图) art_1
            var artPrint = currentDoc.artboards.add([x0, topY, x0 + _w, bottomY]);
            artPrint.name = "art_1_" + line2;

            // 下：轮廓(刀线) art_2 —— 直接接着往下放一个同尺寸
            var topY2 = bottomY - art_offset;           // 中间留 art_offset 间距
            var bottomY2 = topY2 - _h;

            var artLine = currentDoc.artboards.add([x0, topY2, x0 + _w, bottomY2]);
            artLine.name = "art_2_" + line2;

            // 下一组继续往下
            y0 = bottomY2 - art_offset;

            // 记录本列最右侧，方便换列
            x1 = x0 + _w;
        }
    }

    return "200;新画板创建完成(上原图/下轮廓)";
}

function renameAllItems() {
    if (!currentDoc) return "1;未找到文档;未找到打开的文档";

    function judgeInCutMap(p, lines, ib) {
        for (var name in lines) {
            if (p.name == name) return;

            var b = lines[name].bounds;
            if (ib[0] < b.minX || ib[2] > b.maxX || ib[3] < b.minY || ib[1] > b.maxY) continue;

            if (p.name && p.name.indexOf("_g_") !== -1) return;

            p.name = name + "_g_" + (lines[name].itemsToGroup.length);
            lines[name].itemsToGroup.push(p.name);
            return;
        }
    }

    var totalCount = currentDoc.pageItems.length;
    for (var i = 0; i < totalCount; i++) {
        var p = currentDoc.pageItems[i];
        if (!p) continue;
        if (p.parent.typename == "GroupItem" || p.parent.typename == "CompoundPathItem") continue;

        var ib = findBounds(p);
        for (var art in currentCutMap) {
            judgeInCutMap(p, currentCutMap[art].lines, ib);
        }
    }
}

function isMqxName(n) { return n && n.indexOf("mqx_") === 0; }

function groupAllItemByLine() {
    if (!currentDoc) return "1;未找到文档;未找到打开的文档";

    for (var art in currentCutMap) {
        for (var line in currentCutMap[art].lines) {
            app.executeMenuCommand("deselectall");
            while (currentDoc.selection.length > 0) currentDoc.selection[0].selected = false;

            var cutData = currentCutMap[art].lines[line];

            for (var i = cutData.itemsToGroup.length - 1; i >= 0; i--) {
                var nm = cutData.itemsToGroup[i];
                var p = getByNameSafe(currentDoc.pageItems, nm);
                if (!p) continue;
                p.selected = true;
            }

            try { app.executeMenuCommand("group"); } catch(e) {}

            var sel = currentDoc.selection;
            for (var s = 0; s < sel.length; s++) {
                if (!sel[s].name) {
                    sel[s].name = "group_" + line;
                    sel[s].selected = false;
                    break;
                }
            }
            app.executeMenuCommand("deselectall");
        }
    }
}

function moveLineToNewArt(cx) {
    if (!currentDoc) return "1;未找到文档;未找到打开的文档";

    for (var art in currentCutMap) {
        for (var line in currentCutMap[art].lines) {
            var mqx = getByNameSafe(currentDoc.pathItems, line);
            var ab2 = getByNameSafe(currentDoc.artboards, "art_2_" + line); // ✅ 轮廓 -> 下方 art_2
            if (!mqx || !ab2) continue;

            var pos = ab2.artboardRect;
            var bounds = mqx.geometricBounds;

            if (currentCutMap[art].lines[line].flip) {
                try { mqx.rotate(90, true, true, true, true, Transformation.CENTER); } catch(e) {}
            }

            var globX = (pos[0] + pos[2]) / 2;
            var globY = (pos[1] + pos[3]) / 2;
            var curX = (bounds[0] + bounds[2]) / 2;
            var curY = (bounds[1] + bounds[3]) / 2;
            try { mqx.translate(globX - curX, globY - curY); } catch(e2) {}
        }
    }
}

function moveAllItemByLine() {
    if (!currentDoc) return "1;未找到文档;未找到打开的文档";

    for (var art in currentCutMap) {
        for (var line in currentCutMap[art].lines) {
            var cutData = currentCutMap[art].lines[line];
            var group = getByNameSafe(currentDoc.groupItems, "group_" + line);
            var ab1 = getByNameSafe(currentDoc.artboards, "art_1_" + line); // ✅ 原图 -> 上方 art_1
            if (!group || !ab1) continue;

            if (cutData.flip) {
                try { group.rotate(90, true, true, true, true, Transformation.CENTER); } catch(e) {}
            }

            var pos = ab1.artboardRect;
            var bounds = findBounds(group);

            var globX = (pos[0] + pos[2]) / 2;
            var globY = (pos[1] + pos[3]) / 2;
            var curX = (bounds[0] + bounds[2]) / 2;
            var curY = (bounds[1] + bounds[3]) / 2;
            try { group.translate(globX - curX, globY - curY); } catch(e2) {}
        }
    }
}

function ungroup(obj) {
    var elements = [];
    var items = obj.pageItems;
    for (var i = 0; i < items.length; i++) elements.push(items[i]);

    if (elements.length < 1) { try { obj.remove(); } catch(e0) {} ; return; }

    for (var j = 0; j < elements.length; j++) {
        try {
            if (elements[j].parent.typename != "Layer") elements[j].moveBefore(obj);
            if (elements[j].typename == "GroupItem" && !elements[j].clipped && elements[j].opacity == 100) {
                ungroup(elements[j]);
            }
        } catch (e) {}
    }
}

function countLines(linesObj) {
    var c = 0;
    for (var k in linesObj) c++;
    return c;
}

function main(data) {
    ensureIllustratorEnv();

    // 支持 COM main(arguments)
    if (data && typeof data !== "string" && data.length && data[0]) data = data[0];

    try {
        var parts = String(data).split(";");
        var cx = +parts[0];
        var filePath = parts[1];

        // ✅ 第3段：OUT_DIR（由 run.py 传入，PDF 输出目录，也就是后续拼图的 PDF 输入目录）
        if (parts.length >= 3 && parts[2]) {
            OUT_DIR = parts[2];
            OUT_DIR = OUT_DIR.replace(/^\s+|\s+$/g, "");
            // 去掉可能的引号
            if ((OUT_DIR.charAt(0) === '"' && OUT_DIR.charAt(OUT_DIR.length - 1) === '"') ||
                (OUT_DIR.charAt(0) === "'" && OUT_DIR.charAt(OUT_DIR.length - 1) === "'")) {
                OUT_DIR = OUT_DIR.substring(1, OUT_DIR.length - 1);
            }
        }

        startForDoc(filePath);
        if (!currentDoc) return "500;未找到文档;未找到打开的文档";

        var artboardsLength = currentDoc.artboards.length;

        if (artboardsLength == 1) {
            currentArtboard = currentDoc.artboards[0];

            var filename = filePath.split("\\").pop();
            var failName = filename.substring(0, filename.lastIndexOf("."));

            var dimsStr = "0x0";
            try { dimsStr = failName.split("^")[6]; } catch(e1) {}
            var dims = (dimsStr && dimsStr.indexOf("x") !== -1) ? dimsStr.split("x") : ["0", "0"];

            currentCutMap[failName] = {
                index: 0,
                mmL: +dims[0] || 0,
                mmS: +dims[1] || 0,
                mmCx: cx,
                kinds: +failName.split("^")[2] || 1,
                rect: currentArtboard.artboardRect,
                lines: {}
            };
        } else {
            for (var a = 0; a < artboardsLength; a++) {
                currentArtboard = currentDoc.artboards[a];
                var nm = currentArtboard.name;

                var dims2 = ["0", "0"];
                try {
                    var ds2 = nm.split("^")[6];
                    dims2 = ds2.split("x");
                } catch(e2) {}

                currentCutMap[nm] = {
                    index: a,
                    mmL: +dims2[0] || 0,
                    mmS: +dims2[1] || 0,
                    mmCx: cx,
                    kinds: +nm.split("^")[2] || 1,
                    rect: currentArtboard.artboardRect,
                    lines: {}
                };
            }
        }

        // 解组
        for (var li = currentDoc.layers.length - 1; li >= 0; li--) {
            if (currentDoc.groupItems.length) ungroup(currentDoc.layers[li]);
        }

        findAllDieLines();
        var r1 = getArtLines();
        var r2 = addNewArt(cx);

        renameAllItems();
        groupAllItemByLine();
        moveLineToNewArt(cx);
        moveAllItemByLine();

        // 删除旧画板
        if (artboardsLength > 1) {
            var remove_ids = [];
            for (var art in currentCutMap) remove_ids.push(currentCutMap[art].index);
            for (var ri = remove_ids.length - 1; ri >= 0; ri--) {
                try { currentDoc.artboards.remove(remove_ids[ri]); } catch(e3) {}
            }
        } else {
            try { currentDoc.artboards.remove(0); } catch(e4) {}
        }

        // 统一输出目录
        ensureDir(OUT_DIR);

        // ✅ 按“实际识别到的刀线条数”导出（避免 kinds 越界）
        for (var art2 in currentCutMap) {
            var found = countLines(currentCutMap[art2].lines);
            if (found <= 0) {
                $.writeln("WARN: no lines for art=" + art2 + " , skip export.");
                continue;
            }

            var outName = safeFileName(art2) + ".pdf";
            var saveFile = new File(OUT_DIR + "\\" + outName);

            var pdfSaveOpts = new PDFSaveOptions();
            pdfSaveOpts.compatibility = PDFCompatibility.ACROBAT6;
            pdfSaveOpts.acrobatLayers = true;
            pdfSaveOpts.viewAfterSaving = false;
            pdfSaveOpts.saveMultipleArtboards = false;
            pdfSaveOpts.cropToArtboard = true;
            pdfSaveOpts.preserveEditability = false;
            pdfSaveOpts.generateThumbnails = true;

            // found * 2 个画板
            pdfSaveOpts.artboardRange = end + "-" + (found * 2 + end - 1);
            end = end + found * 2;

            try { currentDoc.saveAs(saveFile, pdfSaveOpts); } catch(e5) {
                $.writeln("ERR saveAs: " + e5);
            }
        }

        // 关闭源文件，避免 AI 越跑越卡
        try { currentDoc.close(SaveOptions.DONOTSAVECHANGES); } catch(e6) {}
        currentDoc = null;

        return "200;OK;" + (r1 || "") + "|" + (r2 || "");

    } catch (e) {
        try { if (currentDoc) currentDoc.close(SaveOptions.DONOTSAVECHANGES); } catch(e7) {}
        currentDoc = null;
        return "500;ERR;" + e;
    }
}

// 手动测试：
// main("2;D:\\test_data\\src\\太阳人旗舰店^fang_wow^4^打样^...^SJ2601252110.ai;D:\\test_data\\src");