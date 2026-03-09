
    var srcPath = "D:\\我的数据\\桌面\\原稿路径\\群耀旗舰店^关忆北zy^3^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^182x66^110^乳白PVC不干胶^3260138991677124582^SJ2026022578J3.ai";
    var cols = 17;
    var rows = 1;
    var margin = 30;
    var keepOriginalAtTopLeft = true;
    var names = ["群耀旗舰店^关忆北zy^3^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^182x66^110^乳白PVC不干胶^3260138991677124582^SJ2026022578J3", "群耀旗舰店^关忆北zy^2^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^170x65^210^乳白PVC不干胶^3260138991677124582^SJ2026022578J1", "群耀旗舰店^关忆北zy^2^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^170x65^110^乳白PVC不干胶^3260138991677124582^SJ2026022578HZ", "群耀旗舰店^关忆北zy^2^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^170x65^60^乳白PVC不干胶^3260138991677124582^SJ2026022578HX", "群耀旗舰店^关忆北zy^1^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^182x66^310^乳白PVC不干胶^3260138991677124582^SJ2026022578HV", "群耀旗舰店^关忆北zy^2^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^182x66^210^乳白PVC不干胶^3260138991677124582^SJ2026022578HT", "群耀旗舰店^关忆北zy^2^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^182x66^60^乳白PVC不干胶^3260138991677124582^SJ2026022578HR", "群耀旗舰店^关忆北zy^1^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^145x70^510^乳白PVC不干胶^3260138991677124582^SJ2026022578HM", "群耀旗舰店^关忆北zy^1^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^145x70^310^乳白PVC不干胶^3260138991677124582^SJ2026022578HL", "群耀旗舰店^关忆北zy^3^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^145x70^210^乳白PVC不干胶^3260138991677124582^SJ2026022578HJ", "群耀旗舰店^关忆北zy^6^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^145x70^110^乳白PVC不干胶^3260138991677124582^SJ2026022578HA", "群耀旗舰店^关忆北zy^3^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^145x70^60^乳白PVC不干胶^3260138991677124582^SJ2026022578H6", "群耀旗舰店^关忆北zy^1^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^205x75^410^乳白PVC不干胶^3260138991677124582^SJ2026022578H4", "群耀旗舰店^关忆北zy^2^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^205x75^310^乳白PVC不干胶^3260138991677124582^SJ2026022578H2", "群耀旗舰店^关忆北zy^1^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^205x75^210^乳白PVC不干胶^3260138991677124582^SJ2026022578H0", "群耀旗舰店^关忆北zy^2^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^205x75^110^乳白PVC不干胶^3260138991677124582^SJ2026022578GY", "群耀旗舰店^关忆北zy^1^拼版^8丝乳白PVC不干胶^印刷,覆膜,模切^205x75^60^乳白PVC不干胶^3260138991677124582^SJ2026022578GW"]; // 来自 Python 的名称列表

    function main() {
        var f = new File(srcPath);
        if (!f.exists) {
            alert("源文件不存在: " + srcPath);
            return;
        }
        app.open(f);
        var doc = app.activeDocument;
        doc.documentColorSpace = DocumentColorSpace.CMYK
        if (doc.artboards.length < 1) {
            alert("文档中没有画板。");
            doc.close(SaveOptions.DONOTSAVECHANGES);
            return;
        }

        var refArt = doc.artboards[0];
        var rect = refArt.artboardRect;
        var width = rect[2] - rect[0];
        var height = rect[1] - rect[3];
        var startLeft = rect[0];
        var startTop = rect[1];

        var created = 0;
        var totalBefore = doc.artboards.length;
        var nameIndex = 0;

        // Illustrator坐标限制（大约±16384点）
        var MAX_COORDINATE = 8000;
        var MIN_COORDINATE = -8000;
        for (var r = 0; r < rows; r++) {
            for (var c = 0; c < cols; c++) {
                if (keepOriginalAtTopLeft && r === 0 && c === 0) {
                    // 改名第一个原始画板
                    if (nameIndex < names.length) {
                        doc.artboards[0].name = names[nameIndex++];
                    }
                    continue;
                }
                var left = startLeft + c * (width + margin);
                var top = startTop - r * (height + margin);
                if (left > MAX_COORDINATE || left < MIN_COORDINATE || top > MAX_COORDINATE || top < MIN_COORDINATE 
                || left + width < MIN_COORDINATE || left + width > MAX_COORDINATE 
                || top - height < MIN_COORDINATE || top - height > MAX_COORDINATE) {
                    startLeft = MIN_COORDINATE;
                    startTop = startTop + height + margin;
                    var left = startLeft + c * (width + margin);
                    var top = startTop - r * (height + margin);
                }
                var newRect = [left, top, left + width, top - height];
                var newArtboard = doc.artboards.add(newRect);
                if (nameIndex < names.length) {
                    newArtboard.name = names[nameIndex++];
                } else {
                    newArtboard.name = "Artboard_" + (totalBefore + created + 1);
                }
                created++;
            }
        }

        try { doc.save(); } catch(e) { $.writeln("保存出错：" + e); }
        //alert("完成：新增 " + created + " 个画板。总数：" + doc.artboards.length);
    }
    main();
    