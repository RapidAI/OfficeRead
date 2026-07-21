# officeread 鎬ц兘浼樺寲鎶ュ憡

鏃ユ湡锛?026-06-26

## 鑳屾櫙

鍏煎鎬ф祴璇曚腑鍙戠幇閮ㄥ垎 Office 鏍锋湰瀛樺湪鏄庢樉闀垮熬鑰楁椂锛屼富瑕侀泦涓湪瓒呭ぇ `.xlsx/.xls` 鏂囨湰杈撳嚭銆乣.docx` 鍙鎬у垎鏋愪互鍙婂祵鍏?legacy Office 鏂囨。鎻愬彇銆?
## 鏈疆鏂板浼樺寲

- `.docx` 鍏宠仈鏂囨湰鍙鎬х紦瀛橈細`docxVisibleRelatedTextParts` 澧炲姞浠?`word/document.xml` 涓洪敭鐨勭煭鐢熷懡鍛ㄦ湡缂撳瓨锛岄伩鍏嶆枃鏈彁鍙栭樁娈靛拰 Markdown 闃舵閲嶅鎵弿 charts/drawings/diagrams 鍏崇郴銆?- `.docx` 鍏宠仈鏂囨湰蹇€熷垽鏂細`collectDocxSourceVisibleRelatedTextParts` 鍦ㄥ畬鏁磋В鏋愬浘鐗囧叧绯诲墠鍏堝仛 `likelyImageRelationshipMarkup` 瀛楄妭绾у垽鏂紱娌℃湁鍏崇郴 marker 鐨?XML 鐩存帴璧版櫘閫氬叧绯诲彲杈?fallback銆?- OOXML 鍥剧墖鍏崇郴 marker 鏀剁獎锛氱Щ闄よ繃瀹界殑 `" id"` 棰勫垽椤癸紝淇濈暀 `:id`銆乪mbed/link/relid 绛夌湡姝ｄ細杩涘叆鍏崇郴瑙ｆ瀽鐨勬爣璁帮紝鍑忓皯鏅€?XML 琚鍒ゅ悗杩涘叆瀹屾暣 decoder銆?- OOXML 鍥剧墖 alt 鏄犲皠蹇€熻烦杩囷細`ooxmlVisibleImageAltMap` 鍦ㄨВ鏋?`visibleImageRelationshipAlts` 鍓嶅鐢ㄥ浘鐗囧叧绯?marker 棰勫垽锛岃烦杩囨病鏈夊浘鐗囧叧绯诲睘鎬х殑鍊欓€?XML銆?- OOXML 鍥剧墖 alt 鍒楄〃鎳掑姞杞斤細`extractOOXMLImages` 涓嶅啀鎻愬墠瑙ｆ瀽閫氱敤鍥剧墖 alt 鍒楄〃锛屽彧鏈夊綋鏌愬紶鍥剧墖鏃犳硶閫氳繃濯掍綋鍏崇郴鏄犲皠鎷垮埌 alt 鏃舵墠鎸夐『搴?fallback锛涙棤鍥剧墖鏂囨。鍙洿鎺ョ渷鎺夎璺緞銆?- `.xlsx` simple inlineStr worksheet 蹇矾寰勶細瀵规病鏈夐殣钘忚鍒椼€佸叕寮忋€佸叡浜瓧绗︿覆銆佽秴閾炬帴銆佹暟鎹獙璇併€乭eader/footer 鍜屽鏉傚瘜鏂囨湰鏍囪鐨勭函 `inlineStr` worksheet锛屼娇鐢ㄥ瓧鑺傜骇鎵弿鎻愬彇 `<c t="inlineStr"><is><t>...` 鏂囨湰鍜屽墠 200 琛?Markdown 琛ㄦ牸锛涗笉婊¤冻鏉′欢鏃惰嚜鍔ㄥ洖閫€鍘?XML decoder銆?- `.xlsx` comments 鍙鎬ч噸鎺掞細`xlsxVisibleCommentPartSourcesUncached` 鍏堟鏌?sheet 鍏崇郴涓槸鍚﹀瓨鍦?comments/threadedComments锛屽啀璇诲彇 worksheet 骞惰В鏋愰殣钘忚鍒楋紱娌℃湁 comments 鐨?sheet 涓嶅啀瑙ｅ帇鍜屾壂鎻?worksheet XML銆?- `.xlsx` chart 鍙鎬ч噸鎺掞細`collectXlsxSourceVisibleChartParts` 鍏堟鏌ュ綋鍓?source 鐨勫叧绯荤洰鏍囨槸鍚﹀彲鑳介€氬悜 `xl/charts/` 鎴?`xl/drawings/`锛屾病鏈夌浉鍏冲叧绯绘椂鐩存帴璺宠繃 worksheet XML 瑙ｆ瀽銆?- `.xlsx` worksheet chart 鍏崇郴 refs 瀛楄妭鎵弿锛歚collectXlsxSourceVisibleChartParts` 瀵规櫘閫?`xl/worksheets/*.xml` 浣跨敤 `xlsxWorksheetRelationshipRefsFast` 鐩存帴鎵弿 XML tag 灞炴€ч噷鐨勫叧绯?id锛涢亣鍒?`AlternateContent` 鎴?DOCTYPE 鏃跺洖閫€鍘?XML decoder锛岄伩鍏嶄负甯歌 worksheet 鐨?`<drawing r:id=...>` 璺戝畬鏁?`xlsxImageRelationshipRefs`銆?- `.xlsx` worksheet hidden range 瀛楄妭鎵弿锛歚worksheetHiddenRanges` 瀵?`<col>` / `<row>` 鐨?`hidden`銆乣width`銆乣ht`銆乣min`銆乣max`銆乣r` 灞炴€т娇鐢ㄨ交閲?tag scanner锛屾櫘閫氬悎娉?worksheet 涓嶅啀涓?comments 鍙鎬у垱寤?XML decoder token 娴侊紱鎴柇 tag 浼氬洖閫€鍘?XML decoder銆?- `.xlsx` worksheet 闅愯棌鍗曞厓鏍艰烦杩?CharData锛歚appendWorksheetText` 瀵归殣钘忚鍒椾腑鐨勫崟鍏冩牸涓嶅啀绱Н `<v>/<t>` 瀛楃鏁版嵁锛屽苟寤跺悗 `cur.String()` 鍒扮‘璁ゅ崟鍏冩牸鍙涔嬪悗锛屽噺灏戦殣钘忓崟鍏冩牸璺緞涓婄殑鏃犳晥 builder 鍐欏叆鍜屽瓧绗︿覆鏋勯€犮€?- `.xlsx` 榛樿鏁板€煎崟鍏冩牸杞婚噺杩藉姞锛歚appendWorksheetText` 瀵归粯璁ょ被鍨?`<v>` 涓殑绾暟瀛?绉戝璁℃暟娉曞€艰烦杩囬€氱敤 `cleanText`锛岀洿鎺ヨ拷鍔犳竻鐞嗗悗鐨勬枃鏈潡锛涢潪绾暟鍊笺€佸叡浜瓧绗︿覆銆佸竷灏斿€煎拰 inlineStr 浠嶈蛋鍘熸湁娓呮礂璺緞銆?- `.xlsx` 榛樿鏁板€煎崟鍏冩牸 trim 鍘婚噸锛氶粯璁ゆ暟鍊煎揩璺緞澶嶇敤宸茬粡 `TrimSpace` 鍚庣殑鍊硷紝骞堕€氳繃 `appendTrimmedTextBlock` 杩藉姞锛岄伩鍏?`plainExcelNumberValue` 鍜屾枃鏈潡杩藉姞闃舵閲嶅 trim銆?- `.xlsx` Markdown 鍗曞厓鏍兼敹闆?guard锛歚appendWorksheetText` 澧炲姞 `collectMarkdownCell`锛屽彧鏈夊綋鍓嶈浠嶅湪 Markdown 鍓?200 琛岃寖鍥淬€佸綋鍓嶅垪涓嶈秴杩?50 涓斿崟鍏冩牸鍙鏃舵墠閲嶇疆/鍐欏叆 `markdownCellText`锛岄伩鍏嶅悗缁ぇ琛ㄥ崟鍏冩牸缁х画缁存姢涓嶄細浣跨敤鐨?Markdown 缂撳啿銆?- `.xlsx` 澶у€煎幓閲?map 鎳掑垎閰嶏細`appendWorksheetText` 鍜?simple inlineStr 蹇矾寰勪笉鍐嶉鍏堝垎閰?`seenLargeValues` / `seenLargeSharedIndexes`锛屽彧鏈夐亣鍒拌秴杩?`maxRepeatedTextPartBytes` 鐨勫€兼垨鍏变韩瀛楃涓茬储寮曟椂鎵嶅垱寤?map銆?- 鏂囨湰娓呮礂澶у皬鍐欏垽鏂幓鍒嗛厤锛歚cleanTextFastPath` 鐨?`rid` 妫€鏌ユ敼涓轰笓鐢?ASCII 蹇界暐澶у皬鍐欐壂鎻忥紝`maybeControlFragmentText` 涓殑 OLE/Office 鎺у埗鐗囨鍓嶇紑鍜屽悗缂€鍒ゆ柇鏀逛负 `EqualFold` 鍒囩墖姣旇緝锛岄伩鍏嶄负鐑矾寰勫崟鍏冩牸鏂囨湰鍙嶅鏋勯€?`strings.ToLower` 涓存椂瀛楃涓层€俻prof 涓?`strings.ToLower` 浠庣害 1.55 s cum 闄嶅埌绾?0.66 s cum锛? 涓叧閿牱鏈枃鏈?鍥剧墖璁℃暟淇濇寔涓嶅彉銆?- 浣庝俊鎭墖娈靛垽鏂幓鍒嗛厤锛歚looksLikeLowInformationFragment` 涓嶅啀涓烘瘡涓€欓€夊瓧绗︿覆鏋勯€?`[]rune` 鍜?`map[rune]bool`锛屾敼涓哄崟娆?range 缁熻骞舵渶澶氳窡韪?3 涓笉鍚?rune锛涢噸澶嶅瓧绗︺€佸敮涓€瀛楃鏁般€佹爣鐐规瘮渚嬪拰鏃犲厓闊冲瓧姣嶆瘮渚嬭鍒欎繚鎸佷笉鍙樸€俻prof 涓鍑芥暟浠庣害 1.83 s cum 闄嶅埌绾?0.44 s cum锛? 涓叧閿牱鏈枃鏈?鍥剧墖璁℃暟淇濇寔涓嶅彉銆?- 浜岃繘鍒舵帶鍒剁墖娈靛ぇ灏忓啓鍒ゆ柇鍘诲垎閰嶏細`looksLikeBinaryControlFragment` 涓嶅啀涓€杩涘叆鍑芥暟灏辨瀯閫?`strings.ToLower(trimmed)`锛屾敼鐢?`EqualFold`銆乣hasPrefixFold`銆乣hasSuffixFold` 鍜?ASCII 蹇界暐澶у皬鍐欏瓙涓叉壂鎻忚〃杈惧師鏈夋潯浠讹紱XML 鏍囪鍒ゆ柇鐩存帴妫€鏌?`</` / `/>`銆俻prof 涓鍑芥暟浠庣害 4.18 s cum 闄嶅埌绾?3.31 s cum锛? 涓叧閿牱鏈枃鏈?鍥剧墖璁℃暟淇濇寔涓嶅彉銆?- OLE class 鐗囨鍒ゆ柇鍘诲垎閰嶏細`looksLikeOLEClassFragment` 涓嶅啀瀵瑰€欓€夊瓧绗︿覆鏁翠綋 `ToLower` 鍚庡仛 switch/prefix 鍒ゆ柇锛屾敼涓洪瀛楃鍒嗘淳涓嬬殑 `EqualFold` / `hasPrefixFold` / `hasSuffixFold` 绮剧‘鍖归厤锛沗microsoft forms 2.0` 淇濇寔鍘熸湁鎺т欢鍚嶇櫧鍚嶅崟锛岄伩鍏嶆墿澶ц繃婊よ寖鍥淬€傞噰鏍蜂腑璇ュ嚱鏁板凡涓嶅啀浣滀负鐙珛鐑偣鍑虹幇锛? 涓叧閿牱鏈枃鏈?鍥剧墖璁℃暟淇濇寔涓嶅彉銆?- legacy Cyrillic 濉厖鍣０鍒ゆ柇鍘诲垎閰嶏細`looksLikeLegacyRepeatedCyrillicFill` 涓嶅啀涓哄€欓€夊瓧绗︿覆鏋勯€?`[]rune` 鍜?`map[rune]bool`锛屾敼涓哄崟娆?range 缁熻骞剁敤 4 妲芥暟缁勮窡韪笉鍚?rune锛涒€滃叏 Cyrillic銆侀暱搴﹁嚦灏?6銆佷笉鍚屽瓧绗︿笉瓒呰繃 3鈥濈殑瑙勫垯淇濇寔涓嶅彉銆? 涓叧閿牱鏈枃鏈?鍥剧墖璁℃暟淇濇寔涓嶅彉銆?- legacy glyph soup 琛屽垽鏂幓鍒嗛厤锛歚looksLikeLegacyShortGlyphSoupLine` 鍜?`looksLikeLegacyGlyphSoupContinuationLine` 涓嶅啀棰勫厛鏋勯€?`[]rune(strings.TrimSpace(s))`锛屾敼涓哄崟娆?range 缁熻闀垮害涓庡瓧绗﹀垎绫伙紱鍘熸湁闀垮害銆佽剼鏈€佺┖鐧姐€佹暟瀛楀拰 marker 鏉′欢淇濇寔涓嶅彉銆? 涓叧閿牱鏈枃鏈?鍥剧墖璁℃暟淇濇寔涓嶅彉銆?- OOXML 鍥剧墖 alt 瑙ｆ瀽澶嶇敤锛歚ooxmlVisibleImageAltData` 鍦ㄦ瀯寤?media->alt 鏄犲皠鏃堕『鎵嬩繚瀛樺凡鎵弿 part 鐨勫彲瑙?alt 鍒楄〃锛涘綋鍥剧墖闇€瑕佹寜椤哄簭 fallback alt 鏃讹紝澶嶇敤杩欎簺缁撴灉锛岄伩鍏嶅鍚屼竴 DOCX/PPTX/XLSX XML part 浜屾瑙ｆ瀽銆?- OOXML 鍥剧墖鍏崇郴 refs 鎸?part 缂撳瓨锛歚imageRelationshipRefsFromPartBytes` 浠?`*zip.File` 涓洪敭缂撳瓨宸茶В鏋愮殑鍥剧墖鍏崇郴 refs锛屽苟鍦?DOCX 鍏宠仈鏂囨湰銆佸彲瑙佸獟浣撳拰鍥剧墖 alt 璺緞涔嬮棿澶嶇敤锛涘凡缁忚鍏?XML 瀛楄妭鐨勮矾寰勭洿鎺ヨВ鏋愬苟鍐欏叆缂撳瓨锛岄伩鍏嶅悓涓€ part 澶氭瑙ｅ帇鍜?XML decoder 鎵弿銆?- `.xls` BIFF 鏂囨湰 emit 鍘讳复鏃跺垏鐗囷細`biffText` 鐑矾寰勬敼鐢?`forEachBIFFText` 鍥炶皟鐩存帴鍐欏叆 `biffTextPart`锛屼繚鐣欏師鏈夋竻娲?鍣０杩囨护瑙勫垯锛屽悓鏃堕伩鍏嶆瘡涓崟鍏冩牸鍏堟瀯閫犱复鏃?`[]string` 鍐嶄簩娆?append銆?- `.xls` 鍙鍗曞厓鏍煎垽鏂幓閲嶏細`biffText` 鍐呴儴 `addCellText` 涓嶅啀閲嶅鎵ц璋冪敤鐐瑰凡缁忓畬鎴愮殑闅愯棌琛屽垪鍒ゆ柇锛屽噺灏?BIFF 鍗曞厓鏍肩儹璺緞涓婄殑 row/col 瑙ｆ瀽鍜岄殣钘忓垪閬嶅巻銆?- `.xls` 鏅€?OLE 鍘婚噸鐭矾锛歚extractLegacyTextWithMetadata` 鍦ㄦ櫘閫?`.xls`銆佹湭璇锋眰 metadata 涓斿灞傛病鏈夐澶?BoundSheet 鍚嶇О鏃讹紝鐩存帴杩斿洖 workbook BIFF 鏂囨湰锛岄伩鍏嶅澶ф枃鏈粨鏋滃啀娆℃瀯寤哄幓閲?map銆?- `.xls` BoundSheet 鎵弿璺冲€欓€夛細`biffBoundSheetNames` 鏀圭敤 `bytes.IndexByte` 璺冲埌 `0x85` 鍊欓€変綅缃悗鍐嶆墽琛屽師鏈?BIFF 璁板綍鏍￠獙锛屽噺灏戝 OLE 澶栧眰鏁版嵁鐨勯€愬瓧鑺傛棤鏁堟鏌ャ€?- legacy OLE WMF 鎵弿棰勫垽锛歚carveSizedImages` 鍦ㄩ€愬瓧鑺傚皾璇?WMF 瑙ｆ瀽鍓嶅厛妫€鏌?placeable/standard WMF 澶撮儴锛屽噺灏戞棤鍥剧墖 OLE 娴佷笂鐨?`wmfDeclaredSize` 璋冪敤銆?- legacy `.doc` Markdown 澶嶇敤锛歚extractLegacyWithDepth` 鍦ㄦ櫘閫?`.doc` 涓旀湭璇锋眰 metadata 鏃讹紝鐩存帴鐢ㄥ凡鎻愬彇鐨?`textParts` 鐢熸垚 `## Document` Markdown锛岄伩鍏?`extractLegacyMarkdown` 鍐嶆瑙ｆ瀽 WordDocument/piece table锛涘祵鍏?DOCX 涓殑 legacy DOC 瀵硅薄鍙楃泭鏄庢樉銆?- `.xlsx` chart 鍏崇郴瑙ｆ瀽锛氭柊澧?`xlsxImageRelationshipRefs`锛屽 Excel 鍥捐〃/缁樺浘鍙鎬т娇鐢ㄦ洿绐勭殑 `RawToken()` 瑙ｆ瀽鍣紝閬垮厤濂楃敤 Word 淇闅愯棌閫昏緫鍜?namespace 缈昏瘧銆?- `.xlsx` hidden range 蹇€熷垽鏂細`worksheetHiddenRanges` 鍦?XML 瑙ｆ瀽鍓嶅厛妫€鏌ユ槸鍚﹀瓨鍦?`hidden` 鎴?`width/ht` 涓?0 鐨勮抗璞★紱鏅€?worksheet 鐩存帴杩斿洖绌洪殣钘忚寖鍥达紝鍑忓皯 comments 鍙鎬у垎鏋愭垚鏈€?- `.xlsx` comments/chart 鍙鎬х紦瀛橈細`xlsxVisibleCommentPartSources` 鍜?`xlsxVisibleChartPartNames` 澧炲姞鐭敓鍛藉懆鏈熺紦瀛橈紝閬垮厤鏂囨湰闃舵鍜?Markdown 闃舵閲嶅璁＄畻鍚屼竴宸ヤ綔绨跨殑 comments/chart 鍙鎬с€?- `.xlsx` hidden range 鎵弿锛歚worksheetHiddenRanges` 鏀圭敤 `xml.Decoder.RawToken()`锛屽噺灏?comments 鍙鎬у垎鏋愭椂鐨勫懡鍚嶇┖闂寸炕璇戝紑閿€銆?- `.xlsx` chart 鍙鎬э細`collectXlsxSourceVisibleChartParts` 鍦ㄨ繘鍏ュ浘鐗囧叧绯诲彲瑙佹€цВ鏋愬墠鍏堝仛 `likelyImageRelationshipMarkup` 瀛楄妭绾у垽鏂紱娌℃湁鍏崇郴 id/embed/link 鏍囪鏃剁洿鎺ヨ蛋鏅€氬叧绯诲彲杈捐矾寰勩€?- `.xlsx` worksheet 涓绘壂鎻忥細`appendWorksheetText` 鏀圭敤 `xml.Decoder.RawToken()`锛岄伩鍏嶆爣鍑?`Token()` 鐨勫懡鍚嶇┖闂寸炕璇戝紑閿€銆傝璺緞鍙緷璧栧厓绱犳湰鍦板悕銆佸睘鎬у悕鍜屽瓧绗︽暟鎹紝鍥犳涓嶉渶瑕?namespace 瑙ｆ瀽銆?- `.xlsx` shared strings 璇诲彇锛歚readSharedStrings` 鍚屾牱鏀圭敤 `RawToken()`锛岄檷浣庡叡浜瓧绗︿覆鍨嬪伐浣滅翱鐨?XML token 鎴愭湰銆?- 宸查獙璇佷絾鏈繚鐣欙細鏃╂湡鏇惧皾璇曚负 OOXML 鍥剧墖鍏崇郴 refs 澧炲姞閫氱敤缂撳瓨锛屼絾鐩爣鏍锋湰娌℃湁绋冲畾鏀剁泭涓斿鍔?map clone 寮€閿€锛屽洜姝ゅ凡鎾ゅ洖锛涙湰杞繚鐣欑殑鏄洿绐勭殑 `*zip.File` part 绾х紦瀛樸€?- 宸查獙璇佷絾鏈繚鐣欙細鏇惧皾璇曚负 `.rels` 鍏崇郴鏄犲皠澧炲姞 clone 缂撳瓨鍜屽彧璇荤紦瀛橈紱clone 鐗堟湰鍦ㄥ叧閿牱鏈笂鏄庢樉鍙樻參锛屽彧璇荤増鏈敹鐩婁笉绋冲畾涓旂淮鎶ら闄╂洿楂橈紝鍥犳宸叉挙鍥炪€?- 宸查獙璇佷絾鏈繚鐣欙細鏇惧皾璇曚负 worksheet 涓绘壂鎻忓鍔犳暣鏂囦欢瀛楄妭棰勫垽銆侀殣钘忓垪鍒ゆ柇 fastpath锛屼互鍙婁负 DOCX 鍥剧墖鍏崇郴澧炲姞 RawToken guard 蹇矾寰勶紱鍏抽敭鏍锋湰娌℃湁绋冲畾鏀剁泭锛屽洜姝ゅ凡鎾ゅ洖銆?- 宸查獙璇佷絾鏈繚鐣欙細鏇惧皾璇曡 BIFF 鏁板€兼樉绀哄€艰烦杩囦簩杩涘埗鍣０杩囨护锛宍008055.xls` 铏芥槑鏄惧彉蹇絾鏂囨湰瀛楄妭鏁版敼鍙橈紝鍥犳宸叉挙鍥炪€?- 宸查獙璇佷絾鏈繚鐣欙細鏇惧皾璇曚负 BIFF 鏁板瓧鏄剧ず鐗囨澧炲姞鏇村鐨勪簩杩涘埗鍣０杩囨护蹇矾寰勶紝`008055.xls` 鑰楁椂闄嶈嚦绾?1.2 s锛屼絾鏂囨湰瀛楄妭鏁颁粠 `17684747` 鍙樹负 `19558754`锛屽洜姝ゅ凡鎾ゅ洖銆?- 宸查獙璇佷絾鏈繚鐣欙細鏇惧皾璇曚负 `sharedStrings.xml` 澧炲姞 `<si>/<t>` 瀛楄妭鎵弿蹇矾寰勶紝`00012389.xlsx` 鑰楁椂鐣ラ檷浣嗘枃鏈瓧鑺傛暟浠?`11148655` 鍙樹负 `11148682`锛屽洜姝ゅ凡鎾ゅ洖銆?- 绋冲畾鎬т慨澶嶏細`utf16Strings` 鍦ㄤ簩杩涘埗 fallback 涓亣鍒拌繛缁垚瀵?ASCII 瀛楄妭鏃跺厛 flush 宸茶瘑鍒殑 UTF-16 鏂囨湰锛岄伩鍏嶇湡瀹?Unicode 鏂囨湰鍜屽悗缁?ASCII 璧勬簮璺緞绮樿繛鍚庤璇垽涓哄櫔澹般€?
## 宸叉湁浼樺寲姹囨€?
- `.xlsx`锛氬叡浜瓧绗︿覆鎸夌储寮曞幓閲嶏紝Markdown 琛屾暟杈惧埌涓婇檺鍚庡仠姝㈡敹闆嗗崟鍏冩牸鍐呭锛屼富鏂囨湰杈撳嚭浣跨敤 `strings.Builder` 鐩存帴鍐欏叆骞舵寜 worksheet 瑙ｅ帇澶у皬棰勫垎閰嶃€?- `.xlsx`锛氭枃鏈壂鎻忓悓姝ュ噯澶?Markdown 闇€瑕佺殑鍓?200 琛屽拰闄勫姞鏂囨湰锛屽噺灏戝ぇ鍨?worksheet 鐨勯噸澶嶈В鍘嬪拰瑙ｆ瀽銆?- `.xlsx`锛歝harts銆乨rawings銆乼ables銆乧omments銆乸ivots銆乻licers 绛夐檮鍔犻儴浠朵笉瀛樺湪鏃舵彁鍓嶈烦杩囧叧绯诲拰 XML 鍒嗘瀽銆?- `.xlsx`锛氬彧鏈?`hyperlink` 鍜?`dataValidation` 鍙兘璐＄尞 worksheet 鍙灞炴€ф枃鏈紝鍏朵粬 start token 璺宠繃灞炴€ф枃鏈拰鍗曞厓鏍煎紩鐢ㄥ彲瑙佹€у垽鏂€?- `.xls`锛欱IFF 鏂囨湰宸叉竻鐞嗗悗浣跨敤杞婚噺鎷兼帴璺緞锛涙甯?BIFF 鏂囨湰鍜屽灞?`.xls` 鍚堝苟鏂囨湰浣跨敤 `uniqueCleanedStrings` 杞婚噺鍘婚噸锛屽苟鎸夎緭鍏ヨ妯￠鍒嗛厤銆?- `.xls`锛歚MULRK` 绛夎矾寰勫噺灏戜复鏃跺垏鐗囧拰浜屾娓呯悊锛沴egacy/OLE 鎺у埗鏂囨湰鍒ゆ柇鐢ㄦ墜鍐?`isLegacyObjectReference` 鏇夸唬鐑矾寰勬鍒欍€?- `.docx`锛氬浘鐗囧昂瀵告帰娴嬫敼鐢?`image.DecodeConfig`锛涘紩鐢ㄥ垎鏋愬鍔犲瓧鑺傚寘鍚揩閫熷垽鏂紱Word 鍩熸竻鐞嗗鍔?marker 蹇€熷垽鏂紱header/footer 鍙鎬у鍔犵煭鐢熷懡鍛ㄦ湡缂撳瓨銆?
## 澶嶆祴缁撴灉

| 鏍锋湰 | 鏍煎紡 | 浼樺寲鍓?| 涓婅疆缁撴灉 | 鏈疆缁撴灉 | 鎬绘彁鍗?|
|---|---|---:|---:|---:|---:|
| `testRecordSizeExceeded.xlsx` | `.xlsx` | 1012.6 s | 6.4 s | 5.2 s | 绾?196.2x |
| `00012389.xlsx` | `.xlsx` | 70.6 s | 8.6 s | 1.9 s | 绾?37.9x |
| `00014878.xlsx` | `.xlsx` | 60.0 s | 4.3 s | 0.5 s | 绾?114.2x |
| `008055.xls` | `.xls` | 167.4 s | 2.3 s | 2.5 s | 绾?67.0x |
| `00003763.docx` | `.docx` | 110.7 s | 4.7 s | 4.1 s | 绾?26.9x |
| `223624.docx` | `.docx` | 22.8 s | 3.3 s | 2.7 s | 绾?8.4x |
| `016161.xls` | `.xls` | 66.7 s | 1.3 s | 1.3 s | 绾?50.7x |
| `002505.xls` | `.xls` | 42.4 s | 0.9 s | 0.9 s | 绾?47.6x |

鏈疆鐩爣澶嶆祴鏂囦欢锛?
- `testdata/web-samples/reports/perf-current-continue-targets.json`
- `testdata/web-samples/reports/perf-current-continue-targets.csv`
- `testdata/web-samples/reports/perf-after-xlsx-rawtoken-targets.json`
- `testdata/web-samples/reports/perf-after-xlsx-rawtoken-targets.csv`
- `testdata/web-samples/reports/perf-after-xlsx-rawtoken-rerun.json`
- `testdata/web-samples/reports/perf-after-xlsx-rawtoken-rerun.csv`
- `testdata/web-samples/reports/perf-after-xlsx-rawtoken-final.json`
- `testdata/web-samples/reports/perf-after-xlsx-rawtoken-final.csv`
- `testdata/web-samples/reports/perf-after-xlsx-visibility-cache-00012389-rerun.json`
- `testdata/web-samples/reports/perf-after-xlsx-visibility-cache-00012389-rerun.csv`
- `testdata/web-samples/reports/perf-after-xlsx-visibility-cache-raw-hiddenranges-00012389-rerun.json`
- `testdata/web-samples/reports/perf-after-xlsx-visibility-cache-raw-hiddenranges-00012389-rerun.csv`
- `testdata/web-samples/reports/perf-after-xlsx-chart-markup-guard.json`
- `testdata/web-samples/reports/perf-after-xlsx-chart-markup-guard.csv`
- `testdata/web-samples/reports/perf-after-xlsx-cache-guard-huge-rerun.json`
- `testdata/web-samples/reports/perf-after-xlsx-cache-guard-huge-rerun.csv`
- `testdata/web-samples/reports/perf-after-xlsx-image-refs-raw.json`
- `testdata/web-samples/reports/perf-after-xlsx-image-refs-raw.csv`
- `testdata/web-samples/reports/perf-after-worksheet-hidden-precheck.json`
- `testdata/web-samples/reports/perf-after-worksheet-hidden-precheck.csv`
- `testdata/web-samples/reports/perf-after-docx-related-cache-targets.json`
- `testdata/web-samples/reports/perf-after-docx-related-cache-targets.csv`
- `testdata/web-samples/reports/perf-after-docx-alt-precheck-targets.json`
- `testdata/web-samples/reports/perf-after-docx-alt-precheck-targets.csv`
- `testdata/web-samples/reports/perf-after-docx-related-precheck-targets.json`
- `testdata/web-samples/reports/perf-after-docx-related-precheck-targets.csv`
- `testdata/web-samples/reports/perf-after-image-marker-tighten-targets.json`
- `testdata/web-samples/reports/perf-after-image-marker-tighten-targets.csv`
- `testdata/web-samples/reports/perf-final-optimized-targets.json`
- `testdata/web-samples/reports/perf-final-optimized-targets.csv`
- `testdata/web-samples/reports/perf-after-lazy-image-alts-targets.json`
- `testdata/web-samples/reports/perf-after-lazy-image-alts-targets.csv`
- `testdata/web-samples/reports/perf-final-current-targets.json`
- `testdata/web-samples/reports/perf-final-current-targets.csv`
- `testdata/web-samples/reports/perf-after-simple-inline-xlsx-targets.json`
- `testdata/web-samples/reports/perf-after-simple-inline-xlsx-targets.csv`
- `testdata/web-samples/reports/perf-after-xlsx-visibility-reorder-targets.json`
- `testdata/web-samples/reports/perf-after-xlsx-visibility-reorder-targets.csv`
- `testdata/web-samples/reports/perf-final-xlsx-visibility-reorder.json`
- `testdata/web-samples/reports/perf-final-xlsx-visibility-reorder.csv`
- `testdata/web-samples/reports/perf-after-xlsx-visibility-reorder-final-targets.json`
- `testdata/web-samples/reports/perf-after-xlsx-visibility-reorder-final-targets.csv`
- `testdata/web-samples/reports/perf-after-ooxml-alt-reuse-targets.json`
- `testdata/web-samples/reports/perf-after-ooxml-alt-reuse-targets.csv`
- `testdata/web-samples/reports/perf-after-doc-legacy-markdown-reuse-single-223624.json`
- `testdata/web-samples/reports/perf-after-doc-legacy-markdown-reuse-single-223624.csv`
- `testdata/web-samples/reports/perf-after-doc-legacy-markdown-reuse-targets.json`
- `testdata/web-samples/reports/perf-after-doc-legacy-markdown-reuse-targets.csv`
- `testdata/web-samples/reports/perf-after-image-refs-part-cache-targets.json`
- `testdata/web-samples/reports/perf-after-image-refs-part-cache-targets.csv`
- `testdata/web-samples/reports/perf-after-image-refs-part-cache-single-00003763.json`
- `testdata/web-samples/reports/perf-after-image-refs-part-cache-single-00003763.csv`
- `testdata/web-samples/reports/perf-after-biff-text-emit-fastpath-targets.json`
- `testdata/web-samples/reports/perf-after-biff-text-emit-fastpath-targets.csv`

## 2026-07-06 Current Retained Status

This section summarizes the latest retained state after the newer `.xlsx` optimization pass.

Retained new optimization:
- conservative shared-string worksheet scan in `appendWorksheetText(...)`
- activation is intentionally narrow:
  - no formulas
  - no `inlineStr`
  - no bool / error-string cell types
  - no hyperlink / dataValidation worksheet text path
  - no hidden row / column markers
  - no non-empty header / footer text sections
- when the guard matches, worksheet `<row>/<c>/<v>` content is scanned directly instead of paying
  the full generic `encoding/xml.Decoder` token stream cost

Why it was retained:
- focused hotspot benchmarks improved materially:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    `1367983300 ns/op -> 961121100 ns/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`:
    `832965300 ns/op -> 536413900 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    `2334292400 ns/op -> 1803280200 ns/op`
- same-session serial `.xlsx` keyset comparison on identical inputs confirmed a real win:
  - reverted baseline:
    `testdata/web-samples/reports/perf-exp-ai-assistant-current-baseline-xlsx-keyset-serial-rerun.json`
    - `.xlsx millis = 27461`
  - retained candidate:
    `testdata/web-samples/reports/perf-exp-ai-assistant-sharedscan-xlsx-keyset-serial.json`
    - `.xlsx millis = 23930`
  - same-session net gain:
    `3531 ms`
- output parity held on the decisive `.xlsx` keyset:
  - `textBytes = 279385415`
  - `images = 1`
  - `ok = 21`
  - `errors = 0`
  - `panics = 0`

Latest compatibility evidence:
- `.xlsx` full rerun:
  - `testdata/web-samples/reports/compat-xlsx-20260706-after-sharedscan-candidate.json`
  - result:
    - `1000 / 1000`
    - `errors = 0`
    - `panics = 0`
- full cross-format rerun:
  - `testdata/web-samples/reports/compat-all-20260706-after-sharedscan.json`
  - result:
    - `.doc`: `1000 / 1000`
    - `.docx`: `1000 / 1000`
    - `.ppt`: `1000 / 1000`
    - `.pptx`: `1008 / 1008`
    - `.xls`: `1000 / 1000`
    - `.xlsx`: `1000 / 1000`
    - total: `6008 / 6008`
    - `errors = 0`
    - `panics = 0`

Latest mixed-load evidence:
- mixed 30-file rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-sharedscan-mixed-keyset-serial.json`
- note:
  - this mixed run is still noisy on `.xls` / `.docx`, so it is supporting evidence only
  - the main retention decision continues to rest on the same-session `.xlsx`-only baseline vs
    candidate rerun plus the clean full compatibility sweep

Current authoritative report trail:
- broad experiment log:
  - `testdata/web-samples/reports/followup-report-20260705.md`
- original optimization report:
  - `testdata/web-samples/reports/performance-optimization-report.md`
- `testdata/web-samples/reports/perf-after-biff-text-emit-fastpath-008055-rerun1.json`
- `testdata/web-samples/reports/perf-after-biff-text-emit-fastpath-008055-rerun1.csv`
- `testdata/web-samples/reports/perf-after-biff-text-emit-fastpath-008055-rerun2.json`
- `testdata/web-samples/reports/perf-after-biff-text-emit-fastpath-008055-rerun2.csv`
- `testdata/web-samples/reports/perf-after-biff-hidden-check-dedupe-targets.json`
- `testdata/web-samples/reports/perf-after-biff-hidden-check-dedupe-targets.csv`
- `testdata/web-samples/reports/perf-after-biff-hidden-check-dedupe-008055-rerun.json`
- `testdata/web-samples/reports/perf-after-biff-hidden-check-dedupe-008055-rerun.csv`
- `testdata/web-samples/reports/perf-after-xls-skip-redundant-unique-targets.json`
- `testdata/web-samples/reports/perf-after-xls-skip-redundant-unique-targets.csv`
- `testdata/web-samples/reports/perf-after-xls-skip-redundant-unique-008055-rerun.json`
- `testdata/web-samples/reports/perf-after-xls-skip-redundant-unique-008055-rerun.csv`
- `testdata/web-samples/reports/perf-after-biff-boundsheet-indexbyte-targets.json`
- `testdata/web-samples/reports/perf-after-biff-boundsheet-indexbyte-targets.csv`
- `testdata/web-samples/reports/perf-after-biff-boundsheet-indexbyte-008055-rerun1.json`
- `testdata/web-samples/reports/perf-after-biff-boundsheet-indexbyte-008055-rerun1.csv`
- `testdata/web-samples/reports/perf-after-biff-boundsheet-indexbyte-008055-rerun2.json`
- `testdata/web-samples/reports/perf-after-biff-boundsheet-indexbyte-008055-rerun2.csv`
- `testdata/web-samples/reports/perf-after-xlsx-worksheet-relrefs-fast-targets.json`
- `testdata/web-samples/reports/perf-after-xlsx-worksheet-relrefs-fast-targets.csv`
- `testdata/web-samples/reports/perf-after-xlsx-worksheet-relrefs-fast-00012389-rerun.json`
- `testdata/web-samples/reports/perf-after-xlsx-worksheet-relrefs-fast-00012389-rerun.csv`
- `testdata/web-samples/reports/perf-after-xlsx-worksheet-relrefs-fast-huge-rerun.json`
- `testdata/web-samples/reports/perf-after-xlsx-worksheet-relrefs-fast-huge-rerun.csv`
- `testdata/web-samples/reports/perf-after-worksheet-hiddenranges-byte-scan-targets.json`
- `testdata/web-samples/reports/perf-after-worksheet-hiddenranges-byte-scan-targets.csv`
- `testdata/web-samples/reports/perf-after-worksheet-skip-hidden-chardata-targets.json`
- `testdata/web-samples/reports/perf-after-worksheet-skip-hidden-chardata-targets.csv`
- `testdata/web-samples/reports/perf-after-worksheet-skip-hidden-chardata-00012389-rerun.json`
- `testdata/web-samples/reports/perf-after-worksheet-skip-hidden-chardata-00012389-rerun.csv`
- `testdata/web-samples/reports/perf-after-worksheet-skip-hidden-chardata-targets-final.json`
- `testdata/web-samples/reports/perf-after-worksheet-skip-hidden-chardata-targets-final.csv`
- `testdata/web-samples/reports/perf-after-xlsx-numeric-value-fastpath-targets.json`
- `testdata/web-samples/reports/perf-after-xlsx-numeric-value-fastpath-targets.csv`
- `testdata/web-samples/reports/perf-after-xlsx-number-trim-skip-targets.json`
- `testdata/web-samples/reports/perf-after-xlsx-number-trim-skip-targets.csv`
- `testdata/web-samples/reports/perf-after-xlsx-markdown-cell-guard-targets.json`
- `testdata/web-samples/reports/perf-after-xlsx-markdown-cell-guard-targets.csv`
- `testdata/web-samples/reports/perf-after-xlsx-lazy-large-dedupe-maps-targets.json`
- `testdata/web-samples/reports/perf-after-xlsx-lazy-large-dedupe-maps-targets.csv`
- `testdata/web-samples/reports/perf-after-xlsx-lazy-large-dedupe-maps-rerun.json`
- `testdata/web-samples/reports/perf-after-xlsx-lazy-large-dedupe-maps-rerun.csv`
- `testdata/web-samples/reports/perf-after-cleantext-fold-helpers-targets.json`
- `testdata/web-samples/reports/perf-after-cleantext-fold-helpers-targets.csv`
- `testdata/web-samples/reports/perf-after-cleantext-fold-helpers-rerun.json`
- `testdata/web-samples/reports/perf-after-cleantext-fold-helpers-rerun.csv`
- `testdata/web-samples/reports/perf-after-cleantext-fold-scan-targets.json`
- `testdata/web-samples/reports/perf-after-cleantext-fold-scan-targets.csv`
- `testdata/web-samples/reports/perf-after-cleantext-fold-scan-final-targets.json`
- `testdata/web-samples/reports/perf-after-cleantext-fold-scan-final-targets.csv`
- `testdata/web-samples/reports/perf-after-lowinfo-nomap-targets.json`
- `testdata/web-samples/reports/perf-after-lowinfo-nomap-targets.csv`
- `testdata/web-samples/reports/perf-after-binary-control-nolower-targets.json`
- `testdata/web-samples/reports/perf-after-binary-control-nolower-targets.csv`
- `testdata/web-samples/reports/perf-after-oleclass-nolower-targets.json`
- `testdata/web-samples/reports/perf-after-oleclass-nolower-targets.csv`
- `testdata/web-samples/reports/perf-after-cyrillic-fill-nomap-targets.json`
- `testdata/web-samples/reports/perf-after-cyrillic-fill-nomap-targets.csv`
- `testdata/web-samples/reports/perf-after-glyphsoup-norunes-targets.json`
- `testdata/web-samples/reports/perf-after-glyphsoup-norunes-targets.csv`

## 楠岃瘉

- `go test -buildvcs=false -run 'Test.*XLSX|Test.*Xlsx|Test.*Excel|Test.*Worksheet|Test.*Shared|Test.*Markdown|Test.*Visible|Test.*DOCX|Test.*Docx' .`
- `go test -buildvcs=false -run 'Test.*XLSX|Test.*Xlsx|Test.*Excel|Test.*Worksheet|Test.*Shared|Test.*Markdown|Test.*Visible' .`
- `go test -buildvcs=false -run 'Test.*XLSX|Test.*Xlsx|Test.*Excel|Test.*Worksheet|Test.*Shared|Test.*Markdown|Test.*Visible|Test.*Comment|Test.*Chart|Test.*Drawing' .`
- `go test -buildvcs=false -run 'Test.*XLSX|Test.*Xlsx|Test.*Excel|Test.*Worksheet|Test.*Shared|Test.*Markdown|Test.*Visible|Test.*Comment|Test.*Chart|Test.*Hidden' .`
- `go test -buildvcs=false -run 'Test.*DOCX|Test.*Docx|Test.*Visible|Test.*Header|Test.*Footer|Test.*Embedded|Test.*Drawing|Test.*Image|Test.*Alt|Test.*Media|Test.*Object' .`
- `go test -buildvcs=false ./...`
- `go test -buildvcs=false -run 'Test.*Legacy|Test.*Unicode|Test.*Noise|TestTextQualityFiltersBinaryNoise|TestUTF16NoiseFiltersMisalignedASCII|Test.*XLSX|Test.*Xlsx|Test.*Worksheet|Test.*Hidden' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-image-refs-part-cache-targets.json -csv testdata\web-samples\reports\perf-after-image-refs-part-cache-targets.csv ...`
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Legacy' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-text-emit-fastpath-targets.json -csv testdata\web-samples\reports\perf-after-biff-text-emit-fastpath-targets.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-boundsheet-indexbyte-targets.json -csv testdata\web-samples\reports\perf-after-biff-boundsheet-indexbyte-targets.csv ...`
- `go test -buildvcs=false -run 'Test.*XLSX|Test.*Xlsx|Test.*Excel|Test.*Worksheet|Test.*Comment|Test.*Hidden|Test.*Markdown|Test.*Visible' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-worksheet-hiddenranges-byte-scan-targets.json -csv testdata\web-samples\reports\perf-after-worksheet-hiddenranges-byte-scan-targets.csv ...`
- `go test -buildvcs=false -run 'Test.*XLSX|Test.*Xlsx|Test.*Excel|Test.*Worksheet|Test.*Markdown|Test.*Visible|Test.*Text' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-xlsx-numeric-value-fastpath-targets.json -csv testdata\web-samples\reports\perf-after-xlsx-numeric-value-fastpath-targets.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-xlsx-number-trim-skip-targets.json -csv testdata\web-samples\reports\perf-after-xlsx-number-trim-skip-targets.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-xlsx-markdown-cell-guard-targets.json -csv testdata\web-samples\reports\perf-after-xlsx-markdown-cell-guard-targets.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-xlsx-lazy-large-dedupe-maps-rerun.json -csv testdata\web-samples\reports\perf-after-xlsx-lazy-large-dedupe-maps-rerun.csv ...`
- `go test -buildvcs=false -run 'Test.*XLSX|Test.*Xlsx|Test.*Excel|Test.*Worksheet|Test.*Markdown|Test.*Visible|Test.*Text|Test.*OLE|Test.*Noise' -count=1 .`
- `go test -buildvcs=false -run TestMarkdownBackfillTreatsTableCellBreaksAsVisibleText -count=20 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-cleantext-fold-scan-final-targets.json -csv testdata\web-samples\reports\perf-after-cleantext-fold-scan-final-targets.csv ...`
- `go test -buildvcs=false -run 'Test.*Text|Test.*OLE|Test.*Noise|Test.*Legacy|Test.*Unicode|Test.*XLSX|Test.*Xlsx|Test.*Worksheet|Test.*Markdown|Test.*Visible' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-lowinfo-nomap-targets.json -csv testdata\web-samples\reports\perf-after-lowinfo-nomap-targets.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-binary-control-nolower-targets.json -csv testdata\web-samples\reports\perf-after-binary-control-nolower-targets.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-oleclass-nolower-targets.json -csv testdata\web-samples\reports\perf-after-oleclass-nolower-targets.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-cyrillic-fill-nomap-targets.json -csv testdata\web-samples\reports\perf-after-cyrillic-fill-nomap-targets.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-glyphsoup-norunes-targets.json -csv testdata\web-samples\reports\perf-after-glyphsoup-norunes-targets.csv ...`

## 2026-06-26 杩藉姞浼樺寲璁板綍

- Cyrillic 鎺у埗鍣０鏃╄繃婊わ細`looksLikeCyrillicEncodingTableNoise` 澧炲姞瀵圭煭 PPT 鎺у埗鐗囨锛堟暟瀛?ASCII 鏍囩偣 + 灏戦噺 ASCII 瀛楁瘝 + Cyrillic mojibake锛夌殑绐勫尮閰嶏紝淇 `53446.ppt` 鍓嶇紑鍣０澶栨硠锛屽苟璁╄繖绫荤墖娈垫洿鏃╁湪浜岃繘鍒舵帶鍒跺垽鏂腑杩斿洖锛屽噺灏戝悗缁?Unicode 鍣０鍒嗘敮鍙嶅鎵弿銆?- font table 鐗囨鍒ゆ柇鍘诲垎閰嶏細`looksLikeFontTableFragment` 涓嶅啀涓烘煡鎵?`calibri` 鏋勯€犳暣涓?`strings.ToLower`锛屼篃涓嶅啀鎶婂墠缂€杞崲涓?`[]rune`锛涙敼涓?`indexASCIIFold` 瀛楄妭鎵弿鍜屽墠缂€ range 缁熻锛岃鍒欎繚鎸佷负鈥淐alibri 鍓嶇紑涓瓨鍦ㄩ潪 ASCII 瀛楁瘝涓斿瓧姣嶅潎涓洪潪 ASCII鈥濄€?
鏈疆鐩爣澶嶆祴鏂囦欢锛?- `testdata/web-samples/reports/perf-after-fonttable-index-fold-targets.json`
- `testdata/web-samples/reports/perf-after-fonttable-index-fold-targets.csv`

鏈疆鍏抽敭鏍锋湰澶嶆祴缁撴灉锛?- `00003763.docx`: text=452979, images=10
- `223624.docx`: text=757152, images=33
- `008055.xls`: text=17684747, images=0
- `00012389.xlsx`: text=11148655, images=0
- `testRecordSizeExceeded.xlsx`: text=181333430, images=0

鏈疆楠岃瘉锛?- `go test -buildvcs=false -run 'TestLegacyPPTSampleDropsMojibakeControlLines|TestMojibakePunctuationIsRepairedAndBinaryNoiseDropped|TestMarkdownOutputNormalizesUnicodeSpaces' -count=1 .`
- `go test -buildvcs=false -run 'Test.*Text|Test.*OLE|Test.*Noise|Test.*Legacy|Test.*Unicode|Test.*XLSX|Test.*Xlsx|Test.*Worksheet|Test.*Markdown|Test.*Visible' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-fonttable-index-fold-targets.json -csv testdata\web-samples\reports\perf-after-fonttable-index-fold-targets.csv ...`
- `go test -buildvcs=false ./...`

## 2026-06-27 鎬ц兘澶嶆祴

鏈疆澶嶆祴鏂囦欢锛?- `testdata/web-samples/reports/perf-rerun-20260627-targets.json`
- `testdata/web-samples/reports/perf-rerun-20260627-targets.csv`
- `testdata/web-samples/reports/perf-rerun-20260627-223624-only.json`
- `testdata/web-samples/reports/perf-rerun-20260627-223624-only.csv`

鏈疆缁撴灉涓?2026-06-26 鍩虹嚎瀵规瘮锛?- `00003763.docx`: `7105 ms -> 5474 ms`锛宼ext=`452979`锛宨mages=`10`
- `223624.docx`: `2937 ms -> 4961 ms`锛宼ext=`757152 -> 757034`锛宨mages=`33`
- `008055.xls`: `2746 ms -> 2335 ms`锛宼ext=`17684747`锛宨mages=`0`
- `00012389.xlsx`: `2051 ms -> 1736 ms`锛宼ext=`11148655`锛宨mages=`0`
- `testRecordSizeExceeded.xlsx`: `3482 ms -> 3124 ms`锛宼ext=`181333430`锛宨mages=`0`

缁撹锛?- 4/5 鍏抽敭鏍锋湰鑰楁椂缁х画涓嬮檷锛屼笖鏂囨湰/鍥剧墖璁℃暟绋冲畾銆?- `223624.docx` 鐨?`text=757034` 涓嶆槸姝ｆ枃鍥炲綊锛涘畾浣嶅悗纭锛岃緝鍘嗗彶鍊煎噺灏戠殑涓昏鏉ユ簮鏄袱琛岄噸澶嶇殑 Cyrillic mojibake 鍣０ `00褟褟褘0褟褟褟褟鈥?褮0褝0褞00褯00褜0 褟0=褟]褟 00` 琚?`looksLikeCyrillicEncodingTableNoise` 姝ｅ父杩囨护銆?- 鍥犳鏈疆 `.docx` 鏂囨湰瀛楄妭鏁板彉鍖栧簲瑙嗕负鍣０娓呯悊鍚庣殑鏂扮粨鏋滐紝鑰屼笉鏄吋瀹规€ч€€鍖栵紱鍚庣画鑻ョ户缁互瀛楄妭鏁板仛鍥炲綊闂ㄦ锛岄渶瑕佸悓姝ユ洿鏂?`223624.docx` 鐨勫熀绾胯鐭ャ€?
鏈疆楠岃瘉锛?- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260627-targets.json -csv testdata\web-samples\reports\perf-rerun-20260627-targets.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260627-223624-only.json -csv testdata\web-samples\reports\perf-rerun-20260627-223624-only.csv testdata\web-samples\samples\docx\223624.docx`
- `go test -buildvcs=false -run 'Test.*DOCX|Test.*Docx|Test.*Visible|Test.*Header|Test.*Footer|Test.*Embedded|Test.*Drawing|Test.*Image|Test.*Alt|Test.*Media|Test.*Object|Test.*Text' -count=1 .`

## 2026-06-30 杩藉姞浼樺寲璁板綍

- inline hidden Office 寮曠敤娓呯悊鍏ュ彛 guard锛氫负 `stripInlineHiddenOfficeReferences` 鍜?`looksLikeInlineHiddenOfficeReferenceToken` 澧炲姞 `maybeInlineHiddenOfficeReferenceMarker` 棰勫垽锛涙櫘閫氭鏂囪嫢涓嶅寘鍚祫婧愯矾寰勩€丱ffice 鍏崇郴瀛楁銆丮HTML 澶淬€乣url(...)`銆乣rId`銆佸寘瑁呭紩鐢ㄦ垨缂栫爜璺緞 marker锛屽垯鐩存帴璺宠繃姝ｅ垯鏇挎崲涓?token 灞曞紑銆?- 璇ヤ紭鍖栦富瑕佸墛鍑?`.xlsx` 鏂囨湰涓?Markdown 鐢熸垚涓鏅€氬崟鍏冩牸鍐呭鍙嶅鎵ц闅愯棌 Office 寮曠敤姝ｅ垯娓呯悊鐨勬垚鏈紝涓嶆敼鍙樼湡姝ｅ懡涓殣钘忓紩鐢ㄦ椂鐨勬棦鏈夎繃婊よ鍒欍€?
鏈疆鐩爣澶嶆祴鏂囦欢锛?- `testdata/web-samples/reports/perf-after-inline-office-guard-targets.json`
- `testdata/web-samples/reports/perf-after-inline-office-guard-targets.csv`
- `testdata/web-samples/reports/perf-after-inline-office-guard-disabled-targets.json`
- `testdata/web-samples/reports/perf-after-inline-office-guard-disabled-targets.csv`

鏈疆鍏抽敭鏍锋湰 on/off 澶嶆祴缁撴灉锛?- `00003763.docx`: on=`452979/10`, off=`452979/10`, `8016 ms -> 9679 ms`
- `223624.docx`: on=`760288/33`, off=`760288/33`, `7909 ms -> 12633 ms`
- `008055.xls`: on=`17684747/0`, off=`17684747/0`, `34558 ms -> 99013 ms`
- `00012389.xlsx`: on=`11148655/0`, off=`11148655/0`, `11483 ms -> 23592 ms`
- `testRecordSizeExceeded.xlsx`: on=`181333430/0`, off=`181333430/0`, `17667 ms -> 69303 ms`

鏈疆 profiling 瑙傚療锛?- `testRecordSizeExceeded.xlsx` 涓婏紝`stripInlineHiddenOfficeReferences` 鐨勭疮绉€楁椂浠庣害 `41.36 s` 闄嶅埌绾?`0.30 s`銆?- 鍚屼竴鏍锋湰涓婏紝`extractXlsxMarkdown` 鐨勭疮绉€楁椂浠庣害 `25.17 s` 闄嶅埌绾?`6.62 s`锛宍appendWorksheetText` 浠庣害 `31.27 s` 闄嶅埌绾?`9.94 s`銆?- `223624.docx` 鐨勫綋鍓嶆枃鏈瓧鑺傛暟涓?`760288`锛沢uard 寮€鍏冲墠鍚庣粨鏋滀竴鑷达紝璇存槑璇ユ牱鏈殑褰撳墠鍩虹嚎婕傜Щ涓庢湰杞紭鍖栨棤鍥犳灉鍏崇郴銆?
鏈疆楠岃瘉锛?- `go test -buildvcs=false -run 'TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleTextStripsInlineHiddenOfficeReferences|TestMarkdownImageAltStripsInlineHiddenOfficeReferences|TestLegacyDOCTextStripsInlineHiddenOfficeReferences|TestLegacyXLSTextStripsInlineHiddenOfficeReferences|TestLegacyPPTSampleDropsMojibakeControlLines|TestDocx223624SampleDropsCyrillicMojibakeControlLines|TestMarkdownOutputNormalizesUnicodeSpaces' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-inline-office-guard-targets.json -csv testdata\web-samples\reports\perf-after-inline-office-guard-targets.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-inline-office-guard-disabled-targets.json -csv testdata\web-samples\reports\perf-after-inline-office-guard-disabled-targets.csv ...`

## 2026-06-30 鍐嶆杩藉姞浼樺寲璁板綍

- Markdown 琛ㄦ牸鍗曞厓鏍煎幓閲嶆竻娲楋細涓哄凡閫氳繃 `cleanMarkdownTableCellValue` 鐨?`.xlsx/.xls/.docx` 琛ㄦ牸鍗曞厓鏍兼柊澧?`prepareMarkdownTableCellValue` 鍜?`markdownTablePrepared` 璺緞锛岄伩鍏嶅湪娓叉煋 Markdown 琛ㄦ牸鏃跺啀娆℃墽琛?`cleanText`銆侀殣钘?Office 寮曠敤娓呯悊銆佷簩杩涘埗鎺у埗鍒ゆ柇涓庨噸澶嶈鍘嬬缉銆?- 璇ヤ紭鍖栧彧浣滅敤浜庡凡鍦ㄨ〃鏍奸噰闆嗛樁娈靛畬鎴愭竻娲楃殑鍗曞厓鏍硷紱鏅€?`markdownTable` 璺緞淇濇寔鍘熸湁琛屼负锛岄檷浣庡鍏朵粬 Markdown 鐢熸垚鍦烘櫙鐨勫奖鍝嶈寖鍥淬€?
鏈疆鐩爣澶嶆祴鏂囦欢锛?- `testdata/web-samples/reports/perf-after-markdown-prepared-cells-targets.json`
- `testdata/web-samples/reports/perf-after-markdown-prepared-cells-targets.csv`

鏈疆鍏抽敭鏍锋湰澶嶆祴缁撴灉锛?- `00003763.docx`: text=`452979`, images=`10`, `8016 ms -> 6377 ms`
- `223624.docx`: text=`760288`, images=`33`, `7909 ms -> 4317 ms`
- `008055.xls`: text=`17684747`, images=`0`, `34558 ms -> 23482 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `11483 ms -> 5494 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `17667 ms -> 8553 ms`

鏈疆楠岃瘉锛?- `go test -buildvcs=false -run 'Test.*Markdown|Test.*Worksheet|Test.*XLSX|Test.*Xlsx|TestMarkdownOutputNormalizesUnicodeSpaces|TestCleanVisibleTextStripsInlineHiddenOfficeReferences|TestDocx223624SampleDropsCyrillicMojibakeControlLines' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-markdown-prepared-cells-targets.json -csv testdata\web-samples\reports\perf-after-markdown-prepared-cells-targets.csv ...`

## 鍚庣画寤鸿

- `.xlsx` 褰撳墠鐡堕涓昏鏄?`encoding/xml` 瀵瑰ぇ鍨?worksheet 鐨勯€?token 瑙ｇ爜鍜?181MB 缁撴灉鍐欏叆锛涚户缁ぇ骞呬紭鍖栭渶瑕佽瘎浼版槸鍚﹀紩鍏ヤ笓鐢?worksheet 娴佸紡瑙ｆ瀽鍣ㄣ€?- 涓哄叕寮€ API 澧炲姞鍙厤缃殑鏈€澶ф枃鏈緭鍑洪噺鎴栬秴鏃舵帶鍒讹紝閬垮厤寮傚父瓒呭ぇ鏂囨。鍦ㄩ粯璁よ皟鐢ㄤ腑鍗犵敤杩囦箙銆?## 2026-06-30 BIFF 闅愯棌琛屼紭鍖栬褰?
- `.xls` 鏂囨湰鎻愬彇涓殑闅愯棌琛屽垹闄や粠鈥滄瘡閬囧埌涓€鏉￠殣钘忚璁板綍灏辨暣娈甸噸鎵苟鎼Щ `parts` 鍒囩墖鈥濇敼涓衡€滄寜琛岀紦瀛樺凡杩藉姞鐗囨绱㈠紩骞跺欢杩熸爣璁伴殣钘忊€濓紝閬垮厤 `removeHiddenRowText` 鍦ㄥぇ閲忛殣钘忚宸ヤ綔绨夸腑閫€鍖栦负 O(n^2)銆?- 杩欐鏀瑰姩鍙Е鍙?`biffText` 鐨勯殣钘忚杩囨护瀹炵幇锛屼笉璋冩暣 BIFF 鏂囨湰瑙ｇ爜銆丮arkdown 瑙勫垯鎴栧櫔澹版竻娲楄鍒欍€?
鏈疆澶嶆祴鏂囦欢锛?- `testdata/web-samples/reports/perf-after-biff-hidden-row-marking-targets.json`
- `testdata/web-samples/reports/perf-after-biff-hidden-row-marking-targets.csv`
- `testdata/web-samples/reports/perf-after-biff-hidden-row-marking-xls-anomalies.json`
- `testdata/web-samples/reports/perf-after-biff-hidden-row-marking-xls-anomalies.csv`

鍏抽敭鏍锋湰澶嶆祴缁撴灉锛?- `00003763.docx`: text=`452979`, images=`10`, `5863 ms`
- `223624.docx`: text=`760288`, images=`33`, `4551 ms`
- `008055.xls`: text=`17684747`, images=`0`, `24939 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `5710 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `9172 ms`

寮傚父 `.xls` 鏍锋湰澶嶆祴缁撴灉锛?- `016161.xls`: `219561 ms -> 10633 ms`, text=`4939910 -> 8179182`
- `019088.xls`: `101787 ms -> 98973 ms`, text=`281048`
- `019089.xls`: `107306 ms -> 108061 ms`, text=`306058`

profiling 缁撹锛?- `016161.xls` 浼樺寲鍓?CPU 鏃堕棿绾?`95%` 闆嗕腑鍦?`biffText.removeHiddenRowText`锛屽搴旈殣钘忚鍙嶅閲嶆壂宸叉敹闆嗘枃鏈紱浼樺寲鍚庤閫€鍖栨秷澶憋紝鏂囨湰瀛楄妭鎭㈠鍒板巻鍙插熀绾裤€?- `019088.xls` / `019089.xls` 浠嶄富瑕佽€楁椂鍦?`biffMarkdown -> missingMarkdownText(...)` 鐨?Markdown 鍥炲～鍖呭惈鍒ゆ柇锛岃鏄?`.xls` 闀垮熬鍓╀綑鐡堕宸蹭粠闅愯棌琛屽垹闄よ浆绉诲埌鍥炲～姣斿闃舵銆?
鏈疆楠岃瘉锛?- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*Markdown|Test.*Worksheet|Test.*Visible|Test.*Legacy' -count=1 .`
- `go test -buildvcs=false -timeout 12m ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-hidden-row-marking-targets.json -csv testdata\web-samples\reports\perf-after-biff-hidden-row-marking-targets.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-hidden-row-marking-xls-anomalies.json -csv testdata\web-samples\reports\perf-after-biff-hidden-row-marking-xls-anomalies.csv ...`

## 2026-06-30 Markdown 鍥炲～閲嶅琛岀紦瀛樿褰?
- `missingMarkdownText(...)` 鏂板鎸?`line` 缂撳瓨鈥滆琛屾槸鍚﹂渶瑕佽ˉ鍥?Additional Text鈥濈殑鍒ゅ畾缁撴灉锛岄伩鍏?`.xls` 闀垮熬鏍锋湰鍦ㄥぇ閲忛噸澶嶆枃鏈涓婂弽澶嶆墽琛?`markdownBackfillVisibleText`銆乣markdownVisibleLineText`銆乧overage/containment 鏌ヨ銆?- 杩欐鏀瑰姩涓嶈皟鏁?Markdown 鍥炲～瑙勫垯锛屽彧澶嶇敤鍚屼竴瑙勮寖鍖栬鐨勬棦鏈夊垽鏂粨鏋溿€?
鏈疆澶嶆祴鏂囦欢锛?- `testdata/web-samples/reports/perf-after-backfill-line-cache-xls-019088-019089.json`
- `testdata/web-samples/reports/perf-after-backfill-line-cache-xls-019088-019089.csv`
- `testdata/web-samples/reports/perf-after-backfill-line-cache-xls-anomalies.json`
- `testdata/web-samples/reports/perf-after-backfill-line-cache-xls-anomalies.csv`

寮傚父 `.xls` 鏍锋湰澶嶆祴缁撴灉锛?- `019088.xls`: `98973 ms -> 86751 ms -> 90159 ms`, text=`281048`
- `019089.xls`: `108061 ms -> 93806 ms -> 99481 ms`, text=`306058`
- `016161.xls`: text=`8179182`, `10977 ms`
- `008055.xls`: text=`17684747`, `25654 ms`

缁撹锛?- 璇ヤ紭鍖栧 `019088.xls` / `019089.xls` 鏈夌ǔ瀹氫絾鏈夐檺鐨勬敹鐩婏紝璇存槑杩欎袱浠芥牱鏈‘瀹炲瓨鍦ㄥぇ閲忛噸澶嶈鍥炲～鍒ゅ畾鎴愭湰銆?- `.xls` 褰撳墠鍓╀綑涓荤摱棰堜粛鍦?`missingMarkdownText(...)` 璺緞鏈韩锛屽悗缁户缁紭鍖栧簲浼樺厛鍥寸粫 containment / coverage 鐨勬洿绮楃矑搴︾煭璺垨鎸?sheet/table 绾у埆鍑忓皯鍥炲～鍊欓€夎銆?
鏈疆楠岃瘉锛?- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*Markdown|Test.*Worksheet|Test.*Visible|Test.*Legacy' -count=1 .`
- `go test -buildvcs=false -timeout 12m ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-line-cache-xls-019088-019089.json -csv testdata\web-samples\reports\perf-after-backfill-line-cache-xls-019088-019089.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-line-cache-xls-anomalies.json -csv testdata\web-samples\reports\perf-after-backfill-line-cache-xls-anomalies.csv ...`

## 2026-06-30 鍗曞瓧绗︽姌鍙犳瑕嗙洊缂撳瓨璁板綍

- `missingMarkdownText(...)` 涓 `shouldCollapseMarkdownSingleCharacterRun(run)` 鍛戒腑鐨?`collapsed := strings.Join(run, "")` 鏂板瑕嗙洊鍒ゆ柇缂撳瓨锛岄伩鍏?`.xls` 闀垮熬鏍锋湰瀵瑰悓涓€鎶樺彔瀛楃涓插弽澶嶆墽琛?`coverage` / `containment` 鏌ヨ銆?- 杩欐鏀瑰姩鍙懡涓€滃崟瀛楃琛岃繛缁鎶樺彔鈥濆垎鏀紝涓嶆敼鍙樻櫘閫氳鍥炲～琛屼负銆?
鏈疆 profiling 缁撹锛?- `019089.xls` 鍦ㄤ笂涓€鐗堜腑绾?`93%` CPU 鏃堕棿闆嗕腑鍦?`missingMarkdownText(...)`锛屽叾涓害 `88s` 钀藉湪鍗曞瓧绗︽姌鍙犳鐨?`markdownBackfillCoverageCoversLine(...) || containment...` 鍒ゆ柇銆?- 鍔犲叆 `collapsedCoveredCache` 鍚庯紝杩欎釜鐑偣琚樉钁楀帇浣庯紝`.xls` 闀垮熬鏍锋湰鑰楁椂鍥炶惤鍒?2-3 涓囨绉掗噺绾с€?
鏈疆澶嶆祴鏂囦欢锛?- `testdata/web-samples/reports/perf-after-backfill-collapsed-cache-xls-019089.json`
- `testdata/web-samples/reports/perf-after-backfill-collapsed-cache-xls-019089.csv`
- `testdata/web-samples/reports/perf-after-backfill-collapsed-cache-xls-anomalies.json`
- `testdata/web-samples/reports/perf-after-backfill-collapsed-cache-xls-anomalies.csv`

寮傚父 `.xls` 鏍锋湰澶嶆祴缁撴灉锛?- `019088.xls`: `90159 ms -> 29213 ms`, text=`281048`
- `019089.xls`: `99481 ms -> 22521 ms -> 24632 ms`, text=`306058`
- `016161.xls`: text=`8179182`, `11466 ms`
- `008055.xls`: text=`17684747`, `25748 ms`

缁撹锛?- 杩欐浼樺寲瀵?`019088.xls` / `019089.xls` 甯︽潵浜嗘槑鏄炬敹鐩婏紝璇存槑杩欐壒鏍锋湰鐨勫ぇ澶村苟涓嶅湪鏅€?Markdown 鍥炲～锛岃€屽湪閲嶅鍑虹幇鐨勫崟瀛楃鎶樺彔娈佃鐩栧垽鏂€?- 褰撳墠 `.xls` 璺緞鍓╀綑鐑偣浠嶅湪 `missingMarkdownText(...)`锛屼絾宸茬粡浠?9-10 涓囨绉掔骇杩涗竴姝ュ帇鍒颁簡 2-3 涓囨绉掔骇銆?
鏈疆楠岃瘉锛?- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*Markdown|Test.*Worksheet|Test.*Visible|Test.*Legacy' -count=1 .`
- `go test -buildvcs=false -timeout 12m ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-collapsed-cache-xls-019089.json -csv testdata\web-samples\reports\perf-after-backfill-collapsed-cache-xls-019089.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-collapsed-cache-xls-anomalies.json -csv testdata\web-samples\reports\perf-after-backfill-collapsed-cache-xls-anomalies.csv ...`

## 2026-06-30 containment 鏈€澶ч暱搴︾煭璺褰?
- `markdownBackfillContainment` 鏂板 `tableRawMaxLen`銆乣tableVisibleMaxLen`銆乣tableComparableMaxLen`銆乣visibleMaxLen`锛屽湪 `tableTextContainsLine(...)` / `visibleLineContainsLine(...)` 鎵ц `strings.Contains(...)` 鍓嶅厛鍋氣€滃緟鏌ヤ覆闀垮害鏄惁鍙兘钀藉叆浠讳竴鍗曡鈥濈殑鐭矾銆?- 杩欐鏀瑰姩涓嶆敼鍙樺寘鍚垽鏂涔夛紝鍙伩鍏嶆槑鏄句笉鍙兘鍛戒腑鐨勯暱瀛楃涓插幓鎵弿鏁存鎷兼帴缂撳瓨銆?
鏈疆澶嶆祴鏂囦欢锛?- `testdata/web-samples/reports/perf-after-backfill-maxlen-xls-019088-019089.json`
- `testdata/web-samples/reports/perf-after-backfill-maxlen-xls-019088-019089.csv`
- `testdata/web-samples/reports/perf-after-backfill-maxlen-xls-anomalies.json`
- `testdata/web-samples/reports/perf-after-backfill-maxlen-xls-anomalies.csv`

寮傚父 `.xls` 鏍锋湰澶嶆祴缁撴灉锛?- `019088.xls`: `29213 ms -> 23411 ms -> 25500 ms`, text=`281048`
- `019089.xls`: `24632 ms -> 22750 ms -> 24323 ms`, text=`306058`
- `016161.xls`: text=`8179182`, `11009 ms`
- `008055.xls`: text=`17684747`, `24204 ms`

缁撹锛?- 杩欐潯鐭矾瀵?`019088.xls` 鍜?`019089.xls` 缁х画甯︽潵绋冲畾鏀剁泭锛岃鏄庡綋鍓嶉暱灏炬牱鏈噷浠嶆湁涓嶅皯鈥滈暱搴︿笂涓嶅彲鑳藉懡涓€濈殑 containment 鏌ヨ銆?- `.xls` 鍥炲～鐑偣杩樺湪锛屼絾杩欐潯绾垮凡缁忎粠鏈€鍒濈殑 9-10 涓囨绉掔骇鎸佺画鍘嬪埌浜嗙害 2.3-2.5 涓囨绉掔骇銆?
鏈疆楠岃瘉锛?- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*Markdown|Test.*Worksheet|Test.*Visible|Test.*Legacy' -count=1 .`
- `go test -buildvcs=false -timeout 12m ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-maxlen-xls-019088-019089.json -csv testdata\web-samples\reports\perf-after-backfill-maxlen-xls-019088-019089.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-maxlen-xls-anomalies.json -csv testdata\web-samples\reports\perf-after-backfill-maxlen-xls-anomalies.csv ...`

## 2026-06-30 candidateLine 缂撳瓨璁板綍

- `missingMarkdownText(...)` 鏂板鎸夊師濮嬭緭鍏ヨ缂撳瓨 `markdownBackfillCandidateLine(...)` 缁撴灉锛屽噺灏?`.xls` 鍥炲～闃舵瀵瑰悓涓€鍘熷鏂囨湰閲嶅鎵ц `cleanText + stripInlineHiddenOfficeReferences + TrimSpace`銆?- 杩欐鏀瑰姩鍙紦瀛樻簮琛岃鑼冨寲缁撴灉锛屼笉鏀瑰彉鍥炲～鍒ゆ柇椤哄簭涓庤鍒欍€?
鏈疆澶嶆祴鏂囦欢锛?- `testdata/web-samples/reports/perf-after-backfill-candidate-cache-xls-019088-019089.json`
- `testdata/web-samples/reports/perf-after-backfill-candidate-cache-xls-019088-019089.csv`
- `testdata/web-samples/reports/perf-after-backfill-candidate-cache-xls-anomalies.json`
- `testdata/web-samples/reports/perf-after-backfill-candidate-cache-xls-anomalies.csv`

鐩爣 `.xls` 鏍锋湰澶嶆祴缁撴灉锛?- `019088.xls`: `25500 ms -> 24642 ms`, text=`281048`
- `019089.xls`: `24323 ms -> 23488 ms`, text=`306058`

寮傚父 `.xls` 鏍锋湰鏁寸粍澶嶆祴缁撴灉锛?- `008055.xls`: text=`17684747`, `30373 ms`
- `013623.xls`: text=`1329843`, `5308 ms`
- `016161.xls`: text=`8179182`, `11663 ms`
- `018548.xls`: text=`1541448`, `8104 ms`
- `019088.xls`: text=`281048`, `26663 ms`
- `019089.xls`: text=`306058`, `28207 ms`

缁撹锛?- 杩欐潯缂撳瓨鏄綆椋庨櫓銆佸皬鏀剁泭浼樺寲锛涚洰鏍囨牱鏈笂鏈夊皬骞呬笅闄嶏紝浣嗘暣缁?`.xls` 澶嶆祴瀛樺湪涓€瀹氭尝鍔紝璇存槑瀹冩洿澶氭槸鍦ㄥ凡浼樺寲璺緞涓婄户缁墛鍑忛噸澶嶆竻娲楀紑閿€锛岃€屼笉鏄敼鍙樹富鐑偣缁撴瀯銆?
鏈疆楠岃瘉锛?- `go test -buildvcs=false -run 'TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|Test.*XLS|Test.*Biff|Test.*Markdown|Test.*Worksheet|Test.*Visible|Test.*Legacy' -count=1 .`
- `go test -buildvcs=false -timeout 12m ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-candidate-cache-xls-019088-019089.json -csv testdata\web-samples\reports\perf-after-backfill-candidate-cache-xls-019088-019089.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-candidate-cache-xls-anomalies.json -csv testdata\web-samples\reports\perf-after-backfill-candidate-cache-xls-anomalies.csv ...`

## 2026-06-30 disallowed line 缂撳瓨璁板綍

- `missingMarkdownText(...)` 灏?`lineMissingCache` 鐨勫懡涓鏌ュ墠绉伙紝骞跺 `markdownBackfillNormalizedLineAllowed(...)` 鍒ゅ畾涓衡€滀笉鍏佽鍥炲～鈥濈殑瑙勮寖鍖栬鐩存帴缂撳瓨 `false`锛岃閲嶅鍑虹幇鐨勪綆淇℃伅/闅愯棌寮曠敤/鍏冩暟鎹櫔澹拌鏇存棭鐭矾銆?- 杩欐鏀瑰姩涓嶆敼鍙樺洖濉鍒欙紝鍙噺灏戦噸澶嶇殑鈥滄棤鏁堣鈥濆垽瀹氫笌鍚庣画鍒嗘敮杩涘叆銆?
鏈疆澶嶆祴鏂囦欢锛?- `testdata/web-samples/reports/perf-after-backfill-disallowed-cache-xls-019088-019089.json`
- `testdata/web-samples/reports/perf-after-backfill-disallowed-cache-xls-019088-019089.csv`
- `testdata/web-samples/reports/perf-after-backfill-disallowed-cache-xls-anomalies.json`
- `testdata/web-samples/reports/perf-after-backfill-disallowed-cache-xls-anomalies.csv`

鐩爣 `.xls` 鏍锋湰澶嶆祴缁撴灉锛?- `019088.xls`: `25500 ms -> 24598 ms`, text=`281048`
- `019089.xls`: `24323 ms -> 22914 ms`, text=`306058`

寮傚父 `.xls` 鏍锋湰鏁寸粍澶嶆祴缁撴灉锛?- `008055.xls`: text=`17684747`, `24279 ms`
- `013623.xls`: text=`1329843`, `4479 ms`
- `016161.xls`: text=`8179182`, `11005 ms`
- `018548.xls`: text=`1541448`, `6916 ms`
- `019088.xls`: text=`281048`, `24841 ms`
- `019089.xls`: text=`306058`, `23763 ms`

缁撹锛?- 杩欐潯浼樺寲灞炰簬浣庨闄┿€佺ǔ瀹氱殑灏忔敹鐩婄紦瀛橈紱瀹冩洿鍋忓悜鍑忓皯閲嶅鍣０琛岀殑鏃犳晥鍒ゅ畾锛屽 `019089.xls` 鐨勫府鍔╂瘮 `019088.xls` 鏇存槑鏄俱€?- `.xls` 鍥炲～閾捐矾鐩墠宸茬粡鍩烘湰杩涘叆鈥滄寔缁姞闆剁寮€閿€鈥濈殑闃舵锛屾敹鐩婂紑濮嬪彉灏忥紝浣嗕粛鐒惰兘鍦ㄤ笉鏀硅涔夌殑鍓嶆彁涓嬬户缁線涓嬪帇銆?
鏈疆楠岃瘉锛?- `go test -buildvcs=false -run 'TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|Test.*XLS|Test.*Biff|Test.*Markdown|Test.*Worksheet|Test.*Visible|Test.*Legacy' -count=1 .`
- `go test -buildvcs=false -timeout 12m ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-disallowed-cache-xls-019088-019089.json -csv testdata\web-samples\reports\perf-after-backfill-disallowed-cache-xls-019088-019089.csv ...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-disallowed-cache-xls-anomalies.json -csv testdata\web-samples\reports\perf-after-backfill-disallowed-cache-xls-anomalies.csv ...`

## 2026-06-30 cleanVisibleText discardable-line guard

- Added a cheap pre-check for hidden Office references and metadata before calling the expensive
  `looksLikeHiddenResourceReference(...)`, `looksLikeRelationshipIDReference(...)`,
  `looksLikeOfficeRelationshipMetadataReference(...)`, and
  `looksLikeOfficeXMLMetadataReference(...)` chain.
- The optimization is intentionally narrow: it does not change the discard rules, it only avoids
  running the expensive detectors for ordinary visible text lines and attribute values that do not
  contain any relevant markers.
- The shared guard is now used by:
  - `cleanVisibleText(...)`
  - `cleanMarkdownVisibleText(...)`
  - `cleanVisibleAttributeValue(...)`

Validation:
- `go test -buildvcs=false -run "TestCleanVisibleTextKeepsRelationshipModeWordsInProse|TestLongVisibleTextWithOfficeMetadataWordsIsKept|TestLongOfficeMetadataReferencesAreStillRecognized|TestCleanVisibleTextDropsOfficeXMLNamespaceMetadata|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns" ./...`
- `go test -buildvcs=false ./...`
- `go run ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-clean-visible-guard-xls.json -csv testdata\web-samples\reports\perf-after-clean-visible-guard-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-clean-visible-guard-keyset.json -csv testdata\web-samples\reports\perf-after-clean-visible-guard-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Anomalous `.xls` sample rerun:
- `008055.xls`: `24279 ms -> 22505 ms -> 22970 ms`, text=`17684747`
- `013623.xls`: `4479 ms -> 4376 ms`, text=`1329843`
- `016161.xls`: `11005 ms -> 10083 ms`, text=`8179182`
- `018548.xls`: `6916 ms -> 6254 ms`, text=`1541448`
- `019088.xls`: `24841 ms -> 24021 ms`, text=`281048`
- `019089.xls`: `23763 ms -> 21867 ms`, text=`306058`

Key sample verification:
- `00003763.docx`: text=`452979`, images=`10`, `5643 ms`
- `223624.docx`: text=`760288`, images=`33`, `4135 ms`
- `008055.xls`: text=`17684747`, images=`0`, `22970 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `5488 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `8589 ms`

Conclusion:
- This is a low-risk optimization with a stable improvement on the current `.xls` long-tail set.
- The largest direct win is on `008055.xls`, where the reruns consistently moved from the previous
  `24.3 s` baseline down into the `22.x s` range without changing output size.
- `016161.xls`, `018548.xls`, and `019089.xls` also improved measurably, while key `docx/xlsx`
  outputs stayed on their expected text/image counts.

## 2026-06-30 markdown table-row guard

- Tightened `markdownVisibleTableCells(...)` so it only parses lines that already satisfy
  `markdownLikelyTableRow(...)`, instead of attempting table-cell extraction for every line that
  merely contains a `|`.
- This keeps the existing table coverage rules unchanged, but avoids repeated table splitting and
  cell normalization work for ordinary prose lines with incidental pipe characters.
- The change is intentionally conservative because the repository already had a dedicated
  single-pipe prose regression test.

Validation:
- `go test -buildvcs=false -run "TestMarkdownBackfillTreatsTableCellBreaksAsVisibleText|TestMarkdownBackfillTreatsHardLineBreakMarkersAsControlText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns" ./...`
- `go test -buildvcs=false ./...`
- `go run ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-markdown-table-row-guard-xls.json -csv testdata\web-samples\reports\perf-after-markdown-table-row-guard-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-markdown-table-row-guard-keyset.json -csv testdata\web-samples\reports\perf-after-markdown-table-row-guard-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Anomalous `.xls` sample rerun:
- `008055.xls`: `22970 ms -> 21787 ms -> 22804 ms`, text=`17684747`
- `013623.xls`: `4376 ms -> 3975 ms`, text=`1329843`
- `016161.xls`: `10083 ms -> 10094 ms`, text=`8179182`
- `018548.xls`: `6254 ms -> 6072 ms`, text=`1541448`
- `019088.xls`: `24021 ms -> 25107 ms`, text=`281048`
- `019089.xls`: `21867 ms -> 21343 ms`, text=`306058`

Key sample verification:
- `00003763.docx`: text=`452979`, images=`10`, `5539 ms`
- `223624.docx`: text=`760288`, images=`33`, `4181 ms`
- `008055.xls`: text=`17684747`, images=`0`, `22804 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `5273 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `8140 ms`

Conclusion:
- This is another low-risk optimization with the clearest benefit on `008055.xls`, whose current
  reruns now reach the `21.8 s` range while preserving output size.
- The gain is smaller and noisier on `019088.xls`, which suggests the remaining `.xls`
  long-tail cost is still dominated by later Markdown backfill logic rather than table-row
  candidate detection alone.

## 2026-06-30 coverage comparable guard

- Added `addMarkdownBackfillCoverageText(...)` to centralize coverage insertion and skip
  `markdownBackfillComparableText(...)` unless the source text contains markers that can actually
  change comparable normalization.
- The new `markdownBackfillComparableMayDiffer(...)` guard is conservative: it only returns false
  for inputs that are already plain visible text, so the optimization reduces redundant comparable
  work without changing coverage semantics.
- This primarily targets the `addMarkdownBackfillCoverage(...)` hotspot identified in the `.xls`
  long-tail samples.

Validation:
- `go test -buildvcs=false -run "TestMarkdownBackfillTreatsHeadingsAsVisibleText|TestMarkdownBackfillTreatsListItemsAsVisibleText|TestMarkdownBackfillTreatsBlockquotesAsVisibleText|TestMarkdownBackfillTreatsFootnoteDefinitionsAsVisibleText|TestMarkdownBackfillTreatsInlineFormattingAsVisibleText|TestMarkdownBackfillTreatsHTMLTagsAsVisibleText|TestMarkdownBackfillTreatsTableCellBreaksAsVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns" ./...`
- `go test -buildvcs=false ./...`
- `go run ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-coverage-comparable-guard-xls.json -csv testdata\web-samples\reports\perf-after-coverage-comparable-guard-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-coverage-comparable-guard-keyset.json -csv testdata\web-samples\reports\perf-after-coverage-comparable-guard-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Anomalous `.xls` sample rerun:
- `008055.xls`: `21787 ms -> 20707 ms`, text=`17684747`
- `013623.xls`: `3975 ms -> 4311 ms`, text=`1329843`
- `016161.xls`: `10094 ms -> 9346 ms`, text=`8179182`
- `018548.xls`: `6072 ms -> 5645 ms`, text=`1541448`
- `019088.xls`: `25107 ms -> 22135 ms`, text=`281048`
- `019089.xls`: `21343 ms -> 20584 ms`, text=`306058`

Key sample output verification:
- `00003763.docx`: text=`452979`, images=`10`
- `223624.docx`: text=`760288`, images=`33`
- `008055.xls`: text=`17684747`, images=`0`
- `00012389.xlsx`: text=`11148655`, images=`0`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`

Note:
- The key-sample reruns in the current sandbox used a workspace-local `GOCACHE`, so their wall
  times were noisier than the earlier warm-cache measurements. They were used here to confirm
  output stability, while the anomalous `.xls` timings above remain the main performance signal.

Conclusion:
- This optimization produced the strongest `.xls` improvement of the recent coverage-focused
  tweaks, pushing `008055.xls` down to about `20.7 s` and `019089.xls` down to about `20.6 s`
  without changing extracted text or image counts.
- The remaining `.xls` hot path is still in Markdown backfill, but this trims another chunk of
  redundant coverage construction work before the later containment checks begin.

## 2026-06-30 visible-text textual guard

- Tightened `looksLikeDiscardableVisibleTextLine(...)` with an additional early return for lines
  that already look like ordinary visible text and do not carry hidden-reference or control-text
  markers.
- The new short-circuit keeps the existing discard rules intact for hidden Office references,
  binary/control fragments, and metadata-like lines, but avoids running the full expensive chain
  on the common case of plain readable text.
- This specifically targets the `cleanVisibleText(...)` single-line hotspot identified in the
  latest `.xls` profiling work.

Validation:
- `go test -buildvcs=false -run "TestCleanVisibleTextDropsInternalOfficePartReferences|TestCleanVisibleTextStripsInlineHiddenOfficeReferences|TestCleanVisibleTextKeepsRelationshipModeWordsInProse|TestLongVisibleTextWithOfficeMetadataWordsIsKept|TestLongOfficeMetadataReferencesAreStillRecognized|TestCleanVisibleTextDropsOfficeXMLNamespaceMetadata|TestOLEIdentifierFragmentsAreFilteredConservatively|TestCleanVisibleTextKeepsShortCJKAndKoreanLines|TestCleanVisibleTextKeepsCyrillicSheetNames|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns" ./...`
- `go test -buildvcs=false ./...`
- `go run ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-visible-text-textual-guard-xls.json -csv testdata\web-samples\reports\perf-after-visible-text-textual-guard-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-visible-text-textual-guard-keyset.json -csv testdata\web-samples\reports\perf-after-visible-text-textual-guard-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Anomalous `.xls` sample rerun:
- `008055.xls`: `20707 ms -> 20307 ms -> 22193 ms`, text=`17684747`
- `013623.xls`: `4311 ms -> 4690 ms`, text=`1329843`
- `016161.xls`: `9346 ms -> 8210 ms`, text=`8179182`
- `018548.xls`: `5645 ms -> 5634 ms`, text=`1541448`
- `019088.xls`: `22135 ms -> 22678 ms`, text=`281048`
- `019089.xls`: `20584 ms -> 21591 ms`, text=`306058`

Key sample verification:
- `00003763.docx`: text=`452979`, images=`10`, `6407 ms`
- `223624.docx`: text=`760288`, images=`33`, `4408 ms`
- `008055.xls`: text=`17684747`, images=`0`, `22193 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `5362 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `8964 ms`

Conclusion:
- This is a low-risk `cleanVisibleText(...)` optimization with stable output and a meaningful gain
  on the current `.xls` long-tail set.
- The clearest wins are `008055.xls` dropping to about `20.3 s` on the anomaly rerun and
  `016161.xls` dropping to about `8.2 s`, while the key `docx/xlsx/xls` sample counts remain
  unchanged.

## 2026-06-30 markdown table cell single-pass cleanup

- Reworked `cleanMarkdownTableCellValue(...)` to avoid the old `Split -> append -> Join` flow on
  every BIFF/XLS table cell.
- The new version keeps the existing filtering rules but uses a single-pass newline scan with a
  dedicated `markdownTableCellValueLineDiscarded(...)` helper, which cuts per-cell slice
  allocation and avoids building temporary line arrays for the common case.
- Single-line cell values now take a cheaper fast path after `cleanText(...)` and
  `stripInlineHiddenOfficeReferences(...)`, which is where the hot `.xls` path spends most of its
  time.

Validation:
- `go test -buildvcs=false -run "TestMarkdownTableCell|TestCleanVisibleTextDropsInternalOfficePartReferences|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns" ./...`
- `go test -buildvcs=false ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-markdown-cell-single-pass-xls.json -csv testdata\web-samples\reports\perf-after-markdown-cell-single-pass-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-markdown-cell-single-pass-keyset.json -csv testdata\web-samples\reports\perf-after-markdown-cell-single-pass-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-markdown-cell-single-pass-008055-rerun.json -csv testdata\web-samples\reports\perf-after-markdown-cell-single-pass-008055-rerun.csv testdata\web-samples\samples\xls\008055.xls`

Anomalous `.xls` sample rerun:
- `008055.xls`: `20955 ms -> 20657 ms`, text=`17684747`
- `013623.xls`: `4244 ms`, text=`1329843`
- `016161.xls`: `9152 ms`, text=`8179182`
- `018548.xls`: `5717 ms`, text=`1541448`
- `019088.xls`: `21728 ms`, text=`281048`
- `019089.xls`: `19454 ms`, text=`306058`

Key sample verification:
- `00003763.docx`: text=`452979`, images=`10`, `5558 ms`
- `223624.docx`: text=`760288`, images=`33`, `4199 ms`
- `008055.xls`: text=`17684747`, images=`0`, `21612 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `5298 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `8221 ms`

Conclusion:
- This change stays within the current stable output envelope and improves the BIFF/XLS table path
  without touching the riskier Markdown backfill logic.
- The clearest gains are on `019089.xls`, `019088.xls`, and the repeat runs of `008055.xls`,
  while the key `docx/xlsx/xls` sample counts remain unchanged.

## 2026-06-30 BIFF cell redundant binary check removal

- Removed a redundant `looksLikeBinaryControlFragment(...)` call from `biffSetMarkdownCellAt(...)`
  after `cleanMarkdownTableCellValue(...)`.
- The markdown table cell cleaner already strips binary-control-only lines and returns `""` when no
  visible text remains, so the follow-up whole-value binary-control check in the BIFF cell hot path
  was repeating work without changing output.
- This keeps the PDF-text and spreadsheet-formula guards in place, but trims one more expensive
  classification call from every surviving BIFF/XLS cell candidate.

Validation:
- `go test -buildvcs=false -run "TestMarkdownTableCell|TestCleanVisibleTextDropsInternalOfficePartReferences|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns" ./...`
- `go test -buildvcs=false ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-cell-redundant-binary-check-removal-xls.json -csv testdata\web-samples\reports\perf-after-biff-cell-redundant-binary-check-removal-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-cell-redundant-binary-check-removal-keyset.json -csv testdata\web-samples\reports\perf-after-biff-cell-redundant-binary-check-removal-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-cell-redundant-binary-check-removal-008055-rerun.json -csv testdata\web-samples\reports\perf-after-biff-cell-redundant-binary-check-removal-008055-rerun.csv testdata\web-samples\samples\xls\008055.xls`

Anomalous `.xls` sample rerun:
- `008055.xls`: `20029 ms`, text=`17684747`
- `013623.xls`: `3975 ms`, text=`1329843`
- `016161.xls`: `8463 ms`, text=`8179182`
- `018548.xls`: `5400 ms`, text=`1541448`
- `019088.xls`: `21660 ms`, text=`281048`
- `019089.xls`: `19291 ms`, text=`306058`

Key sample verification:
- `00003763.docx`: text=`452979`, images=`10`, `6715 ms`
- `223624.docx`: text=`760288`, images=`33`, `4494 ms`
- `008055.xls`: text=`17684747`, images=`0`, `20839 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `5355 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `8534 ms`

Repeat check:
- `008055.xls`: text=`17684747`, images=`0`, `21245 ms`

Conclusion:
- This is another low-risk BIFF/XLS hot-path cleanup with stable output and broad improvement
  across the current anomalous `.xls` set.
- The strongest gains here are `008055.xls` reaching about `20.0 s`, `016161.xls` dropping to
  about `8.5 s`, and `019089.xls` moving under `19.3 s`, while the key sample counts remain
  unchanged.

## 2026-06-30 BIFF comment redundant binary check removal

- Removed redundant `looksLikeBinaryControlFragment(...)` checks from `cleanBIFFCommentText(...)`
  after `cleanVisibleText(...)`.
- `cleanVisibleText(...)` already strips discardable visible-text lines, including binary-control
  fragments, so the extra whole-text and per-line binary-control checks inside BIFF comment cleanup
  were reclassifying text that had already passed through the visible-text filter.
- The BIFF comment path still keeps the `looksLikeEmbeddedPDFText(...)` and
  `looksLikeSpreadsheetFormulaExpression(...)` guards, so this only trims repeated binary-control
  classification work from the comment hot path.

Validation:
- `go test -buildvcs=false -run "TestMarkdownTableCell|TestCleanVisibleTextDropsInternalOfficePartReferences|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestBIFFHeaderFooterControlCodesAreNotVisibleText" ./...`
- `go test -buildvcs=false ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-comment-redundant-binary-check-removal-xls.json -csv testdata\web-samples\reports\perf-after-biff-comment-redundant-binary-check-removal-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-comment-redundant-binary-check-removal-keyset.json -csv testdata\web-samples\reports\perf-after-biff-comment-redundant-binary-check-removal-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-comment-redundant-binary-check-removal-008055-rerun.json -csv testdata\web-samples\reports\perf-after-biff-comment-redundant-binary-check-removal-008055-rerun.csv testdata\web-samples\samples\xls\008055.xls`

Anomalous `.xls` sample rerun:
- `008055.xls`: `16920 ms`, text=`17684747`
- `013623.xls`: `3578 ms`, text=`1329843`
- `016161.xls`: `7860 ms`, text=`8179182`
- `018548.xls`: `4621 ms`, text=`1541448`
- `019088.xls`: `19931 ms`, text=`281048`
- `019089.xls`: `17779 ms`, text=`306058`

Key sample verification:
- `00003763.docx`: text=`452979`, images=`10`, `5039 ms`
- `223624.docx`: text=`760288`, images=`33`, `3971 ms`
- `008055.xls`: text=`17684747`, images=`0`, `18976 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `4946 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `7625 ms`

Repeat check:
- `008055.xls`: text=`17684747`, images=`0`, `18143 ms`

Conclusion:
- This is another low-risk BIFF/XLS cleanup that compounds well with the earlier cell-path change.
- The clearest wins are `008055.xls` moving into the high-`16 s` to low-`18 s` range on reruns,
  `016161.xls` dropping below `8 s`, and `019089.xls` reaching about `17.8 s`, while the key
  sample counts remain unchanged.

## 2026-06-30 BIFF text redundant binary check removal

- Removed redundant `looksLikeBinaryControlFragment(...)` checks from `forEachBIFFText(...)` after
  `cleanVisibleText(...)`.
- `cleanVisibleText(...)` already discards binary-control-like visible text, so the BIFF text
  iterator was re-running the same binary-control classification on the fully cleaned whole value
  and again on each cleaned line before emission.
- The BIFF text path still keeps the `looksLikeEmbeddedPDFText(...)` and
  `looksLikeSpreadsheetFormulaExpression(...)` guards, so this change only removes repeated
  binary-control classification work from emitted BIFF text candidates.

Validation:
- `go test -buildvcs=false -run "TestMarkdownTableCell|TestCleanVisibleTextDropsInternalOfficePartReferences|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestBIFFHeaderFooterControlCodesAreNotVisibleText" ./...`
- `go test -buildvcs=false ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-text-redundant-binary-check-removal-xls.json -csv testdata\web-samples\reports\perf-after-biff-text-redundant-binary-check-removal-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-text-redundant-binary-check-removal-keyset.json -csv testdata\web-samples\reports\perf-after-biff-text-redundant-binary-check-removal-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-text-redundant-binary-check-removal-008055-rerun.json -csv testdata\web-samples\reports\perf-after-biff-text-redundant-binary-check-removal-008055-rerun.csv testdata\web-samples\samples\xls\008055.xls`

Anomalous `.xls` sample rerun:
- `008055.xls`: `15241 ms`, text=`17684747`
- `013623.xls`: `3530 ms`, text=`1329843`
- `016161.xls`: `7109 ms`, text=`8179182`
- `018548.xls`: `3859 ms`, text=`1541448`
- `019088.xls`: `19833 ms`, text=`281048`
- `019089.xls`: `17058 ms`, text=`306058`

Key sample verification:
- `00003763.docx`: text=`452979`, images=`10`, `4933 ms`
- `223624.docx`: text=`760288`, images=`33`, `3878 ms`
- `008055.xls`: text=`17684747`, images=`0`, `16089 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `4747 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `7850 ms`

Repeat check:
- `008055.xls`: text=`17684747`, images=`0`, `15203 ms`

Conclusion:
- This low-risk BIFF/XLS cleanup compounds strongly with the earlier cell-path and comment-path
  changes.
- The clearest wins are `008055.xls` reaching about `15.2 s`, `016161.xls` dropping to about
  `7.1 s`, and `019089.xls` reaching about `17.1 s`, while the key sample counts remain
  unchanged.

## 2026-06-30 mojibake rune-threshold scan

- Reworked the long-string branch inside `looksLikeLegacyMojibakeControlLine(...)` to stop
  allocating `[]rune` just to answer the question 鈥渄oes this string have at least 24 runes?鈥?
- The new helper `stringHasAtLeastRunes(...)` scans only until the threshold is met, and the
  mojibake path now reuses one `noHorizontalSpace` check instead of re-running the same
  `strings.ContainsAny(s, " \t")` test around the punctuation-table probes.
- This keeps all existing detection rules intact, but trims allocation and repeated scanning inside
  one of the hottest subpaths of `looksLikeBinaryControlFragment(...)`.

Validation:
- `go test -buildvcs=false -run "TestMarkdownTableCell|TestUniqueStringsDropsHiddenResourceReferences|TestCleanVisibleTextDropsInternalOfficePartReferences|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise" ./...`
- `go test -buildvcs=false ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-mojibake-rune-threshold-scan-xls.json -csv testdata\web-samples\reports\perf-after-mojibake-rune-threshold-scan-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-mojibake-rune-threshold-scan-keyset.json -csv testdata\web-samples\reports\perf-after-mojibake-rune-threshold-scan-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-mojibake-rune-threshold-scan-008055-rerun.json -csv testdata\web-samples\reports\perf-after-mojibake-rune-threshold-scan-008055-rerun.csv testdata\web-samples\samples\xls\008055.xls`

Anomalous `.xls` sample rerun:
- `008055.xls`: `13298 ms`, text=`17684747`
- `013623.xls`: `3057 ms`, text=`1329843`
- `016161.xls`: `6047 ms`, text=`8179182`
- `018548.xls`: `3877 ms`, text=`1541448`
- `019088.xls`: `18557 ms`, text=`281048`
- `019089.xls`: `16674 ms`, text=`306058`

Key sample verification:
- `00003763.docx`: text=`452979`, images=`10`, `4569 ms`
- `223624.docx`: text=`760288`, images=`33`, `3604 ms`
- `008055.xls`: text=`17684747`, images=`0`, `15729 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `4657 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `7374 ms`

Repeat check:
- `008055.xls`: text=`17684747`, images=`0`, `14950 ms`

Conclusion:
- This is a low-risk internal hot-path optimization inside the mojibake classifier, and it
  compounds well with the recent BIFF cleanup series.
- The clearest wins are `008055.xls` moving into the `13.3 s` to `15.0 s` range on reruns,
  `016161.xls` dropping to about `6.0 s`, and `019089.xls` reaching about `16.7 s`, while the key
  sample counts remain unchanged.

## 2026-06-30 unicode noise trimmed reuse

- Tightened `looksLikeUnicodeBinaryNoise(...)` so the function switches to the already-trimmed
  string once `strings.TrimSpace(...)` has been computed and verified non-empty.
- This is a deliberately tiny change: it does not alter any detection rules, but it avoids
  carrying the pre-trimmed string through the later prefix/substring checks in one of the hottest
  binary-control subpaths.
- The attempted larger follow-up edits in this area were discarded; this section records only the
  retained trimmed-string reuse that passed the full regression and performance sweep.

Validation:
- `go test -buildvcs=false -run "TestMarkdownTableCell|TestUniqueStringsDropsHiddenResourceReferences|TestCleanVisibleTextDropsInternalOfficePartReferences|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise" ./...`
- `go test -buildvcs=false ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-unicode-trimmed-reuse-xls.json -csv testdata\web-samples\reports\perf-after-unicode-trimmed-reuse-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-unicode-trimmed-reuse-keyset.json -csv testdata\web-samples\reports\perf-after-unicode-trimmed-reuse-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-unicode-trimmed-reuse-008055-rerun.json -csv testdata\web-samples\reports\perf-after-unicode-trimmed-reuse-008055-rerun.csv testdata\web-samples\samples\xls\008055.xls`

Anomalous `.xls` sample rerun:
- `008055.xls`: `13465 ms`, text=`17684747`
- `013623.xls`: `3291 ms`, text=`1329843`
- `016161.xls`: `6020 ms`, text=`8179182`
- `018548.xls`: `3678 ms`, text=`1541448`
- `019088.xls`: `18430 ms`, text=`281048`
- `019089.xls`: `16751 ms`, text=`306058`

Key sample verification:
- `00003763.docx`: text=`452979`, images=`10`, `4706 ms`
- `223624.docx`: text=`760288`, images=`33`, `3946 ms`
- `008055.xls`: text=`17684747`, images=`0`, `16142 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `4763 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `7552 ms`

Repeat check:
- `008055.xls`: text=`17684747`, images=`0`, `15682 ms`

Conclusion:
- This retained change is tiny but measurably positive on the current long-tail `.xls` set.
- The clearest wins are `008055.xls` reaching about `13.5 s`, `016161.xls` staying around
  `6.0 s`, and `018548.xls` reaching about `3.7 s`, while the key sample counts remain unchanged.

## 2026-06-30 markdown backfill low-information early filter

- Kept the exact-hit containment maps in `markdownBackfillContainment` and added one more safe
  hot-path guard in `missingMarkdownText(...)`: low-information / disallowed normalized lines now
  exit via `markdownBackfillNormalizedLineAllowed(...)` before entering the single-character-run
  collapse path.
- This matters for the long-tail `.xls` samples because the old hotspot was spending most of its
  time collapsing single-character ASCII runs and then checking backfill coverage / containment for
  lines that would be rejected anyway.
- The retained change does not alter the existing filtering rules. It only moves the
  low-information gate earlier so obviously discardable lines skip the expensive collapse and
  containment work.

Validation:
- `go test -buildvcs=false -run 'TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill' -count=1 .`
- `go run -buildvcs=false ./cmd/profileextract testdata\web-samples\samples\xls\019088.xls -cpuprofile testdata\web-samples\reports\cpu-019088-lowinfo-early.pprof`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-continue-xls-anomalies-lowinfo-early.json -csv testdata\web-samples\reports\perf-continue-xls-anomalies-lowinfo-early.csv testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-early.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-early.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30.json | ConvertFrom-Json).inputs)`

Profile check:
- Before this change, `cpu-019088-current.pprof` showed the dominant hotspot inside
  `missingMarkdownText(...)` on the single-character-run collapse branch.
- After the change, `cpu-019088-lowinfo-early.pprof` no longer shows that branch dominating, and
  the same single-file `019088.xls` run dropped to about `5.1 s`.

Target `.xls` rerun:
- `016161.xls`: `7036 ms`, text=`8179182`
- `019088.xls`: `5949 ms`, text=`281048`
- `019089.xls`: `5676 ms`, text=`306058`
- Group summary: total=`18661 ms`, max=`7036 ms`, `over10Sec=0`

Top-30 rerun summary vs `perf-rerun-20260630-top30.json`:
- `.docx`: `10658 ms -> 9362 ms`
- `.xls`: `488637 ms -> 56762 ms`
- `.xlsx`: `50486 ms -> 48356 ms`

Key sample verification:
- `00003763.docx`: text=`452979`, images=`10`, `5245 ms`
- `223624.docx`: text=`760288`, images=`33`, `4117 ms`
- `008055.xls`: text=`17684747`, images=`0`, `18734 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `5237 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `6910 ms`

Largest wins in the rerun set:
- `016161.xls`: `233475 ms -> 8458 ms`
- `019088.xls`: `98412 ms -> 6310 ms`
- `019089.xls`: `105226 ms -> 6192 ms`
- `008055.xls`: `26581 ms -> 18734 ms`

Notes:
- The old top-30 baseline for `016161.xls` reported `textBytes=4939910`, while the new rerun
  reports `textBytes=8179182`. This matches the corrected/stable output seen in the narrower reruns
  during this session, so the earlier top-30 baseline should be treated as stale rather than as a
  regression introduced by this change.
- The broader rerun still has only one `.xls` file above `10 s` (`008055.xls`), and the three
  previously anomalous `.xls` files are now all below `7.1 s`.

Conclusion:
- This is the first retained change in this phase that directly attacks the markdown backfill
  hotspot rather than another downstream BIFF cleanup, and it compounds cleanly with the earlier
  `.xls` optimizations.
- The clearest wins are `019088.xls` and `019089.xls` moving from roughly `100 s` into the `6 s`
  range in the top-30 rerun, `016161.xls` collapsing from `233.5 s` to `8.5 s`, and the top-30
  `.xls` aggregate dropping from `488.6 s` to `56.8 s`.

## 2026-06-30 hidden-office precheck before deep reference classification

- Added a cheap `maybeDiscardableHiddenOfficeText(...)` guard before the deeper hidden-reference /
  Office-metadata classifiers in two hot paths:
  `markdownBackfillNormalizedLineAllowed(...)` and `markdownTableCellValueLineDiscarded(...)`.
- Before this change, ordinary visible lines could still pay for
  `looksLikeHiddenResourceReference(...)`,
  `looksLikeRelationshipIDReference(...)`,
  `looksLikeOfficeRelationshipMetadataReference(...)`, and
  `looksLikeOfficeXMLMetadataReference(...)` even when the line had no hidden-resource markers at
  all.
- The retained change keeps the same deep checks for suspicious lines, but skips them for clearly
  ordinary text, which helps both the BIFF markdown cell cleaner and the markdown backfill pass on
  large `.xls` outputs.

Validation:
- `go test -buildvcs=false -run 'TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-hidden-office-precheck-xls.json -csv testdata\web-samples\reports\perf-after-hidden-office-precheck-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-hidden-office-precheck-keyset.json -csv testdata\web-samples\reports\perf-after-hidden-office-precheck-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-hidden-office-precheck.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-hidden-office-precheck.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-early.json | ConvertFrom-Json).inputs)`

Target `.xls` rerun:
- `008055.xls`: `15039 ms`, text=`17684747`
- `013623.xls`: `3809 ms`, text=`1329843`
- `016161.xls`: `6617 ms`, text=`8179182`
- `018548.xls`: `4470 ms`, text=`1541448`
- `019088.xls`: `4570 ms`, text=`281048`
- `019089.xls`: `4715 ms`, text=`306058`
- Group summary: total=`39220 ms`, max=`15039 ms`, `over10Sec=1`

Top-30 rerun summary vs `perf-rerun-20260630-top30-after-lowinfo-early.json`:
- `.docx`: `9362 ms -> 10009 ms`
- `.xls`: `56762 ms -> 47976 ms`
- `.xlsx`: `48356 ms -> 43172 ms`

Key sample verification:
- `00003763.docx`: text=`452979`, images=`10`, `5406 ms`
- `223624.docx`: text=`760288`, images=`33`, `4603 ms`
- `008055.xls`: text=`17684747`, images=`0`, `16311 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `4549 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `6804 ms`

Notable deltas inside the top-30 rerun:
- `008055.xls`: `18734 ms -> 16311 ms`
- `002505.xls`: `7513 ms -> 6127 ms`
- `016161.xls`: `8458 ms -> 7125 ms`
- `018548.xls`: `5383 ms -> 4375 ms`
- `019088.xls`: `6310 ms -> 4661 ms`
- `019089.xls`: `6192 ms -> 4535 ms`
- `00012389.xlsx`: `5237 ms -> 4549 ms`
- `testRecordSizeExceeded.xlsx`: `6910 ms -> 6804 ms`

Notes:
- A few samples regressed slightly in wall-clock time during this rerun, notably
  `013623.xls` (`4172 ms -> 4842 ms`) and the two key `.docx` files. Their extracted text/image
  counts remain unchanged, and the aggregate top-30 runtime still improves materially.
- The top-30 rerun still has only one `.xls` file above `10 s` (`008055.xls`), and its runtime
  dropped again from `18.7 s` to `16.3 s`.

Conclusion:
- This is a conservative 鈥渁void deep classification unless the line looks suspicious鈥?change, and
  it compounds well with the earlier low-information early filter.
- The strongest gains are another `8.8 s` off the top-30 `.xls` aggregate, `019088.xls` and
  `019089.xls` moving into the mid-`4 s` range, and `008055.xls` dropping into the mid-`15 s` to
  low-`16 s` range while keeping output stable.

## 2026-06-30 non-ASCII guard before unicode noise classifiers

- Added a cheap byte-level `containsNonASCIIByte(...)` precheck in `looksLikeBinaryControlFragment(...)`
  before calling the two heaviest non-ASCII noise classifiers:
  `looksLikeLegacyMojibakeControlLine(...)` and `looksLikeUnicodeBinaryNoise(...)`.
- Those classifiers are only intended to catch mojibake / Unicode binary-noise fragments. For
  ASCII-only lines, they were doing expensive rune conversion and multi-branch classification work
  that could never produce a positive result.
- The retained change keeps all earlier ASCII-path checks exactly as they were, and only skips the
  non-ASCII classifiers when the candidate line contains no non-ASCII bytes at all.

Validation:
- `go test -buildvcs=false -run 'TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-nonascii-guard-xls.json -csv testdata\web-samples\reports\perf-after-nonascii-guard-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-nonascii-guard-keyset.json -csv testdata\web-samples\reports\perf-after-nonascii-guard-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-nonascii-guard.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-nonascii-guard.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-hidden-office-precheck.json | ConvertFrom-Json).inputs)`

Target `.xls` rerun:
- `008055.xls`: `12737 ms`, text=`17684747`
- `013623.xls`: `4117 ms`, text=`1329843`
- `016161.xls`: `5407 ms`, text=`8179182`
- `018548.xls`: `3138 ms`, text=`1541448`
- `019088.xls`: `3486 ms`, text=`281048`
- `019089.xls`: `3526 ms`, text=`306058`
- Group summary: total=`32411 ms`, max=`12737 ms`, `over10Sec=1`

Top-30 rerun summary vs `perf-rerun-20260630-top30-after-hidden-office-precheck.json`:
- `.docx`: `10009 ms -> 8601 ms`
- `.xls`: `47976 ms -> 34402 ms`
- `.xlsx`: `43172 ms -> 32702 ms`

Key sample verification:
- `00003763.docx`: text=`452979`, images=`10`, `4766 ms`
- `223624.docx`: text=`760288`, images=`33`, `3761 ms`
- `008055.xls`: text=`17684747`, images=`0`, `11677 ms`
- `00012389.xlsx`: text=`11148655`, images=`0`, `3641 ms`
- `testRecordSizeExceeded.xlsx`: text=`181333430`, images=`0`, `6379 ms`

Notable deltas inside the top-30 rerun:
- `008055.xls`: `16311 ms -> 12007 ms`
- `002505.xls`: `6127 ms -> 4399 ms`
- `013623.xls`: `4842 ms -> 3563 ms`
- `016161.xls`: `7125 ms -> 5061 ms`
- `018548.xls`: `4375 ms -> 2771 ms`
- `019088.xls`: `4661 ms -> 3028 ms`
- `019089.xls`: `4535 ms -> 3573 ms`
- `00012389.xlsx`: `4549 ms -> 2957 ms`
- `testRecordSizeExceeded.xlsx`: `6804 ms -> 5765 ms`

Conclusion:
- This is a small code change with outsized effect because it prunes a very hot branch from the
  overwhelmingly ASCII-heavy BIFF/XLS workload before any rune-heavy Unicode noise analysis begins.
- The strongest wins are the top-30 `.xls` aggregate dropping from `48.0 s` to `34.4 s`,
  `008055.xls` moving down to about `12.0 s`, `016161.xls` reaching about `5.1 s`, and
  `019088.xls` / `019089.xls` landing around `3.0 s` to `3.6 s` while extracted text and image
  counts remain unchanged.

## 2026-06-30 low-information precheck before `looksLikeLowInformationFragment`

- Added one more cheap guard inside `looksLikeBinaryControlFragment(...)` before the expensive
  `looksLikeLowInformationFragment(...)` fallback:
  skip that fallback when the candidate is shorter than 6 bytes, contains horizontal whitespace, or
  already contains non-ASCII bytes.
- This is semantics-preserving because `looksLikeLowInformationFragment(...)` already returns
  `false` for those cases; the change only avoids sending obviously ineligible inputs through a hot
  byte-scan classifier.
- The effect compounds with the earlier non-ASCII guard because `looksLikeDiscardableBinaryControlLine(...)`
  is still a major cost center in the BIFF-heavy `.xls` path.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-lowinfo-precheck-in-binary-control-xls.json -csv testdata\web-samples\reports\perf-after-lowinfo-precheck-in-binary-control-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-lowinfo-precheck-in-binary-control-keyset.json -csv testdata\web-samples\reports\perf-after-lowinfo-precheck-in-binary-control-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-nonascii-guard.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 10911 ms`
  - `013623.xls`: `4117 ms -> 3425 ms`
  - `016161.xls`: `5407 ms -> 4962 ms`
  - `018548.xls`: `3138 ms -> 3467 ms`
  - `019088.xls`: `3486 ms -> 3295 ms`
  - `019089.xls`: `3526 ms -> 3610 ms`
  - group summary: `32411 ms -> 29670 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `00003763.docx`: `4766 ms -> 5019 ms`
  - `223624.docx`: `3761 ms -> 3431 ms`
  - `008055.xls`: `11677 ms -> 11717 ms`
  - `00012389.xlsx`: `3641 ms -> 3799 ms`
  - `testRecordSizeExceeded.xlsx`: `6379 ms -> 5792 ms`
- Top-30 rerun vs `perf-rerun-20260630-top30-after-nonascii-guard.json`:
  - `.docx`: `8601 ms -> 7993 ms`
  - `.xls`: `34402 ms -> 32513 ms`
  - `.xlsx`: `32702 ms -> 33598 ms`

Conclusion:
- This is a small but real net win across the top-30 rerun, even though a subset of `.xlsx`
  samples regressed modestly.
- The strongest wins are another roughly `1.9 s` off the top-30 `.xls` aggregate, about `0.6 s`
  off the `.docx` aggregate, and `008055.xls` moving down to roughly `11.0 s` while output counts
  remain unchanged.

## 2026-06-30 rejected: BIFF markdown display-value cleaning fast path

- Tried a BIFF-only markdown-table optimization for generated display values: numbers, RK values,
  bool/error values, and formula display values would bypass `cleanMarkdownTableCellValue(...)` and
  use a much smaller cleaner before `prepareMarkdownTableCellValue(...)`.
- This was intentionally scoped so it would affect only legacy BIFF markdown table construction and
  only values already synthesized by the extractor.
- The `.xls` gains were real and correctness stayed stable, but the broader top-30 rerun still
  ended slightly worse overall because `.xlsx` and `.docx` drifted upward more than the `.xls`
  improvement recovered.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-display-clean-fastpath-check-008055.json -csv testdata\web-samples\reports\perf-after-biff-display-clean-fastpath-check-008055.csv testdata\web-samples\samples\xls\008055.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-display-clean-fastpath-xls.json -csv testdata\web-samples\reports\perf-after-biff-display-clean-fastpath-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-display-clean-fastpath-keyset.json -csv testdata\web-samples\reports\perf-after-biff-display-clean-fastpath-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-biff-display-clean-fastpath.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-biff-display-clean-fastpath.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Single-file check:
  - `008055.xls`: text bytes unchanged at `17684747`, `9988 ms`
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 10244 ms`
  - `013623.xls`: `4117 ms -> 3423 ms`
  - `016161.xls`: `5407 ms -> 4890 ms`
  - `018548.xls`: `3138 ms -> 3139 ms`
  - `019088.xls`: `3486 ms -> 3358 ms`
  - `019089.xls`: `3526 ms -> 3222 ms`
  - group summary: `32411 ms -> 28276 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `00003763.docx`: `4766 ms -> 4497 ms`
  - `223624.docx`: `3761 ms -> 3633 ms`
  - `008055.xls`: `11677 ms -> 11124 ms`
  - `00012389.xlsx`: `3641 ms -> 3683 ms`
  - `testRecordSizeExceeded.xlsx`: `6379 ms -> 5692 ms`
- Top-30 incremental rerun vs `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json`:
  - `.docx`: `7993 ms -> 8149 ms`
  - `.xls`: `32513 ms -> 32399 ms`
  - `.xlsx`: `33598 ms -> 34520 ms`

Decision:
- Reverted. The BIFF-local win was not enough to offset the shared top-30 regressions.

## 2026-06-30 rejected: BIFF markdown plain-cell fast path

- Tried adding an early return in `prepareMarkdownTableCellValue(...)` that skipped
  `markdownText(...)` and the subsequent split/dedupe/escape path when the cell looked like plain
  text.
- The idea was sound for isolated `.xls` workloads, and a narrow 6-file `.xls` rerun improved from
  `32411 ms` to `28675 ms`.
- However, the broader top-30 rerun regressed overall wall-clock time versus
  `perf-rerun-20260630-top30-after-nonascii-guard.json`, so this experiment was reverted.

Validation:
- `go test -buildvcs=false -run 'TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestMarkdownTableCell' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-cell-plain-fastpath-xls.json -csv testdata\web-samples\reports\perf-after-biff-cell-plain-fastpath-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-cell-plain-fastpath-keyset.json -csv testdata\web-samples\reports\perf-after-biff-cell-plain-fastpath-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-biff-cell-plain-fastpath.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-biff-cell-plain-fastpath.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-nonascii-guard.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun:
  - `008055.xls`: `12007 ms -> 10823 ms`
  - `013623.xls`: `3563 ms -> 3409 ms`
  - `016161.xls`: `5061 ms -> 4811 ms`
  - `018548.xls`: `2771 ms -> 3139 ms`
  - `019088.xls`: `3028 ms -> 3204 ms`
  - `019089.xls`: `3573 ms -> 3289 ms`
- Top-30 summary:
  - `.docx`: `8601 ms -> 9153 ms`
  - `.xls`: `34402 ms -> 34599 ms`
  - `.xlsx`: `32702 ms -> 35972 ms`

Decision:
- Reverted. The targeted `.xls` gain is not worth the broader `.xlsx` and `.docx` regressions.

## 2026-06-30 rejected: BIFF text reuse between text output and markdown backfill

- Tried reusing the already extracted BIFF workbook text when building `.xls` structured markdown,
  so `biffMarkdown(...)` would not need to call `biffText(...)` a second time just to feed
  `missingMarkdownText(...)`.
- This was intentionally scoped to `.xls` extraction and did produce a strong narrow-file gain on
  the classic BIFF hotspot set.
- However, repeated top-30 reruns did not hold up as an overall win, so the experiment was
  reverted instead of being retained on a theory.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-xls-bifftext-reuse-xls.json -csv testdata\web-samples\reports\perf-after-xls-bifftext-reuse-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-xls-bifftext-reuse-keyset.json -csv testdata\web-samples\reports\perf-after-xls-bifftext-reuse-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-xls-bifftext-reuse.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-xls-bifftext-reuse.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-nonascii-guard.json | ConvertFrom-Json).inputs)`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-xls-bifftext-reuse-rerun2.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-xls-bifftext-reuse-rerun2.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-nonascii-guard.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 9355 ms`
  - `013623.xls`: `4117 ms -> 2633 ms`
  - `016161.xls`: `5407 ms -> 4590 ms`
  - `018548.xls`: `3138 ms -> 2615 ms`
  - `019088.xls`: `3486 ms -> 2795 ms`
  - `019089.xls`: `3526 ms -> 2908 ms`
  - group summary: `32411 ms -> 24896 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `008055.xls`: `11677 ms -> 10573 ms`
  - `00012389.xlsx`: `3641 ms -> 3888 ms`
  - `testRecordSizeExceeded.xlsx`: `6379 ms -> 5934 ms`
- Top-30 rerun 1 vs `perf-rerun-20260630-top30-after-nonascii-guard.json`:
  - `.docx`: `8601 ms -> 8682 ms`
  - `.xls`: `34402 ms -> 32012 ms`
  - `.xlsx`: `32702 ms -> 38651 ms`
- Top-30 rerun 2 vs the same baseline:
  - `.docx`: `8601 ms -> 9223 ms`
  - `.xls`: `34402 ms -> 28403 ms`
  - `.xlsx`: `32702 ms -> 39622 ms`

Decision:
- Reverted. The `.xls` gain was real, but the broader top-30 reruns stayed materially worse on the
  `.xlsx` aggregate across repeated runs.

## 2026-06-30 rejected: empty hidden-state fast path in BIFF hidden-cell checks

- Tried a very small fast path in `biffCellHidden(...)` and `biffBIFFCellRecordHidden(...)` for the
  common case where there are no hidden rows or hidden columns yet.
- The change was deliberately tiny and `.xls`-only, but the measured result was still negative on
  the hotspot set, so it was reverted immediately.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-hidden-empty-fastpath-xls.json -csv testdata\web-samples\reports\perf-after-biff-hidden-empty-fastpath-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-hidden-empty-fastpath-keyset.json -csv testdata\web-samples\reports\perf-after-biff-hidden-empty-fastpath-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 11754 ms`
  - `013623.xls`: `4117 ms -> 4118 ms`
  - `016161.xls`: `5407 ms -> 5374 ms`
  - `018548.xls`: `3138 ms -> 3375 ms`
  - `019088.xls`: `3486 ms -> 3597 ms`
  - `019089.xls`: `3526 ms -> 3622 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `008055.xls`: `11677 ms -> 13119 ms`
  - `00012389.xlsx`: `3641 ms -> 3734 ms`
  - `testRecordSizeExceeded.xlsx`: `6379 ms -> 6120 ms`

Decision:
- Reverted. The supposed fast path did not produce a net gain even on the narrow `.xls` targets.

## 2026-06-30 rejected: remove per-line reclean in `forEachBIFFText`

- Tried simplifying `forEachBIFFText(...)` so that after the initial `cleanVisibleText(s)` call, the
  multi-line branch would no longer call `cleanVisibleText(line)` again for each emitted line.
- This looked promising from the profile because `forEachBIFFText(...)` is hot in BIFF text
  extraction, but the measured result was negative on the `.xls` hotspot set and on the keyset.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-foreach-line-reclean-removal-xls.json -csv testdata\web-samples\reports\perf-after-biff-foreach-line-reclean-removal-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-foreach-line-reclean-removal-keyset.json -csv testdata\web-samples\reports\perf-after-biff-foreach-line-reclean-removal-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 11805 ms`
  - `013623.xls`: `4117 ms -> 3566 ms`
  - `016161.xls`: `5407 ms -> 5474 ms`
  - `018548.xls`: `3138 ms -> 3428 ms`
  - `019088.xls`: `3486 ms -> 3795 ms`
  - `019089.xls`: `3526 ms -> 3653 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `008055.xls`: `11677 ms -> 12632 ms`
  - `00012389.xlsx`: `3641 ms -> 3867 ms`
  - `testRecordSizeExceeded.xlsx`: `6379 ms -> 6246 ms`

Decision:
- Reverted. The removed per-line reclean did not translate into a net performance win.

## 2026-06-30 rejected: direct-append BIFF display values in text extraction

- Tried bypassing `forEachBIFFText(...)` for BIFF cell display values that looked intrinsically safe:
  numeric cells, RK cells, formula display values, and bool/error display values.
- The performance signal was initially attractive on `.xls`, but this changed extracted output size on
  the hotspot workbook, which means the existing `forEachBIFFText(...)` path is still doing
  user-visible normalization even for these apparently simple display strings.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-display-directtext-xls.json -csv testdata\web-samples\reports\perf-after-biff-display-directtext-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-display-directtext-keyset.json -csv testdata\web-samples\reports\perf-after-biff-display-directtext-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun immediately showed output drift:
  - `008055.xls`: text bytes `17684747 -> 19558754`, `9397 ms`
  - `013623.xls`: text bytes `1329843 -> 1329934`, `3213 ms`
  - `016161.xls`: text bytes `8179182 -> 9264536`, `4561 ms`
  - `018548.xls`: text bytes `1541448 -> 1712046`, `2679 ms`
- Keyset showed the same correctness problem:
  - `008055.xls`: text bytes `17684747 -> 19558754`, `10139 ms`
  - `00012389.xlsx`: text bytes unchanged, `3447 ms`

Decision:
- Reverted immediately. This was not a semantics-preserving optimization.

## 2026-06-30 rejected: single-line `cleanVisibleText` hidden-office early return

- Tried adding a fast return in the single-line branch of `cleanVisibleText(...)` when
  `maybeDiscardableHiddenOfficeText(...)` said the line did not resemble hidden Office metadata.
- The intent was to skip `stripInlineHiddenOfficeReferences(...)` and the subsequent discardable-line
  classification for the common visible-text case.
- Despite looking conservative, this changed extracted `.xls` output size on the hotspot sample, so
  it was reverted immediately.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-cleanvisible-singleline-hiddenoffice-guard-xls.json -csv testdata\web-samples\reports\perf-after-cleanvisible-singleline-hiddenoffice-guard-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-cleanvisible-singleline-hiddenoffice-guard-keyset.json -csv testdata\web-samples\reports\perf-after-cleanvisible-singleline-hiddenoffice-guard-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun showed correctness drift immediately:
  - `008055.xls`: text bytes `17684747 -> 19558790`, `11164 ms`
  - `013623.xls`: text bytes `1329843 -> 1332839`, `3636 ms`
  - `016161.xls`: text bytes `8179182 -> 9265368`, `5499 ms`
  - `018548.xls`: text bytes `1541448 -> 1712127`, `3323 ms`
- Keyset showed the same drift:
  - `008055.xls`: text bytes `17684747 -> 19558790`, `12207 ms`
  - `00012389.xlsx`: text bytes unchanged, `3803 ms`

Decision:
- Reverted immediately. This guard is not semantics-preserving for legacy `.xls` extraction.

## 2026-06-30 rejected: byte-scan helper for `maybeHiddenOrControlText`

- Replaced the `strings.ContainsAny(s, "/\\\\<>#%")` check inside `maybeHiddenOrControlText(...)`
  with a tiny byte-scan helper.
- This was semantics-preserving in the tested outputs and looked promising on the `.xls` hotspot set
  and the small cross-format keyset.
- However, the broader top-30 rerun still regressed overall because the shared-path `.xlsx`
  aggregate moved the wrong way, so this change was reverted rather than kept as a local win.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-hidden-control-marker-byte-scan-xls.json -csv testdata\web-samples\reports\perf-after-hidden-control-marker-byte-scan-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-hidden-control-marker-byte-scan-keyset.json -csv testdata\web-samples\reports\perf-after-hidden-control-marker-byte-scan-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-hidden-control-marker-byte-scan.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-hidden-control-marker-byte-scan.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-nonascii-guard.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 10840 ms`
  - `013623.xls`: `4117 ms -> 3459 ms`
  - `016161.xls`: `5407 ms -> 4863 ms`
  - `018548.xls`: `3138 ms -> 3175 ms`
  - `019088.xls`: `3486 ms -> 3486 ms`
  - `019089.xls`: `3526 ms -> 3477 ms`
  - group summary: `32411 ms -> 29300 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `008055.xls`: `11677 ms -> 11467 ms`
  - `00012389.xlsx`: `3641 ms -> 3623 ms`
  - `testRecordSizeExceeded.xlsx`: `6379 ms -> 6037 ms`
- Top-30 rerun vs `perf-rerun-20260630-top30-after-nonascii-guard.json`:
  - `.docx`: `8601 ms -> 8525 ms`
  - `.xls`: `34402 ms -> 33364 ms`
  - `.xlsx`: `32702 ms -> 35081 ms`

Decision:
- Reverted. The `.xls` gain was real, but the shared-path `.xlsx` regression made the overall
  top-30 result worse.

## 2026-06-30 rejected: BIFF worksheet map preallocation

- Tried modest preallocation inside `biffWorksheetMarkdownTables(...)` for the per-sheet `rows` and
  `hiddenRows` maps, using a size hint derived from workbook byte length.
- This was intentionally limited to legacy `.xls` worksheet markdown construction and did preserve
  extracted output exactly on the measured sets.
- The narrow `.xls` hotspot batch improved, but the broader cross-format keyset regressed, so the
  change was reverted instead of being promoted to a larger rerun.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*' -count=1 .`
- `go test -buildvcs=false ./...`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-worksheet-map-prealloc-xls.json -csv testdata\web-samples\reports\perf-after-biff-worksheet-map-prealloc-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-worksheet-map-prealloc-keyset.json -csv testdata\web-samples\reports\perf-after-biff-worksheet-map-prealloc-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 11345 ms`
  - `013623.xls`: `4117 ms -> 3600 ms`
  - `016161.xls`: `5407 ms -> 5229 ms`
  - `018548.xls`: `3138 ms -> 3588 ms`
  - `019088.xls`: `3486 ms -> 3847 ms`
  - `019089.xls`: `3526 ms -> 3704 ms`
  - group summary: `32411 ms -> 31313 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 8981 ms`
  - `.xls`: `11677 ms -> 12209 ms`
  - `.xlsx`: `10020 ms -> 10457 ms`

Decision:
- Reverted. The local `.xls` win was not strong enough to justify a consistent regression on the
  cross-format keyset.

## 2026-06-30 rejected: direct `MULRK` iteration in BIFF markdown cells

- Tried removing the temporary `[]string` allocation in `biffSetMarkdownMulRKCells(...)` by
  iterating the `MULRK` record directly and immediately calling `biffSetMarkdownCellAt(...)` for
  each decoded RK value.
- The change was narrowly scoped to legacy `.xls` markdown table extraction and preserved
  `textBytes` on every measured sample.
- Early narrow runs looked encouraging, but the comparison that matters against the latest retained
  baseline was negative, so the change was reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-markdown-mulrk-direct-xls.json -csv testdata\web-samples\reports\perf-after-biff-markdown-mulrk-direct-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-markdown-mulrk-direct-keyset.json -csv testdata\web-samples\reports\perf-after-biff-markdown-mulrk-direct-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-markdown-mulrk-direct-keyset-rerun2.json -csv testdata\web-samples\reports\perf-after-biff-markdown-mulrk-direct-keyset-rerun2.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-biff-markdown-mulrk-direct.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-biff-markdown-mulrk-direct.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 11634 ms`
  - `013623.xls`: `4117 ms -> 3586 ms`
  - `016161.xls`: `5407 ms -> 5220 ms`
  - `018548.xls`: `3138 ms -> 3377 ms`
  - `019088.xls`: `3486 ms -> 3451 ms`
  - `019089.xls`: `3526 ms -> 3619 ms`
  - group summary: `32411 ms -> 30887 ms`
- Keyset rerun 2 vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 8455 ms`
  - `.xls`: `11677 ms -> 11331 ms`
  - `.xlsx`: `10020 ms -> 9214 ms`
- Top-30 rerun vs `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json`:
  - `.docx`: `7993 ms -> 8229 ms`
  - `.xls`: `32513 ms -> 33744 ms`
  - `.xlsx`: `33598 ms -> 35009 ms`

Decision:
- Reverted. A seemingly good local allocation reduction did not survive the top-30 rerun against
  the latest retained baseline.

## 2026-06-30 rejected: conservative fast path for `markdownBackfillComparableText`

- Tried adding a very narrow early return in `markdownBackfillComparableText(...)` for trimmed,
  plain ASCII lines that had no markdown punctuation, escaping, HTML markers, repeated spaces, or
  list/heading prefixes.
- The intent was to avoid repeatedly running `markdownBackfillVisibleText(...)`,
  `markdownVisibleLineText(...)`, and `strings.Fields(...)` for the large number of ordinary lines
  that already appear normalized.
- The change preserved `textBytes` and looked strong on the narrow `.xls` hotspot batch, but it did
  not hold up on the top-30 rerun, so it was reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-comparable-fastpath-xls.json -csv testdata\web-samples\reports\perf-after-comparable-fastpath-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-comparable-fastpath-keyset.json -csv testdata\web-samples\reports\perf-after-comparable-fastpath-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-comparable-fastpath-keyset-rerun2.json -csv testdata\web-samples\reports\perf-after-comparable-fastpath-keyset-rerun2.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-comparable-fastpath.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-comparable-fastpath.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 11432 ms`
  - `013623.xls`: `4117 ms -> 3435 ms`
  - `016161.xls`: `5407 ms -> 5047 ms`
  - `018548.xls`: `3138 ms -> 3211 ms`
  - `019088.xls`: `3486 ms -> 3429 ms`
  - `019089.xls`: `3526 ms -> 3330 ms`
  - group summary: `32411 ms -> 29884 ms`
- Keyset rerun 2 vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 8825 ms`
  - `.xls`: `11677 ms -> 11175 ms`
  - `.xlsx`: `10020 ms -> 9165 ms`
- Top-30 rerun vs `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json`:
  - `.docx`: `7993 ms -> 8388 ms`
  - `.xls`: `32513 ms -> 33822 ms`
  - `.xlsx`: `33598 ms -> 35080 ms`

Decision:
- Reverted. The shared-path win on sampled workloads did not survive the broader top-30 rerun.

## 2026-06-30 rejected: hide-flag hidden-column filtering in BIFF text extraction

- Tried changing `biffText(...)` so hidden-column removal would mark matching `biffTextPart`
  entries as `hide=true` instead of compacting the `parts` slice in place.
- The motivation was to avoid repeated slice rebuilding on legacy `.xls` workbooks with late hidden
  columns while keeping row index bookkeeping stable for subsequent hidden-row filtering.
- Hidden-row/hidden-column correctness tests still passed and `textBytes` stayed unchanged, but the
  cross-format keyset remained too mixed to justify keeping the change.

Validation:
- `go test -buildvcs=false -run 'TestLegacyXLSLateHiddenRowDoesNotLeakText|TestLegacyXLSLateHiddenColumnDoesNotLeakTextOrComments|TestLegacyXLSCommentsAreVisibleAndHiddenCellsAreFiltered|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-hidden-column-hideflag-xls.json -csv testdata\web-samples\reports\perf-after-biff-hidden-column-hideflag-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-hidden-column-hideflag-keyset.json -csv testdata\web-samples\reports\perf-after-biff-hidden-column-hideflag-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-hidden-column-hideflag-keyset-rerun2.json -csv testdata\web-samples\reports\perf-after-biff-hidden-column-hideflag-keyset-rerun2.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 11901 ms`
  - `013623.xls`: `4117 ms -> 3642 ms`
  - `016161.xls`: `5407 ms -> 5119 ms`
  - `018548.xls`: `3138 ms -> 3376 ms`
  - `019088.xls`: `3486 ms -> 3488 ms`
  - `019089.xls`: `3526 ms -> 3524 ms`
  - group summary: `32411 ms -> 31050 ms`
- Keyset rerun 2 vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 9061 ms`
  - `.xls`: `11677 ms -> 11579 ms`
  - `.xlsx`: `10020 ms -> 9665 ms`

Decision:
- Reverted. The `.xls`-only benefit was too small and the keyset remained inconsistent.

## 2026-06-30 rejected: `biffText` row-index map preallocation

- Tried preallocating the `rowPartIndexes` and `hiddenRows` maps inside `biffText(...)`, with a
  modest workbook-size-derived hint reused at BOF/EOF sheet resets.
- This was intended to shave map growth overhead from the legacy `.xls` text path without touching
  shared markdown normalization logic.
- The result was immediately negative even on the narrow `.xls` hotspot batch, so the change was
  reverted without spending a top-30 rerun on it.

Validation:
- `go test -buildvcs=false -run 'TestLegacyXLSLateHiddenRowDoesNotLeakText|TestLegacyXLSLateHiddenColumnDoesNotLeakTextOrComments|TestLegacyXLSCommentsAreVisibleAndHiddenCellsAreFiltered|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-rowindex-map-prealloc-xls.json -csv testdata\web-samples\reports\perf-after-biff-rowindex-map-prealloc-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-rowindex-map-prealloc-keyset.json -csv testdata\web-samples\reports\perf-after-biff-rowindex-map-prealloc-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 12998 ms`
  - `013623.xls`: `4117 ms -> 3840 ms`
  - `016161.xls`: `5407 ms -> 5525 ms`
  - `018548.xls`: `3138 ms -> 3562 ms`
  - `019088.xls`: `3486 ms -> 3869 ms`
  - `019089.xls`: `3526 ms -> 3808 ms`
  - group summary: `32411 ms -> 33602 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 10452 ms`
  - `.xls`: `11677 ms -> 12975 ms`
  - `.xlsx`: `10020 ms -> 10410 ms`

Decision:
- Reverted immediately. This preallocation strategy made the hotspot path slower.

## 2026-06-30 rejected: small-string cache for `biffText` normalization

- Tried adding a local cache in `biffText(...)` for short raw strings so repeated BIFF cell values
  could reuse the output of `forEachBIFFText(...)` instead of repeatedly running
  `cleanVisibleText(...)`, line splitting, and filtering.
- The idea targeted repeated numeric and short textual display values in legacy `.xls` workbooks and
  preserved `textBytes` on the sampled runs.
- In practice, the extra cache bookkeeping was materially slower on the hotspot set, so the change
  was reverted immediately.

Validation:
- `go test -buildvcs=false -run 'TestLegacyXLSLateHiddenRowDoesNotLeakText|TestLegacyXLSLateHiddenColumnDoesNotLeakTextOrComments|TestLegacyXLSCommentsAreVisibleAndHiddenCellsAreFiltered|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-text-small-cache-xls.json -csv testdata\web-samples\reports\perf-after-biff-text-small-cache-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-text-small-cache-keyset.json -csv testdata\web-samples\reports\perf-after-biff-text-small-cache-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 13839 ms`
  - `013623.xls`: `4117 ms -> 3602 ms`
  - `016161.xls`: `5407 ms -> 6147 ms`
  - `018548.xls`: `3138 ms -> 2851 ms`
  - `019088.xls`: `3486 ms -> 3602 ms`
  - `019089.xls`: `3526 ms -> 3889 ms`
  - group summary: `32411 ms -> 33930 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 9204 ms`
  - `.xls`: `11677 ms -> 14174 ms`
  - `.xlsx`: `10020 ms -> 9983 ms`

Decision:
- Reverted immediately. The cache overhead outweighed any reuse benefit on the measured workloads.

## 2026-06-30 rejected: deduplicating `missingMarkdownText` per-line variant checks

- Tried collapsing the repeated per-line checks in `missingMarkdownText(...)` by building a deduped
  variant list from `line`, `visibleLine`, `markdownLine`, and the escaped-table visible form, then
  running the coverage and containment predicates once per unique variant.
- The idea was to reduce obviously repeated work inside the hottest part of markdown backfill
  without changing the underlying predicate implementations.
- In practice this was not semantics-preserving for wide-sheet markdown backfill, and it was also
  dramatically slower on the hotspot `.xls` sample, so it was reverted immediately.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-missing-line-variant-dedupe-xls.json -csv testdata\web-samples\reports\perf-after-missing-line-variant-dedupe-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-missing-line-variant-dedupe-keyset.json -csv testdata\web-samples\reports\perf-after-missing-line-variant-dedupe-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Regression test failure:
  - `TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns` failed because the sheet markdown
    started backfilling truncated table values that the stable logic correctly suppresses.
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 16298 ms`
  - `013623.xls`: `4117 ms -> 4833 ms`
  - `016161.xls`: `5407 ms -> 9438 ms`
  - `018548.xls`: `3138 ms -> 4076 ms`
  - `019088.xls`: `3486 ms -> 3878 ms`
  - `019089.xls`: `3526 ms -> 3749 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 10136 ms`
  - `.xls`: `11677 ms -> 19129 ms`
  - `.xlsx`: `10020 ms -> 10367 ms`

Decision:
- Reverted immediately. This changed behavior and made performance much worse.

## 2026-06-30 rejected: cache only short candidate lines in markdown backfill

- Tried narrowing `missingMarkdownText(...)`'s `candidateLineCache` so only raw lines of length 64
  or less would be cached; longer lines would be normalized on demand and not inserted into the map.
- The intent was to cut `mapassign` churn from one-off long lines while preserving cache hits for
  repeated short labels and simple values.
- This was semantics-preserving in the targeted regression suite, and the narrow `.xls` batch looked
  good, but the top-30 rerun was decisively worse than the retained baseline, so the change was
  reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-candidate-cache-short-only-xls.json -csv testdata\web-samples\reports\perf-after-candidate-cache-short-only-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-candidate-cache-short-only-keyset.json -csv testdata\web-samples\reports\perf-after-candidate-cache-short-only-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-candidate-cache-short-only-keyset-rerun2.json -csv testdata\web-samples\reports\perf-after-candidate-cache-short-only-keyset-rerun2.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-candidate-cache-short-only.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-candidate-cache-short-only.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 11706 ms`
  - `013623.xls`: `4117 ms -> 3645 ms`
  - `016161.xls`: `5407 ms -> 5161 ms`
  - `018548.xls`: `3138 ms -> 3439 ms`
  - `019088.xls`: `3486 ms -> 3617 ms`
  - `019089.xls`: `3526 ms -> 3621 ms`
  - group summary: `32411 ms -> 31189 ms`
- Keyset rerun 2 vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 8761 ms`
  - `.xls`: `11677 ms -> 11750 ms`
  - `.xlsx`: `10020 ms -> 9169 ms`
- Top-30 rerun vs `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json`:
  - `.docx`: `7993 ms -> 8737 ms`
  - `.xls`: `32513 ms -> 37193 ms`
  - `.xlsx`: `33598 ms -> 38246 ms`

Decision:
- Reverted. The narrower cache policy did not survive the broader top-30 rerun.

## 2026-06-30 rejected: cache only skipped lines in markdown backfill

- Tried changing `missingMarkdownText(...)` so the per-line cache would remember only lines known
  not to need backfill, while lines already added to `missing` would rely solely on `seenMissing`
  for duplicate suppression.
- The reasoning was that positive cache entries are redundant once `seenMissing[line] = true`, so
  removing them might cut some `mapassign` and lookup overhead in the hottest loop.
- The change was semantics-preserving in the targeted regression suite, and the narrow `.xls` batch
  looked mildly positive, but the keyset moved clearly in the wrong direction, so it was reverted
  without spending a top-30 rerun.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-skipped-line-cache-only-xls.json -csv testdata\web-samples\reports\perf-after-skipped-line-cache-only-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-skipped-line-cache-only-keyset.json -csv testdata\web-samples\reports\perf-after-skipped-line-cache-only-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 11968 ms`
  - `013623.xls`: `4117 ms -> 3778 ms`
  - `016161.xls`: `5407 ms -> 5223 ms`
  - `018548.xls`: `3138 ms -> 3491 ms`
  - `019088.xls`: `3486 ms -> 3776 ms`
  - `019089.xls`: `3526 ms -> 3693 ms`
  - group summary: `32411 ms -> 31929 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 9589 ms`
  - `.xls`: `11677 ms -> 12326 ms`
  - `.xlsx`: `10020 ms -> 10243 ms`

Decision:
- Reverted. The keyset regression was too clear to justify continuing.

## 2026-06-30 rejected: short plain-text early return in binary control detection

- Tried adding an early `false` return near the top of `looksLikeBinaryControlFragment(...)` for
  short ASCII strings that had neither hidden/control markers nor a positive
  `maybeControlFragmentText(...)` match.
- The goal was to cut repeated binary-control classification work for very short ordinary text
  lines in `cleanVisibleText(...)` and markdown cell filtering, while still allowing known short
  control tokens like `chart` and marker-bearing strings such as `%PDF-` or `<...>`.
- The change was semantics-preserving in the targeted regression suite, but it regressed both the
  narrow `.xls` batch and the cross-format keyset, so it was reverted immediately.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-binary-short-plain-guard-xls.json -csv testdata\web-samples\reports\perf-after-binary-short-plain-guard-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-binary-short-plain-guard-keyset.json -csv testdata\web-samples\reports\perf-after-binary-short-plain-guard-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 13144 ms`
  - `013623.xls`: `4117 ms -> 3739 ms`
  - `016161.xls`: `5407 ms -> 5457 ms`
  - `018548.xls`: `3138 ms -> 3513 ms`
  - `019088.xls`: `3486 ms -> 3390 ms`
  - `019089.xls`: `3526 ms -> 3428 ms`
  - group summary: `32411 ms -> 32671 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 10430 ms`
  - `.xls`: `11677 ms -> 12922 ms`
  - `.xlsx`: `10020 ms -> 10155 ms`

Decision:
- Reverted immediately. This early return made the measured workloads slower overall.

## 2026-06-30 rejected: `<br>` guard in `markdownVisibleTableCells`

- Tried wrapping the segment-splitting branch in `markdownVisibleTableCells(...)` so
  `strings.Split(part, "<br>")` would run only when the already-normalized cell text actually
  contained `<br>`.
- The change was intentionally tiny and semantics-preserving in the targeted markdown/backfill
  regression suite, because cells without `<br>` only ever produced a single redundant segment equal
  to `part`, which the existing code immediately discarded.
- Narrow `.xls` and the second keyset rerun looked acceptable, but the top-30 rerun against the
  retained baseline regressed across all three format groups, so the change was reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-tablecell-br-guard-xls.json -csv testdata\web-samples\reports\perf-after-tablecell-br-guard-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-tablecell-br-guard-keyset.json -csv testdata\web-samples\reports\perf-after-tablecell-br-guard-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-tablecell-br-guard-keyset-rerun2.json -csv testdata\web-samples\reports\perf-after-tablecell-br-guard-keyset-rerun2.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-tablecell-br-guard.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-tablecell-br-guard.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 12630 ms`
  - `013623.xls`: `4117 ms -> 3663 ms`
  - `016161.xls`: `5407 ms -> 5263 ms`
  - `018548.xls`: `3138 ms -> 3339 ms`
  - `019088.xls`: `3486 ms -> 3413 ms`
  - `019089.xls`: `3526 ms -> 3397 ms`
  - group summary: `32411 ms -> 31705 ms`
- Keyset rerun 2 vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 8742 ms`
  - `.xls`: `11677 ms -> 11581 ms`
  - `.xlsx`: `10020 ms -> 9696 ms`
- Top-30 rerun vs `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json`:
  - `.docx`: `7993 ms -> 8649 ms`
  - `.xls`: `32513 ms -> 34126 ms`
  - `.xlsx`: `33598 ms -> 35352 ms`

Decision:
- Reverted. A small local win in the sampled sets did not survive the broader top-30 rerun.

## 2026-06-30 rejected: BIFF NUMBER direct append in `biffText`

- Tried a very narrow fast path in `biffText(...)` for BIFF `0x0203` NUMBER records only.
- The experiment bypassed `forEachBIFFText(...)` for numeric display strings and appended the raw
  `biffNumberDisplayValue(...)` directly as a single cell text part, while leaving BOOLERR,
  formula strings, shared strings, and markdown generation unchanged.
- The goal was to shave time from the hottest `biffText` branch without repeating the broader
  direct-append experiment that had already changed output semantics elsewhere.

Validation:
- `go test -buildvcs=false -run 'TestLegacyXLSLateHiddenRowDoesNotLeakText|TestLegacyXLSLateHiddenColumnDoesNotLeakTextOrComments|TestLegacyXLSCommentsAreVisibleAndHiddenCellsAreFiltered|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-number-direct-append-xls.json -csv testdata\web-samples\reports\perf-after-biff-number-direct-append-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `textBytes 17684747 -> 19557230`, `12737 ms -> 9498 ms`
  - `013623.xls`: `textBytes 1329843 -> 1329843`, `4117 ms -> 3364 ms`
  - `016161.xls`: `textBytes 8179182 -> 9254740`, `5407 ms -> 4735 ms`
  - `018548.xls`: `textBytes 1541448 -> 1708438`, `3138 ms -> 2975 ms`
  - `019088.xls`: `textBytes 281048 -> 281048`, `3486 ms -> 3487 ms`
  - `019089.xls`: `textBytes 306058 -> 306058`, `3526 ms -> 3438 ms`
  - group summary: `textBytes 29322326 -> 32437357`, `32411 ms -> 27497 ms`

Decision:
- Reverted immediately. Although it sped up the sampled `.xls` batch, it changed extracted text
  volume on multiple files, so it was not semantically safe enough to continue to keyset or top-30.

## 2026-06-30 rejected: one-pass `prepareMarkdownTableCellValue`

- Tried replacing the `Split -> TrimSpace/dedupe -> Join -> ReplaceAll` sequence in
  `prepareMarkdownTableCellValue(...)` with a single-pass builder.
- The intended win was lower allocation churn while preserving the exact same markdown-cell output:
  trim per line, drop empties, collapse consecutive duplicates, then escape `\` and `|` and turn
  newlines into `<br>`.
- The targeted regression suite and the narrow `.xls` batch both kept `textBytes` stable, but the
  broader keyset regressed enough across all three formats that the change was not worth keeping.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-markdown-cell-prepare-onepass-xls.json -csv testdata\web-samples\reports\perf-after-markdown-cell-prepare-onepass-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-markdown-cell-prepare-onepass-keyset.json -csv testdata\web-samples\reports\perf-after-markdown-cell-prepare-onepass-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 12338 ms`
  - `013623.xls`: `4117 ms -> 3618 ms`
  - `016161.xls`: `5407 ms -> 5466 ms`
  - `018548.xls`: `3138 ms -> 3369 ms`
  - `019088.xls`: `3486 ms -> 3711 ms`
  - `019089.xls`: `3526 ms -> 3503 ms`
  - group summary: `32411 ms -> 32005 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 9407 ms`
  - `.xls`: `11677 ms -> 13988 ms`
  - `.xlsx`: `10020 ms -> 11453 ms`

Decision:
- Reverted. This looked mildly positive in the narrow `.xls` sample but regressed the broader
  keyset too clearly.

## 2026-06-30 rejected: empty-raw fast path in `missingMarkdownText` candidate caching

- Tried adding a very small fast path in `missingMarkdownText(...)` so the local `candidateLine`
  closure would return immediately for raw empty strings instead of probing and populating
  `candidateLineCache`.
- The intent was to shave map overhead in the hottest per-line loop without changing normalization,
  filtering, or cache behavior for any non-empty line.
- This one actually survived targeted tests, the narrow `.xls` batch, and the cross-format keyset
  with stable `textBytes`, but it still regressed badly enough on the retained top-30 workload that
  it was not kept.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-candidate-empty-raw-fastpath-xls.json -csv testdata\web-samples\reports\perf-after-candidate-empty-raw-fastpath-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-candidate-empty-raw-fastpath-keyset.json -csv testdata\web-samples\reports\perf-after-candidate-empty-raw-fastpath-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-candidate-empty-raw-fastpath.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-candidate-empty-raw-fastpath.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 11679 ms`
  - `013623.xls`: `4117 ms -> 3620 ms`
  - `016161.xls`: `5407 ms -> 5498 ms`
  - `018548.xls`: `3138 ms -> 3504 ms`
  - `019088.xls`: `3486 ms -> 3414 ms`
  - `019089.xls`: `3526 ms -> 3307 ms`
  - group summary: `32411 ms -> 31022 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 7772 ms`
  - `.xls`: `11677 ms -> 11131 ms`
  - `.xlsx`: `10020 ms -> 9256 ms`
- Top-30 rerun vs `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json`:
  - `.docx`: `7993 ms -> 7814 ms`
  - `.xls`: `32513 ms -> 34925 ms`
  - `.xlsx`: `33598 ms -> 36917 ms`

Decision:
- Reverted. The broader top-30 workload remains the deciding signal, and this change lost there.

## 2026-06-30 rejected: marker-free fast path in `markdownBackfillVisibleText`

- Tried a small fast path in `markdownBackfillVisibleText(...)`: when the input line had none of
  `[`, `!`, or `\`, it returned `markdownText(...)` directly instead of first running
  `markdownInlineLinkVisibleText(...)` plus the unescape pass.
- The intent was to avoid repeated reference-definition scanning and inline-link processing for the
  very common case where a backfill candidate line does not contain markdown link or escape markers.
- This change was semantics-preserving in the regression suite and kept `textBytes` stable across
  the sampled perf batches, but it still regressed the retained top-30 workload enough that it was
  not kept.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-visible-fastpath-xls.json -csv testdata\web-samples\reports\perf-after-backfill-visible-fastpath-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-visible-fastpath-keyset.json -csv testdata\web-samples\reports\perf-after-backfill-visible-fastpath-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260630-top30-after-backfill-visible-fastpath.json -csv testdata\web-samples\reports\perf-rerun-20260630-top30-after-backfill-visible-fastpath.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 10907 ms`
  - `013623.xls`: `4117 ms -> 3562 ms`
  - `016161.xls`: `5407 ms -> 4733 ms`
  - `018548.xls`: `3138 ms -> 3117 ms`
  - `019088.xls`: `3486 ms -> 3265 ms`
  - `019089.xls`: `3526 ms -> 3403 ms`
  - group summary: `32411 ms -> 28987 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 8137 ms`
  - `.xls`: `11677 ms -> 10923 ms`
  - `.xlsx`: `10020 ms -> 9097 ms`
- Top-30 rerun vs `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json`:
  - `.docx`: `7993 ms -> 7903 ms`
  - `.xls`: `32513 ms -> 32601 ms`
  - `.xlsx`: `33598 ms -> 35603 ms`

Decision:
- Reverted. The sampled sets looked good, but the broader top-30 rerun still regressed overall.

## 2026-06-30 rejected: preallocate `markdownBackfillCoverageSet` map

- Tried preallocating the `coverage` map in `markdownBackfillCoverageSet(...)` using an estimate
  derived from the markdown line count.
- The goal was to reduce rehashing and bucket growth inside the hottest coverage-building path,
  especially because `addMarkdownBackfillCoverageText(...)` spends a noticeable amount of time on
  repeated `coverage[text] = struct{}{}` writes.
- The change was semantics-preserving in regression tests and mildly positive on the narrow `.xls`
  batch, but the broader keyset regressed enough that it was reverted before any top-30 rerun.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-coverage-map-prealloc-xls.json -csv testdata\web-samples\reports\perf-after-coverage-map-prealloc-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-coverage-map-prealloc-keyset.json -csv testdata\web-samples\reports\perf-after-coverage-map-prealloc-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 12030 ms`
  - `013623.xls`: `4117 ms -> 3730 ms`
  - `016161.xls`: `5407 ms -> 5315 ms`
  - `018548.xls`: `3138 ms -> 3424 ms`
  - `019088.xls`: `3486 ms -> 3649 ms`
  - `019089.xls`: `3526 ms -> 3597 ms`
  - group summary: `32411 ms -> 31745 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 8994 ms`
  - `.xls`: `11677 ms -> 12491 ms`
  - `.xlsx`: `10020 ms -> 11749 ms`

Decision:
- Reverted. The keyset regression was too clear to justify keeping the preallocation.

## 2026-07-01 rejected: guard the escaped-table-cell visible-text branch

- Tried narrowing the expensive `markdownBackfillVisibleText(escapeMarkdownTableCell(line))`
  variant check inside `missingMarkdownText(...)`.
- The idea was to reuse `visibleLine` unless the candidate line actually contained table-escaping
  characters such as `|` or `\`, because the escaped-cell path is intended to match markdown table
  cell coverage and looked redundant for ordinary single-line text.
- That assumption was wrong for at least one wide-sheet backfill case: the change caused a real
  markdown regression before performance evaluation could even continue.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`

Observed results:
- Regression failure:
  - `TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns/49609.xlsx`
- The failure showed that the escaped-table-cell path is still semantically relevant even when the
  raw candidate line itself does not obviously contain the table-escape markers that motivated the
  optimization.

Decision:
- Reverted immediately after the regression failure. No perf follow-up was kept.

## 2026-07-01 rejected: `[` guard in `markdownReferenceDefinitions`

- Tried adding a trivial early return in `markdownReferenceDefinitions(...)` when the input text did
  not contain `[` at all.
- The reasoning was straightforward: without `[` there cannot be markdown reference definitions, so
  the function could skip the line scan and return an empty map immediately.
- Although the change was semantics-preserving in regression tests, it regressed the narrow `.xls`
  batch enough that it was reverted without spending time on keyset or top-30 reruns.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-refdefs-bracket-guard-xls.json -csv testdata\web-samples\reports\perf-after-refdefs-bracket-guard-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 13335 ms`
  - `013623.xls`: `4117 ms -> 4028 ms`
  - `016161.xls`: `5407 ms -> 5672 ms`
  - `018548.xls`: `3138 ms -> 3068 ms`
  - `019088.xls`: `3486 ms -> 3348 ms`
  - `019089.xls`: `3526 ms -> 3888 ms`
  - group summary: `32411 ms -> 33339 ms`

Decision:
- Reverted. Even this tiny early return lost on the narrow `.xls` sample.

## 2026-07-01 rejected: one-pass whitespace collapse in `markdownBackfillComparableText`

- Tried replacing `strings.Join(strings.Fields(s), " ")` at the end of
  `markdownBackfillComparableText(...)` with a custom single-pass whitespace-collapsing helper.
- The goal was to keep the same comparable-text normalization while avoiding the slice allocation
  and second pass implied by `Fields + Join`.
- Regression tests passed and the narrow `.xls` batch improved, but the keyset regressed badly
  enough across all three format groups that the change was reverted immediately.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-comparable-whitespace-collapse-xls.json -csv testdata\web-samples\reports\perf-after-comparable-whitespace-collapse-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-comparable-whitespace-collapse-keyset.json -csv testdata\web-samples\reports\perf-after-comparable-whitespace-collapse-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 12072 ms`
  - `013623.xls`: `4117 ms -> 3905 ms`
  - `016161.xls`: `5407 ms -> 4739 ms`
  - `018548.xls`: `3138 ms -> 3554 ms`
  - `019088.xls`: `3486 ms -> 3550 ms`
  - `019089.xls`: `3526 ms -> 3393 ms`
  - group summary: `32411 ms -> 31213 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 12115 ms`
  - `.xls`: `11677 ms -> 12674 ms`
  - `.xlsx`: `10020 ms -> 10552 ms`

Decision:
- Reverted. The keyset regression was too strong to justify continuing.

## 2026-07-01 rejected: early return in `markdownInlineLinkVisibleText`

- Tried a very local optimization in `markdownInlineLinkVisibleText(...)`: if the input string did
  not contain `[` at all, return it unchanged instead of allocating a builder and copying bytes.
- This was intentionally narrower than the earlier rejected fast path in
  `markdownBackfillVisibleText(...)`, because the downstream unescape and `markdownText(...)` stages
  still ran exactly as before.
- Even so, the narrow `.xls` sample regressed clearly enough that the change was reverted without
  continuing to keyset or top-30.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-inline-link-no-bracket-return-xls.json -csv testdata\web-samples\reports\perf-after-inline-link-no-bracket-return-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 12351 ms`
  - `013623.xls`: `4117 ms -> 4238 ms`
  - `016161.xls`: `5407 ms -> 5811 ms`
  - `018548.xls`: `3138 ms -> 3570 ms`
  - `019088.xls`: `3486 ms -> 4494 ms`
  - `019089.xls`: `3526 ms -> 3706 ms`
  - group summary: `32411 ms -> 34170 ms`

Decision:
- Reverted. The narrow `.xls` sample already made it clear this change was not helping.

## 2026-07-01 rejected: reuse visible text for comparable table-row normalization

- Tried removing duplicate work inside `markdownBackfillContainmentSet(...)` for table rows.
- The specific change was to compute `visible := markdownBackfillVisibleText(markdownLine)` once,
  then derive the comparable form from that visible text instead of calling
  `markdownBackfillComparableText(markdownLine)`, which internally ran
  `markdownBackfillVisibleText(...)` again on the same source line.
- This was meant to preserve semantics while avoiding one repeated visible-text normalization pass
  per markdown table row.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-comparable-visible-reuse-xls.json -csv testdata\web-samples\reports\perf-after-comparable-visible-reuse-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-comparable-visible-reuse-keyset.json -csv testdata\web-samples\reports\perf-after-comparable-visible-reuse-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260701-top30-after-comparable-visible-reuse.json -csv testdata\web-samples\reports\perf-rerun-20260701-top30-after-comparable-visible-reuse.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 11426 ms`
  - `013623.xls`: `4117 ms -> 3193 ms`
  - `016161.xls`: `5407 ms -> 4818 ms`
  - `018548.xls`: `3138 ms -> 2645 ms`
  - `019088.xls`: `3486 ms -> 3400 ms`
  - `019089.xls`: `3526 ms -> 3234 ms`
  - group summary: `32411 ms -> 28716 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 8624 ms`
  - `.xls`: `11677 ms -> 12780 ms`
  - `.xlsx`: `10020 ms -> 8810 ms`
- Top-30 rerun vs `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json`:
  - `.docx`: `7993 ms -> 8741 ms`
  - `.xls`: `32513 ms -> 34447 ms`
  - `.xlsx`: `33598 ms -> 34755 ms`

Decision:
- Reverted. Despite a strong narrow `.xls` result, the broader top-30 rerun regressed across every
  format group.

## 2026-07-01 rejected: early `lineMissingCache` hit for non-single-character lines

- Tried moving one cache check earlier inside `missingMarkdownText(...)`.
- The change let non-single-character lines reuse an existing `lineMissingCache[line]` result before
  re-running `markdownBackfillNormalizedLineAllowed(...)` and the surrounding duplicate work.
- Single-character lines were intentionally excluded so the existing single-character-run collapse
  logic stayed untouched.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-early-nonsingle-line-cache-xls.json -csv testdata\web-samples\reports\perf-after-early-nonsingle-line-cache-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-early-nonsingle-line-cache-keyset.json -csv testdata\web-samples\reports\perf-after-early-nonsingle-line-cache-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 12130 ms`
  - `013623.xls`: `4117 ms -> 4000 ms`
  - `016161.xls`: `5407 ms -> 5573 ms`
  - `018548.xls`: `3138 ms -> 3200 ms`
  - `019088.xls`: `3486 ms -> 3556 ms`
  - `019089.xls`: `3526 ms -> 3592 ms`
  - group summary: `32411 ms -> 32051 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 8712 ms`
  - `.xls`: `11677 ms -> 11951 ms`
  - `.xlsx`: `10020 ms -> 10041 ms`

Decision:
- Reverted. The narrow `.xls` sample looked slightly positive, but the cross-format keyset still
  lost overall.

## 2026-07-01 rejected: remove `candidateLineCache` from `missingMarkdownText`

- Tried deleting the local `candidateLineCache` entirely and calling
  `markdownBackfillCandidateLine(...)` directly at each use site in `missingMarkdownText(...)`.
- The motivation came straight from the stable profile: cache lookup and assignment around
  `candidateLine` were taking more cumulative time than the normalization helper itself, suggesting
  that repeated map traffic might be more expensive than recomputing candidate lines.
- The change was semantics-preserving in regression tests and looked strong on the narrow `.xls`
  sample, but it still regressed the keyset once `.xlsx` was included.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-remove-candidate-line-cache-xls.json -csv testdata\web-samples\reports\perf-after-remove-candidate-line-cache-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-remove-candidate-line-cache-keyset.json -csv testdata\web-samples\reports\perf-after-remove-candidate-line-cache-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 10226 ms`
  - `013623.xls`: `4117 ms -> 3354 ms`
  - `016161.xls`: `5407 ms -> 5435 ms`
  - `018548.xls`: `3138 ms -> 3938 ms`
  - `019088.xls`: `3486 ms -> 4228 ms`
  - `019089.xls`: `3526 ms -> 3945 ms`
  - group summary: `32411 ms -> 31126 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 8457 ms`
  - `.xls`: `11677 ms -> 11482 ms`
  - `.xlsx`: `10020 ms -> 10781 ms`

Decision:
- Reverted. Removing the cache looked promising on `.xls`, but the keyset still lost overall once
  `.xlsx` was included.

## 2026-07-01 rejected: adaptive `candidateLineCache` enablement

- Tried making the `candidateLine` cache in `missingMarkdownText(...)` adaptive instead of always
  on.
- The experiment started without a cache, sampled the first 256 raw lines, and only enabled the
  full cache if repeated raw-line keys crossed a small duplicate threshold. The hope was to get the
  `.xls` win from uncached normalization while still preserving cache benefits for more repetitive
  `.xlsx` workloads.
- Regression tests passed, the narrow `.xls` batch looked strong, and the cross-format keyset also
  improved overall, but the broader top-30 rerun regressed badly and the change was reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-adaptive-candidate-line-cache-xls.json -csv testdata\web-samples\reports\perf-after-adaptive-candidate-line-cache-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-adaptive-candidate-line-cache-keyset.json -csv testdata\web-samples\reports\perf-after-adaptive-candidate-line-cache-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260701-top30-after-adaptive-candidate-line-cache.json -csv testdata\web-samples\reports\perf-rerun-20260701-top30-after-adaptive-candidate-line-cache.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 10176 ms`
  - `013623.xls`: `4117 ms -> 3381 ms`
  - `016161.xls`: `5407 ms -> 4696 ms`
  - `018548.xls`: `3138 ms -> 2948 ms`
  - `019088.xls`: `3486 ms -> 3219 ms`
  - `019089.xls`: `3526 ms -> 3143 ms`
  - group summary: `32411 ms -> 27563 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 9097 ms`
  - `.xls`: `11677 ms -> 10116 ms`
  - `.xlsx`: `10020 ms -> 9120 ms`
  - total summary: `30224 ms -> 28333 ms`
- Top-30 rerun vs `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json`:
  - `.docx`: `7993 ms -> 19965 ms`
  - `.xls`: `32513 ms -> 45927 ms`
  - `.xlsx`: `33598 ms -> 34037 ms`

Decision:
- Reverted. Even though this was the first recent candidate to beat keyset, the top-30 rerun was a
  clear regression overall.

## 2026-07-01 rejected: `.xls`-only removal of `candidateLineCache`

- Tried splitting `missingMarkdownText(...)` behind a small option so BIFF markdown backfill could
  run with `candidateLineCache` disabled while the default path for other formats kept the cache.
- The idea matched the recent measurements: narrow `.xls` samples often got faster when candidate
  lines were recomputed directly, while `.xlsx` tended to benefit from the cache. This experiment
  tested whether a format-specific switch could keep both wins.
- Regression tests passed, the narrow `.xls` batch improved, and the small cross-format keyset also
  improved overall. Even so, the broader top-30 rerun still regressed across every format group, so
  the change was reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-xls-only-no-candidate-cache-xls.json -csv testdata\web-samples\reports\perf-after-xls-only-no-candidate-cache-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-xls-only-no-candidate-cache-keyset.json -csv testdata\web-samples\reports\perf-after-xls-only-no-candidate-cache-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260701-top30-after-xls-only-no-candidate-cache.json -csv testdata\web-samples\reports\perf-rerun-20260701-top30-after-xls-only-no-candidate-cache.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 10810 ms`
  - `013623.xls`: `4117 ms -> 3348 ms`
  - `016161.xls`: `5407 ms -> 5307 ms`
  - `018548.xls`: `3138 ms -> 3800 ms`
  - `019088.xls`: `3486 ms -> 3644 ms`
  - `019089.xls`: `3526 ms -> 3609 ms`
  - group summary: `32411 ms -> 30518 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 9795 ms`
  - `.xls`: `11677 ms -> 10469 ms`
  - `.xlsx`: `10020 ms -> 8747 ms`
  - total summary: `30224 ms -> 29011 ms`
- Top-30 rerun vs `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json`:
  - `.docx`: `7993 ms -> 9004 ms`
  - `.xls`: `32513 ms -> 34175 ms`
  - `.xlsx`: `33598 ms -> 35577 ms`

Decision:
- Reverted. The format-specific switch looked attractive on the narrow `.xls` batch and even on the
  smaller keyset, but the retained top-30 baseline still moved the wrong way across all three
  format groups.

## 2026-07-01 rejected: BIFF exact-match workbook-backfill precheck

- Tried adding a BIFF-only precheck in `biffMarkdown(...)` before calling
  `missingMarkdownText(...)`.
- The experiment built a cheap exact-match set from the structured BIFF markdown and skipped the
  expensive workbook backfill pass whenever every `biffText(...)` line looked directly covered by
  that exact set.
- The motivation came from the hotspot sample `008055.xls`: it does not emit `## Workbook Text` in
  the final markdown, so the full backfill pass appears to spend a lot of time proving that nothing
  is missing.
- The approach stayed semantically safe on the known BIFF regression cases, including a new guard
  test that ensures visible rows past the markdown row limit still reach `## Workbook Text`.

Validation:
- `go test -buildvcs=false -run 'TestLegacyXLSMarkdownUsesBIFFCellGrid|TestLegacyXLSWorkbookTextBackfillsVisibleRowsPastMarkdownLimit|TestLegacyXLSMarkdownKeepsEmptyVisibleSheetsAndSkipsHiddenEmptySheets|TestLegacyXLSCommentsAreVisibleAndHiddenCellsAreFiltered|TestLegacyXLSLateHiddenRowDoesNotLeakText|TestLegacyXLSLateHiddenColumnDoesNotLeakTextOrComments|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSOldBIFFDoesNotEmitUnicodeMojibake|TestLegacyXLSBIFF8CompressedUnicodeKeepsLatin1Characters|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSSampleSheetNamesAreStructured' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-exact-backfill-precheck-008055.json -csv testdata\web-samples\reports\perf-after-biff-exact-backfill-precheck-008055.csv testdata\web-samples\samples\xls\008055.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-exact-backfill-precheck-xls.json -csv testdata\web-samples\reports\perf-after-biff-exact-backfill-precheck-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-biff-exact-backfill-precheck-keyset-rerun.json -csv testdata\web-samples\reports\perf-after-biff-exact-backfill-precheck-keyset-rerun.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260701-top30-after-biff-exact-backfill-precheck.json -csv testdata\web-samples\reports\perf-rerun-20260701-top30-after-biff-exact-backfill-precheck.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260701-top30-after-biff-exact-backfill-precheck-rerun2.json -csv testdata\web-samples\reports\perf-rerun-20260701-top30-after-biff-exact-backfill-precheck-rerun2.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Single-file hotspot rerun:
  - `008055.xls`: `12737 ms -> 12532 ms`
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 10373 ms`
  - `013623.xls`: `4117 ms -> 4359 ms`
  - `016161.xls`: `5407 ms -> 4949 ms`
  - `018548.xls`: `3138 ms -> 2860 ms`
  - `019088.xls`: `3486 ms -> 3222 ms`
  - `019089.xls`: `3526 ms -> 3369 ms`
  - group summary: `32411 ms -> 29132 ms`
- Keyset rerun vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 8857 ms`
  - `.xls`: `11677 ms -> 9716 ms`
  - `.xlsx`: `10020 ms -> 10335 ms`
  - total summary: `30224 ms -> 28908 ms`
- Top-30 rerun #1 vs `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json`:
  - `.docx`: `7993 ms -> 8859 ms`
  - `.xls`: `32513 ms -> 29126 ms`
  - `.xlsx`: `33598 ms -> 37809 ms`
  - total summary: `74104 ms -> 75794 ms`
- Top-30 rerun #2 vs the same baseline:
  - `.docx`: `7993 ms -> 10846 ms`
  - `.xls`: `32513 ms -> 30148 ms`
  - `.xlsx`: `33598 ms -> 44769 ms`
  - total summary: `74104 ms -> 85763 ms`

Decision:
- Reverted. The BIFF-only precheck consistently helped `.xls`, but the broader top-30 reruns were
  not stable enough to justify retaining it, and both reruns regressed overall because unrelated
  groups drifted sharply upward.

## 2026-07-01 rejected: generic exact-match early precheck in `missingMarkdownText`

- Tried moving the exact-match early check down into `missingMarkdownText(...)` itself instead of
  keeping it BIFF-specific.
- The new version normalized input lines once, built a markdown exact-match set up front, and
  returned early before constructing `markdownBackfillCoverageSet(...)` and
  `markdownBackfillContainmentSet(...)` when every allowed candidate line looked directly covered by
  the current markdown.
- This directly targeted the hotspot pattern from `008055.xls`: the final markdown has no
  `## Workbook Text`, so the expensive backfill path is often spending its time proving that
  nothing is missing.
- The approach was regression-safe and materially faster on the narrow `.xls` set, but broad top-30
  reruns still regressed overall once `.docx` and `.xlsx` were included, so the change was
  reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-generic-exact-backfill-precheck-008055.json -csv testdata\web-samples\reports\perf-after-generic-exact-backfill-precheck-008055.csv testdata\web-samples\samples\xls\008055.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-generic-exact-backfill-precheck-xls.json -csv testdata\web-samples\reports\perf-after-generic-exact-backfill-precheck-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-generic-exact-backfill-precheck-keyset.json -csv testdata\web-samples\reports\perf-after-generic-exact-backfill-precheck-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260701-top30-after-generic-exact-backfill-precheck.json -csv testdata\web-samples\reports\perf-rerun-20260701-top30-after-generic-exact-backfill-precheck.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Single-file hotspot rerun:
  - `008055.xls`: `12737 ms -> 9045 ms`
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 8770 ms`
  - `013623.xls`: `4117 ms -> 3299 ms`
  - `016161.xls`: `5407 ms -> 3854 ms`
  - `018548.xls`: `3138 ms -> 2892 ms`
  - `019088.xls`: `3486 ms -> 2739 ms`
  - `019089.xls`: `3526 ms -> 3083 ms`
  - group summary: `32411 ms -> 24637 ms`
- Keyset vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 9164 ms`
  - `.xls`: `11677 ms -> 9097 ms`
  - `.xlsx`: `10020 ms -> 9723 ms`
  - total summary: `30224 ms -> 27984 ms`
- Top-30 rerun vs `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json`:
  - `.docx`: `7993 ms -> 8858 ms`
  - `.xls`: `32513 ms -> 29057 ms`
  - `.xlsx`: `33598 ms -> 37941 ms`
  - total summary: `74104 ms -> 75856 ms`

Decision:
- Reverted. The generic early precheck was the strongest recent `.xls` result and even beat the
  keyset overall, but the broader top-30 rerun still regressed because the gains did not survive
  cross-format aggregate validation.

## 2026-07-01 retained after baseline recalibration: generic exact-match early precheck in `missingMarkdownText`

- Revisited the same exact-match early precheck after rerunning the **current stable** top-30
  baseline in the same environment.
- That rerun showed the original comparison target from
  `perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json` was no longer a fair
  same-environment baseline for go/no-go decisions. The current stable rerun was substantially
  slower across all three format groups, so the earlier 鈥渞egression鈥?call was measuring against a
  stale reference point.
- Against the freshly rerun stable baseline, the exact-match early precheck wins cleanly across
  `.docx`, `.xls`, and `.xlsx`, with the largest gain still coming from BIFF backfill avoidance on
  the hotspot `.xls` samples.

Implementation:
- `missingMarkdownText(...)` now normalizes candidate lines once up front and builds a lightweight
  exact-match markdown set before the expensive coverage/containment structures.
- If every allowed candidate line is already directly covered by markdown-visible text, table-cell
  text, or escaped-table visible text, the function returns early without building
  `markdownBackfillCoverageSet(...)` or `markdownBackfillContainmentSet(...)`.
- If any candidate line is not proven by the exact set, the function falls back to the original full
  logic unchanged.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-generic-exact-backfill-precheck-008055.json -csv testdata\web-samples\reports\perf-after-generic-exact-backfill-precheck-008055.csv testdata\web-samples\samples\xls\008055.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-generic-exact-backfill-precheck-xls.json -csv testdata\web-samples\reports\perf-after-generic-exact-backfill-precheck-xls.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-generic-exact-backfill-precheck-keyset-rerun2.json -csv testdata\web-samples\reports\perf-after-generic-exact-backfill-precheck-keyset-rerun2.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260701-top30-current-stable-rerun.json -csv testdata\web-samples\reports\perf-rerun-20260701-top30-current-stable-rerun.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260701-top30-after-generic-exact-backfill-precheck-rerun2.json -csv testdata\web-samples\reports\perf-rerun-20260701-top30-after-generic-exact-backfill-precheck-rerun2.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260630-top30-after-lowinfo-precheck-in-binary-control.json | ConvertFrom-Json).inputs)`

Observed results:
- Single-file hotspot rerun:
  - `008055.xls`: `12737 ms -> 9045 ms`
- Narrow `.xls` rerun vs `perf-after-nonascii-guard-xls.json`:
  - `008055.xls`: `12737 ms -> 8770 ms`
  - `013623.xls`: `4117 ms -> 3299 ms`
  - `016161.xls`: `5407 ms -> 3854 ms`
  - `018548.xls`: `3138 ms -> 2892 ms`
  - `019088.xls`: `3486 ms -> 2739 ms`
  - `019089.xls`: `3526 ms -> 3083 ms`
  - group summary: `32411 ms -> 24637 ms`
- Keyset rerun #2 vs `perf-after-nonascii-guard-keyset.json`:
  - `.docx`: `8527 ms -> 9231 ms`
  - `.xls`: `11677 ms -> 9724 ms`
  - `.xlsx`: `10020 ms -> 10346 ms`
  - total summary: `30224 ms -> 29301 ms`
- Current stable top-30 rerun vs exact-precheck top-30 rerun #2:
  - `.docx`: `9062 ms -> 8838 ms`
  - `.xls`: `37772 ms -> 28695 ms`
  - `.xlsx`: `41168 ms -> 38167 ms`
  - total summary: `88002 ms -> 75700 ms`

Decision:
- Retained.
- This supersedes the earlier rejection of the same idea. The previous rollback decision was based on
  comparison against an older, faster top-30 baseline rather than the freshly rerun current stable
  baseline. Once recalibrated, the change is a clear net win.

## 2026-07-01 rejected: exact-match guard tightening on retained early precheck

- Tried tightening the retained exact-match early precheck so the exact path would skip expensive
  visible-text/table-cell transformations unless a line actually contained markdown/HTML/table
  markers that could change the visible text.
- The implementation added small guards around `markdownVisibleLineText(...)`,
  `markdownVisibleTableCells(...)`, `markdownBackfillVisibleText(...)`, and escaped-table-cell
  variants used by the exact-match path.
- This version was deliberately conservative and regression-safe, and it improved the hotspot
  `008055.xls` sample further, but the broader retained top-30 comparison only moved by noise-level
  amounts and `.xls` as a group actually got slightly worse.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-exact-guard-tightening-008055.json -csv testdata\web-samples\reports\perf-after-exact-guard-tightening-008055.csv testdata\web-samples\samples\xls\008055.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-exact-guard-tightening-keyset.json -csv testdata\web-samples\reports\perf-after-exact-guard-tightening-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-rerun-20260701-top30-after-exact-guard-tightening.json -csv testdata\web-samples\reports\perf-rerun-20260701-top30-after-exact-guard-tightening.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260701-top30-current-stable-rerun.json | ConvertFrom-Json).inputs)`
- `go test -buildvcs=false ./...`

Observed results:
- Single-file hotspot rerun vs retained exact-precheck build:
  - `008055.xls`: `8759 ms -> 8499 ms`
- Keyset vs retained exact-precheck build:
  - `.docx`: `9231 ms -> 7988 ms`
  - `.xls`: `9724 ms -> 8424 ms`
  - `.xlsx`: `10346 ms -> 9028 ms`
  - total summary: `29301 ms -> 25440 ms`
- Top-30 rerun vs retained exact-precheck build:
  - `.docx`: `8838 ms -> 7910 ms`
  - `.xls`: `28695 ms -> 29420 ms`
  - `.xlsx`: `38167 ms -> 38134 ms`
  - total summary: `75700 ms -> 75464 ms`

Decision:
- Reverted. Even though the hotspot and keyset looked great, the retained top-30 comparison only
  improved by `236 ms` overall and `.xls` itself regressed slightly, which is too thin to justify
  keeping another layer of path-specific complexity.

## 2026-07-01 rejected: single-entry repeat cache before `candidateLineCache`

- Tried adding a one-entry 鈥渓ast raw line / last normalized line鈥?cache in front of
  `missingMarkdownText(...)`'s existing `candidateLineCache`.
- The idea was simple: if the BIFF/text stream contains long runs of identical adjacent raw lines,
  the hot exact-backfill path could avoid both map lookup and map assignment for the common repeated
  case.
- In practice the hotspot workbook did not exhibit enough immediate raw-line repetition for the
  extra branch and state tracking to pay for themselves.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-last-candidate-repeat-cache-008055.json -csv testdata\web-samples\reports\perf-after-last-candidate-repeat-cache-008055.csv testdata\web-samples\samples\xls\008055.xls`

Observed results:
- Single-file hotspot rerun vs retained exact-precheck build:
  - `008055.xls`: `8759 ms -> 10925 ms`

Decision:
- Reverted immediately. The extra repeat-cache branch was a clear loss on the hotspot sample.

## 2026-07-01 rejected: cheap hidden-reference gate in `markdownTableCellValueLineDiscarded`

- Tried replacing the `maybeDiscardableHiddenOfficeText(...)` gate in
  `markdownTableCellValueLineDiscarded(...)` with a cheaper marker/keyword precheck before the
  existing exact hidden-reference classifiers.
- The idea was to short-circuit obvious plain table-cell text before paying for
  `maybeHiddenOrControlText(...)`, `containsRIDFold(...)`, and related hidden-office heuristics in
  the BIFF markdown-cell hot path.
- The narrow regression suite stayed green, but the hotspot workbook got materially slower, which
  means the extra precheck was not actually cheaper on the real `.xls` distribution.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-hidden-check-gate-008055.json -csv testdata\web-samples\reports\perf-exp-hidden-check-gate-008055.csv testdata\web-samples\samples\xls\008055.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-hidden-check-gate-008055-rerun.json -csv testdata\web-samples\reports\perf-exp-hidden-check-gate-008055-rerun.csv testdata\web-samples\samples\xls\008055.xls`

Observed results:
- Single-file hotspot reruns vs retained exact-precheck build:
  - `008055.xls`: `8759 ms -> 10961 ms`
  - `008055.xls` rerun: `8759 ms -> 10632 ms`

Decision:
- Reverted immediately. This gate added branch and substring work without reducing the real hidden
  reference workload enough to pay for itself.

## 2026-07-01 rejected: reorder `maybeControlFragmentText` cheap prefix cases ahead of generic classifiers

- Tried moving the large first-character switch in `maybeControlFragmentText(...)` ahead of the more
  generic `looksLikeOLEIdentifierFragment(...)`, `looksLikeOLEWrapperStreamName(...)`, and
  `looksLikeOOXMLMarkupNameFragment(...)` checks.
- The intent was purely ordering-based: preserve the same match set, but let obvious BIFF/OLE
  control strings short-circuit before paying for the more expensive generic classifiers.
- The narrow regression suite stayed green, but the same-machine `.xls` batch comparison showed the
  reordered version was slower overall, so the branch ordering was reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-maybe-control-reorder-008055.json -csv testdata\web-samples\reports\perf-exp-maybe-control-reorder-008055.csv testdata\web-samples\samples\xls\008055.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-maybe-control-reorder-008055-rerun.json -csv testdata\web-samples\reports\perf-exp-maybe-control-reorder-008055-rerun.csv testdata\web-samples\samples\xls\008055.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-maybe-control-reorder-xls6.json -csv testdata\web-samples\reports\perf-exp-maybe-control-reorder-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-baseline-maybe-control-reorder-xls6.json -csv testdata\web-samples\reports\perf-baseline-maybe-control-reorder-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- Single-file hotspot measurements were noisy and inconclusive:
  - experiment: `10490 ms`, rerun `11402 ms`
  - reverted baseline reruns: `14391 ms`, `11206 ms`
- Same-machine 6-file `.xls` batch settled the result:
  - reordered version: `31649 ms`
  - reverted baseline: `30184 ms`

Decision:
- Reverted. The broader `.xls` batch showed the reorder was a net loss even though some individual
  hotspot runs looked promising.

## 2026-07-01 retained: skip markdown list-marker stripping when a line cannot be a list item

- Added a tiny guard in `looksLikeDiscardableBinaryControlLine(...)` so it only calls
  `stripMarkdownListMarker(...)` when the already-trimmed line actually starts with a plausible
  markdown list marker: `-`, `*`, `+`, or a digit.
- This preserves the old behavior for real list items while avoiding repeated list-marker parsing on
  the much more common ordinary BIFF visible-text lines.
- The change is intentionally narrow and only affects the discardable binary-control classifier; the
  rest of the visible-text pipeline is unchanged.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestCleanTextKeepsMarkdownListRowsWithCyrillicText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-list-marker-guard-008055.json -csv testdata\web-samples\reports\perf-exp-list-marker-guard-008055.csv testdata\web-samples\samples\xls\008055.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-list-marker-guard-xls6.json -csv testdata\web-samples\reports\perf-exp-list-marker-guard-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-list-marker-guard-keyset.json -csv testdata\web-samples\reports\perf-exp-list-marker-guard-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-baseline-list-marker-guard-keyset.json -csv testdata\web-samples\reports\perf-baseline-list-marker-guard-keyset.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-list-marker-guard-top30.json -csv testdata\web-samples\reports\perf-exp-list-marker-guard-top30.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260701-top30-current-stable-rerun.json | ConvertFrom-Json).inputs)`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-baseline-list-marker-guard-top30-samemachine.json -csv testdata\web-samples\reports\perf-baseline-list-marker-guard-top30-samemachine.csv @((Get-Content -Raw testdata\web-samples\reports\perf-rerun-20260701-top30-current-stable-rerun.json | ConvertFrom-Json).inputs)`

Observed results:
- Single-file hotspot:
  - `008055.xls`: `10472 ms`
- 6-file `.xls` batch:
  - experiment total: `29367 ms`
  - same-machine stable baseline from the immediately preceding `.xls` batch: `30184 ms`
- Same-machine keyset:
  - experiment: `.docx 11378 ms`, `.xls 11108 ms`, `.xlsx 11478 ms`, total `33964 ms`
  - baseline: `.docx 10962 ms`, `.xls 11033 ms`, `.xlsx 13637 ms`, total `35632 ms`
- Same-machine top-30:
  - experiment: `.docx 10891 ms`, `.xls 35785 ms`, `.xlsx 52458 ms`, total `99134 ms`
  - baseline: `.docx 11491 ms`, `.xls 38909 ms`, `.xlsx 61888 ms`, total `112288 ms`

Decision:
- Kept. This guard consistently improves the `.xls`-heavy workloads and stays positive through the
  same-machine keyset and top-30 reruns.

## 2026-07-01 rejected: tighten `markdownCouldStartListMarker` to fully validate list syntax

- Tried making `markdownCouldStartListMarker(...)` more exact by verifying the full markdown list
  prefix shape before calling `stripMarkdownListMarker(...)`:
  bullet markers had to be followed by ASCII space, and numeric markers had to scan through the
  full digit run and confirm a trailing `.`/`)` plus space.
- The idea was to avoid false positives on ordinary BIFF text that merely begins with `-` or a
  digit, especially dates and ID-like tokens.
- In practice the extra byte scanning cost more than the saved `stripMarkdownListMarker(...)` calls
  on the `.xls` hot path.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestCleanTextKeepsMarkdownListRowsWithCyrillicText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-list-marker-guard-tightened-xls6.json -csv testdata\web-samples\reports\perf-exp-list-marker-guard-tightened-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained broad guard:
  - retained broad guard total: `29367 ms`
  - tightened validator: `38461 ms`

Decision:
- Reverted immediately. The stronger validator was much more expensive than the saved work.

## 2026-07-01 rejected: reuse repeated marker booleans in `maybeInlineHiddenOfficeReferenceMarker`

- Tried hoisting a few repeated scans in `maybeInlineHiddenOfficeReferenceMarker(...)` into cached
  booleans, mainly `containsRIDFold(...)`, `containsASCIIFold(..., "url(")`, and
  `containsASCIIFold(..., "[content_types].xml")`.
- The change was purely structural and kept the same decision tree, with the goal of reducing
  duplicate substring searches in the hot inline-hidden-reference precheck.
- On the real `.xls` batch this lost noticeably, so the extra cached checks were reverted.

Validation:
- `go test -buildvcs=false -run 'TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleTextStripsInlineHiddenOfficeReferences|TestMarkdownImageAltStripsInlineHiddenOfficeReferences|TestLegacyDOCTextStripsInlineHiddenOfficeReferences|TestLegacyXLSTextStripsInlineHiddenOfficeReferences|TestMarkdownTableCellDropsInternalReferences|TestCleanVisibleTextKeepsRelationshipModeWordsInProse|TestLongVisibleTextWithOfficeMetadataWordsIsKept|TestLongOfficeMetadataReferencesAreStillRecognized|TestCleanVisibleTextDropsOfficeXMLNamespaceMetadata|TestMarkdownBackfillTreats.*|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-inline-marker-bool-reuse-xls6.json -csv testdata\web-samples\reports\perf-exp-inline-marker-bool-reuse-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - boolean-reuse variant: `34950 ms`

Decision:
- Reverted immediately. The duplicated scans were cheaper than eagerly computing all of the cached
  booleans on the common path.

## 2026-07-01 rejected: BIFF-only ASCII list-marker helper for discardable binary-control lines

- Tried replacing the generic `stripMarkdownListMarker(...)` call inside
  `looksLikeDiscardableBinaryControlLine(...)` with a BIFF-specific ASCII helper that assumes the
  line has already passed through `cleanText(...)`.
- The helper mirrored the same bullet, ordered-list, and task-list rules, but used ASCII-space
  checks and a narrower parsing path to reduce generic markdown helper overhead on the `.xls`
  hot path.
- Even with the narrower scope, the specialized helper was slightly slower than the retained broad
  guard plus generic list-marker stripping.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestCleanTextKeepsMarkdownListRowsWithCyrillicText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-discardable-list-helper-xls6.json -csv testdata\web-samples\reports\perf-exp-discardable-list-helper-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - BIFF-only ASCII helper: `30845 ms`

Decision:
- Reverted. The specialized helper added complexity without beating the retained generic path.

## 2026-07-01 rejected: trimmed-input helper for `looksLikeDiscardableBinaryControlLine`

- Tried splitting `looksLikeDiscardableBinaryControlLine(...)` into a tiny wrapper plus a
  `looksLikeDiscardableBinaryControlLineTrimmed(...)` helper, then routing the hot call sites from
  `cleanText(...)` and `looksLikeDiscardableVisibleTextLine(...)` directly to the trimmed helper.
- The reasoning was straightforward: those callers already feed trimmed strings, so the extra
  `strings.TrimSpace(...)` in the common binary-control path looked redundant.
- On the real `.xls` batch this turned out to be a loss, so the helper was reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestCleanTextKeepsMarkdownListRowsWithCyrillicText|TestCleanVisibleText.*' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-trimmed-binary-line-helper-xls6.json -csv testdata\web-samples\reports\perf-exp-trimmed-binary-line-helper-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - trimmed-helper variant: `31751 ms`

Decision:
- Reverted. The removed trim was not worth the extra helper split on the measured hot path.

## 2026-07-01 rejected: skip whole-value PDF/formula checks for multiline `forEachBIFFText`

- Tried narrowing `forEachBIFFText(...)` so the expensive whole-value
  `looksLikeEmbeddedPDFText(...)` and `looksLikeSpreadsheetFormulaExpression(...)` checks would run
  only on the single-line path.
- The multiline branch still kept its existing per-line checks after `cleanVisibleText(...)`, so
  the idea was to remove what looked like redundant whole-string classification for already-split
  BIFF text.
- The targeted regression suite stayed green, but the `.xls` batch regressed slightly, so the
  change was reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Markdown|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestStripInlineHiddenOfficeReferencesFastGuard|TestCleanVisibleText.*|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-multiline-wholecheck-skip-xls6.json -csv testdata\web-samples\reports\perf-exp-biff-multiline-wholecheck-skip-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - multiline whole-check skip variant: `30625 ms`

Decision:
- Reverted. The whole-value check still appears to be a net win once the full multiline path is
  accounted for.

## 2026-07-01 rejected: use `cleanTextNoMojibakeRepair` for multiline `cleanVisibleText` tail cleanup

- Tried replacing the final `cleanText(strings.Join(out, "\n"))` in the multiline branch of
  `cleanVisibleText(...)` with `cleanTextNoMojibakeRepair(...)`.
- The rationale was that each individual line has already passed through `cleanText(line)`, so the
  trailing whole-block cleanup looked like it might only need whitespace/newline normalization
  rather than another full mojibake-repair pass.
- The targeted visible-text and mojibake regression suite stayed green, but the `.xls` hotspot set
  regressed clearly, so the change was reverted.

Validation:
- `go test -buildvcs=false -run 'TestCleanVisibleText.*|TestMojibakePunctuationIsRepairedAndBinaryNoiseDropped|TestMojibakeContractionsAreRepaired|TestDocx223624SampleDropsCyrillicMojibakeControlLines|TestLegacyPPTSampleDropsMojibakeControlLines|TestLegacyDOCWord95CyrillicDoesNotEmitMojibake|TestLegacyXLSUTF8JapaneseCodePageMojibakeIsRepaired|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestCleanTextKeepsMarkdownTableRowsWithCyrillicText|TestCleanTextKeepsMarkdownListRowsWithCyrillicText|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-cleanvisible-nomojibake-tail-xls6.json -csv testdata\web-samples\reports\perf-exp-cleanvisible-nomojibake-tail-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - `cleanTextNoMojibakeRepair` tail variant: `34909 ms`

Decision:
- Reverted. The final whole-block `cleanText(...)` pass is still paying for itself on the measured
  legacy `.xls` workloads.

## 2026-07-01 rejected: preallocate multiline `cleanVisibleText` output slices

- Tried preallocating `out` and `pendingGlyphSoup` in the multiline branch of `cleanVisibleText(...)`
  using the cleaned line count as an initial capacity hint.
- The idea was a simple allocation reduction: the branch already knows it is processing a multiline
  block, so pre-sizing the output slice looked like an easy way to reduce append growth on BIFF
  visible-text cleanup.
- The targeted visible-text and mojibake regression suite stayed green, but the `.xls` batch slowed
  down materially, so the preallocation was reverted.

Validation:
- `go test -buildvcs=false -run 'TestCleanVisibleText.*|TestMojibakePunctuationIsRepairedAndBinaryNoiseDropped|TestMojibakeContractionsAreRepaired|TestDocx223624SampleDropsCyrillicMojibakeControlLines|TestLegacyPPTSampleDropsMojibakeControlLines|TestLegacyDOCWord95CyrillicDoesNotEmitMojibake|TestLegacyXLSUTF8JapaneseCodePageMojibakeIsRepaired|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestCleanTextKeepsMarkdownTableRowsWithCyrillicText|TestCleanTextKeepsMarkdownListRowsWithCyrillicText|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-cleanvisible-prealloc-xls6.json -csv testdata\web-samples\reports\perf-exp-cleanvisible-prealloc-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - multiline preallocation variant: `35495 ms`

Decision:
- Reverted. The extra counting/allocation work lost clearly on the measured `.xls` set.

## 2026-07-01 rejected: switch `looksLikeEmbeddedPDFText` from `ToLower` to repeated ASCII-fold scans

- Tried rewriting `looksLikeEmbeddedPDFText(...)` so it would trim once and then use
  `containsASCIIFold(...)` for each PDF marker instead of allocating a lowercase copy of the whole
  string with `strings.ToLower(...)`.
- The idea looked attractive because `looksLikeEmbeddedPDFText(...)` still shows up in the BIFF
  markdown and text hot paths, and the lowercase allocation is obvious in the profile.
- In practice the repeated ASCII-fold substring scans were much slower than a single lowercase copy
  plus ordinary `strings.Contains(...)`, so the change was reverted immediately.

Validation:
- `go test -buildvcs=false -run 'TestMarkdownTableCell.*|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestXLSXFormulaTextIsNotVisibleText|TestOOXMLChartFormulaReferencesAreNotVisibleText|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-pdf-asciifold-xls6.json -csv testdata\web-samples\reports\perf-exp-pdf-asciifold-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - ASCII-fold PDF variant: `46790 ms`

Decision:
- Reverted immediately. The repeated ASCII-fold scans were dramatically slower on the measured
  BIFF-heavy workloads.

## 2026-07-01 rejected: cheap marker guard before lowercase PDF-text detection

- Tried adding a tiny early return in `looksLikeEmbeddedPDFText(...)` that skipped the later
  `strings.ToLower(...)` work when the candidate text contained none of the obvious marker bytes
  used by the PDF checks: `%`, `/`, `e`, or `E`.
- The guard looked attractive because most ordinary visible text should fail it immediately, while
  real PDF fragments almost always contain at least one of those characters.
- On the measured `.xls` workloads the extra guard still lost badly, so the change was reverted.

Validation:
- `go test -buildvcs=false -run 'TestMarkdownTableCell.*|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestXLSXFormulaTextIsNotVisibleText|TestOOXMLChartFormulaReferencesAreNotVisibleText|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-pdf-marker-guard-xls6.json -csv testdata\web-samples\reports\perf-exp-pdf-marker-guard-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - PDF marker-guard variant: `37633 ms`

Decision:
- Reverted immediately. The cheap-looking marker test was still a clear net loss on the hot BIFF
  workloads.

## 2026-07-01 rejected: BIFF-only direct write for already-safe single-line prepared markdown cells

- Tried adding a very narrow fast path in `biffSetMarkdownCellAt(...)`: after
  `cleanMarkdownTableCellValue(...)`, `looksLikeEmbeddedPDFText(...)`, and
  `looksLikeSpreadsheetFormulaExpression(...)`, single-line values with no `\` or `|` would be
  written directly into the markdown row map instead of flowing through
  `prepareMarkdownTableCellValue(...)`.
- This was intended as a BIFF-only version of the earlier plain-cell fast-path idea, with a much
  tighter equivalence condition and no effect on `.docx` or `.xlsx` markdown cell preparation.
- The targeted regression suite stayed green, but the `.xls` batch was still slower overall, so the
  change was reverted.

Validation:
- `go test -buildvcs=false -run 'TestMarkdownTableCell.*|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-direct-prepared-singleline-xls6.json -csv testdata\web-samples\reports\perf-exp-biff-direct-prepared-singleline-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - BIFF-only direct prepared-cell write: `32763 ms`

Decision:
- Reverted. Even this tighter BIFF-only plain-cell fast path did not pay for itself on the measured
  legacy `.xls` workloads.

## 2026-07-01 rejected: switch exact-precheck `seen` set from `map[string]bool` to `map[string]struct{}`

- Tried changing the local `seen` set inside `markdownBackfillExactLinesCovered(...)` from
  `map[string]bool` to `map[string]struct{}`.
- The idea was as small as it sounds: the exact-match precheck only needs membership, so using an
  empty-struct set looked like a cheap way to shave a bit of map value overhead in the hottest
  retained backfill loop.
- The targeted regression suite stayed green, but the measured `.xls` batch was slower, so the
  change was reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-exact-seen-struct-xls6.json -csv testdata\web-samples\reports\perf-exp-exact-seen-struct-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - `map[string]struct{}` seen-set variant: `36050 ms`

Decision:
- Reverted. This tiny set representation swap did not help the hot exact-precheck loop.

## 2026-07-01 rejected: skip exact-set table-cell extraction when a markdown line has no pipe

- Tried adding a very small guard in `markdownBackfillExactSet(...)` so it would only call
  `markdownVisibleTableCells(...)` when the current markdown line actually contained a `|`.
- The rationale was simple: lines without a pipe cannot be markdown table rows, so the expensive
  table-cell extraction looked avoidable in the hot exact-set builder.
- The targeted regression suite stayed green, but the `.xls` batch regressed clearly, so the guard
  was reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-exactset-pipe-guard-xls6.json -csv testdata\web-samples\reports\perf-exp-exactset-pipe-guard-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - exact-set pipe guard variant: `33812 ms`

Decision:
- Reverted. Even this obvious-looking table-cell guard was a clear loss on the measured exact-set
  hotspot workloads.

## 2026-07-01 rejected: preallocate `markdownBackfillExactSet` from markdown line count

- Tried giving `markdownBackfillExactSet(...)` a coarse initial map capacity based on
  `strings.Count(markdown, "\n")+1`.
- The change did not alter any matching behavior; it only aimed to reduce rehashing and bucket
  growth while the exact-set builder inserts visible lines, visible markdown lines, and extracted
  table cells.
- On the measured `.xls` workloads this was still a clear loss, so the preallocation was reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-exactset-prealloc-xls6.json -csv testdata\web-samples\reports\perf-exp-exactset-prealloc-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - exact-set preallocation variant: `40192 ms`

Decision:
- Reverted. The coarse exact-set capacity hint was a strong net loss on the hotspot set.

## 2026-07-01 rejected: reuse exact-set output as the source for `markdownBackfillCoverageSet`

- Tried reusing the retained exact-precheck work in `missingMarkdownText(...)` more directly:
  compute `exactSet` once, pass it into the exact-coverage check, and if the exact precheck fails,
  build `coverage` from the already-populated exact strings instead of rescanning the full markdown
  source a second time.
- Conceptually this looked clean because `coverageSet` is mostly the exact strings plus comparable
  variants, so the extra full markdown pass seemed redundant.
- The targeted regression suite stayed green, but the `.xls` batch regressed sharply, so the reuse
  scheme was reverted.

Validation:
- `go test -buildvcs=false -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText' -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-reuse-exactset-for-coverage-xls6.json -csv testdata\web-samples\reports\perf-exp-reuse-exactset-for-coverage-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - reused exact-set coverage variant: `43414 ms`

Decision:
- Reverted. The seemingly redundant second markdown pass still outperformed the exact-set reuse
  approach on the measured hotspot workloads.

## 2026-07-01 rejected: visible-scalar early keep in `cleanText` / `cleanVisibleText`

- Tried adding a helper that treated obvious scalar display lines as always-visible text before the
  binary-control and short-glyph-soup filters. The helper accepted plain numeric/date/time-like
  strings plus BIFF boolean and Excel error display literals, and it was wired into both
  `cleanText(...)` and `cleanVisibleText(...)`.
- The motivation came directly from the current BIFF profile: `biffNumberDisplayValue(...)` drives a
  large share of `forEachBIFFText(...)`, and the hottest single-line cost sits inside
  `looksLikeDiscardableVisibleTextLine(...)` / `looksLikeBinaryControlFragment(...)`. On paper this
  looked like a cheap way to spare obviously visible scalar cell values from that classification
  path.
- A focused regression run stayed green after adjusting the multiline path, but the measured
  six-file `.xls` batch regressed sharply, so the experiment was reverted.

Validation:
- `go test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText" -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-visible-scalar-guard-xls6.json -csv testdata\web-samples\reports\perf-exp-visible-scalar-guard-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - visible-scalar early-keep variant: `37551 ms`

Decision:
- Reverted. Even though the heuristic matched the current BIFF hotspot story, its extra branching
  and repeated classification checks were a clear net loss on the measured legacy `.xls` workloads.

## 2026-07-01 rejected: fold `cleanTextFastPath` slash-marker scan into the main byte loop

- Tried simplifying `cleanTextFastPath(...)` by removing the separate
  `strings.ContainsAny("/\\%[]()")` precheck and instead rejecting those bytes inside the existing
  per-byte loop.
- The intended win was straightforward: keep the fast-path semantics the same while avoiding one
  extra whole-string scan plus the `ContainsAny` ASCII-set setup that shows up in the current CPU
  profile.
- The targeted regression suite stayed green, but the six-file `.xls` batch was still slower, and
  the measured text-length totals on those legacy samples shifted enough that the change was not
  worth keeping. The experiment was reverted immediately.

Validation:
- `go test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText" -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-cleantext-fastpath-inline-scan-xls6.json -csv testdata\web-samples\reports\perf-exp-cleantext-fastpath-inline-scan-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - in-loop slash-marker scan variant: `30489 ms`

Decision:
- Reverted. The extra up-front `ContainsAny` scan was cheaper in practice than pushing more checks
  into the hot byte loop.

## 2026-07-01 retained: single-line fast path in `prepareMarkdownTableCellValue`

- Added a narrow fast path at the top of `prepareMarkdownTableCellValue(...)` for single-line cell
  values that do not contain `\r` or `\n` and do not look like raw RTF payloads.
- The retained path keeps the existing trimming behavior through `normalizeMarkdownTextLine(...)`
  and still escapes backslashes and pipes exactly as before, but it skips the heavier
  `markdownText(...) -> strings.Split(...) -> compact duplicate lines -> strings.Join(...)` flow
  when that machinery cannot change the result for an ordinary single-line cell value.
- Raw single-line RTF snippets still fall back to the old path, so visible `\line` output and other
  RTF normalization behavior stay intact.

Validation:
- `go test -buildvcs=false -run "TestPrepareMarkdownTableCellValueSingleLineFastPath|TestMarkdownTableCell.*|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|Test.*Markdown|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable" -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-prepare-cell-singleline-fastpath-xls6.json -csv testdata\web-samples\reports\perf-exp-prepare-cell-singleline-fastpath-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-prepare-cell-singleline-fastpath-keyset.json -csv testdata\web-samples\reports\perf-exp-prepare-cell-singleline-fastpath-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`
- `go test -buildvcs=false ./...`

Observed results:
- 6-file `.xls` batch vs retained list-marker guard build:
  - retained list-marker guard total: `29367 ms`
  - single-line `prepareMarkdownTableCellValue` fast path: `26588 ms`
- 30-file keyset/top30 rerun:
  - current stable rerun total: `88002 ms`
  - single-line `prepareMarkdownTableCellValue` fast path: `84927 ms`
- Full suite:
  - `ok officeread 356.276s`

Decision:
- Retained. This was a real same-machine win on the hot legacy spreadsheet workloads without
  disturbing targeted regressions or the full test suite.

## 2026-07-01 rejected: BIFF raw single-line formula/PDF prefilter before cell cleaning

- Tried adding a narrow BIFF-only prefilter at the top of `biffSetMarkdownCellAt(...)`: for raw
  single-line cell values, trim once and skip the full `cleanMarkdownTableCellValue(...)` path when
  the original value is already empty, embedded-PDF-like, or a spreadsheet formula/range
  expression.
- The idea was to save the hot cell-cleaning cost for values that the existing post-clean check was
  already going to throw away anyway.
- The targeted regression suite stayed green, and the six-file `.xls` hotspot batch even looked
  good, but the broader keyset regressed enough that the prefilter was not a trustworthy net win.
  The experiment was reverted.

Validation:
- `go test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText" -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-raw-formula-pdf-prefilter-xls6.json -csv testdata\web-samples\reports\perf-exp-biff-raw-formula-pdf-prefilter-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-raw-formula-pdf-prefilter-keyset.json -csv testdata\web-samples\reports\perf-exp-biff-raw-formula-pdf-prefilter-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained single-line cell-prepare fast path build:
  - retained single-line cell-prepare fast path total: `26588 ms`
  - raw BIFF formula/PDF prefilter variant: `26269 ms`
- 30-file keyset/top30 rerun:
  - retained single-line cell-prepare fast path total: `84927 ms`
  - raw BIFF formula/PDF prefilter variant: `94597 ms`

Decision:
- Reverted. The prefilter helped the smallest hotspot batch but lost badly enough on the broader
  mixed workload that it was not worth keeping.

## 2026-07-01 retained: direct visible-return fast path for simple single-line cleaned table cells

- Added a narrow fast path inside `cleanMarkdownTableCellValue(...)` for single-line values after
  `cleanText(...)` and `stripInlineHiddenOfficeReferences(...)`.
- When the trimmed line is non-empty, does not look like discardable hidden Office metadata, and
  also does not match the cheap control-fragment prefilter, the function now returns the cleaned
  line directly instead of falling through the full `markdownTableCellValueLineDiscarded(...)`
  chain.
- The fast path still preserves the original byte-limit behavior by applying the existing
  `truncateUTF8Bytes(...)` plus trailing `...` logic before returning, which was important for the
  large-cell truncation tests.
- Conceptually this keeps the expensive binary-control classifier focused on candidates that already
  look suspicious, instead of making every ordinary single-line cell pay that cost.

Validation:
- `go test -buildvcs=false -run "TestPrepareMarkdownTableCellValueSingleLineFastPath|TestMarkdownTableCell.*|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|Test.*Markdown|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestDOCXMarkdownTruncatesLargeTableCells|TestXLSXMarkdownTruncatesLargeCells" -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-clean-cell-singleline-direct-visible-return-xls6.json -csv testdata\web-samples\reports\perf-exp-clean-cell-singleline-direct-visible-return-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-clean-cell-singleline-direct-visible-return-keyset.json -csv testdata\web-samples\reports\perf-exp-clean-cell-singleline-direct-visible-return-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`
- `go test -buildvcs=false ./...`

Observed results:
- 6-file `.xls` batch vs retained single-line cell-prepare fast path build:
  - retained single-line cell-prepare fast path total: `26588 ms`
  - direct visible-return cell-clean fast path: `24501 ms`
- 30-file keyset/top30 rerun:
  - retained single-line cell-prepare fast path total: `84927 ms`
  - direct visible-return cell-clean fast path: `78589 ms`
- Full suite:
  - `ok officeread 352.266s`

Decision:
- Retained. This meaningfully reduced the hot single-line BIFF cell-clean path, held up on the
  broader keyset, and passed the full suite once truncation behavior was preserved.

## 2026-07-01 rejected: backslash guard in unescapeMarkdownInlineFormattingMarkers

- Tried adding a tiny early return in unescapeMarkdownInlineFormattingMarkers(...): if the string
  did not contain a backslash at all, return it unchanged instead of paying for four unconditional
  strings.ReplaceAll(...) passes.
- This helper sits directly inside the hot markdown table-cell extraction path, both for whole cell
  values and <br>-split segments, so the idea looked promising after the surrounding BIFF
  cell-clean path had already been trimmed down.
- The focused regression suite stayed green and the six-file .xls hotspot batch improved
  dramatically, but the broader keyset regressed enough that the guard was not a trustworthy net
  win. The change was reverted.

Validation:
- go test -buildvcs=false -run "TestPrepareMarkdownTableCellValueSingleLineFastPath|TestMarkdownTableCell.*|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|Test.*Markdown|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestDOCXMarkdownTruncatesLargeTableCells|TestXLSXMarkdownTruncatesLargeCells" -count=1 .
- go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-unescape-inline-formatting-backslash-guard-xls6.json -csv testdata\web-samples\reports\perf-exp-unescape-inline-formatting-backslash-guard-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls
- go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-unescape-inline-formatting-backslash-guard-keyset.json -csv testdata\web-samples\reports\perf-exp-unescape-inline-formatting-backslash-guard-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx

Observed results:
- 6-file .xls batch vs retained direct visible-return cell-clean fast path build:
  - retained direct visible-return cell-clean fast path total: 24501 ms
  - backslash guard in unescapeMarkdownInlineFormattingMarkers: 20274 ms
- 30-file keyset/top30 rerun:
  - retained direct visible-return cell-clean fast path total: 78589 ms
  - backslash guard in unescapeMarkdownInlineFormattingMarkers: 92811 ms

Decision:
- Reverted. The hotspot-batch gain did not survive the broader mixed-workload rerun.
## 2026-07-01 rejected: marker-presence guard in stripMarkdownInlineWrappers

- Tried adding a cheap early return to stripMarkdownInlineWrappers(...): if the input contained
  none of *, _, ~, or `  `, return it unchanged instead of entering the wrapper-stripping
  loop.
- The idea was to mirror the existing guard already used in stripMarkdownInlineFormatting(...)
  and cut down pointless per-cell wrapper checks inside markdownVisibleTableCells(...).
- The focused regression suite stayed green, but the six-file .xls hotspot batch regressed
  clearly, so the guard was reverted immediately without promoting it to the broader keyset.

Validation:
- go test -buildvcs=false -run "TestPrepareMarkdownTableCellValueSingleLineFastPath|TestMarkdownTableCell.*|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|Test.*Markdown|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestDOCXMarkdownTruncatesLargeTableCells|TestXLSXMarkdownTruncatesLargeCells" -count=1 .
- go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-inline-wrapper-marker-guard-xls6.json -csv testdata\web-samples\reports\perf-exp-inline-wrapper-marker-guard-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls

Observed results:
- 6-file .xls batch vs retained direct visible-return cell-clean fast path build:
  - retained direct visible-return cell-clean fast path total: 24501 ms
  - marker-presence guard in stripMarkdownInlineWrappers: 29867 ms

Decision:
- Reverted. Even this seemingly harmless marker check was a clear loss on the measured hotspot
  spreadsheets.
## 2026-07-01 rejected: backslash guard in unescapeMarkdownVisibleText

- Tried adding a narrow early return to unescapeMarkdownVisibleText(...): if the string did not
  contain a backslash at all, return it unchanged instead of paying for the whole sequence of
  markdown escape ReplaceAll(...) calls.
- The current profile showed this helper at about 330 ms cumulative inside
  markdownVisibleTableCells(...), so it looked like a plausible sibling to the earlier rejected
  unescapeMarkdownInlineFormattingMarkers(...) guard.
- The focused regression suite stayed green, but the six-file .xls hotspot batch regressed
  immediately, so the change was reverted without promoting it to a broader rerun.

Validation:
- go test -buildvcs=false -run "TestPrepareMarkdownTableCellValueSingleLineFastPath|TestMarkdownTableCell.*|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|Test.*Markdown|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestDOCXMarkdownTruncatesLargeTableCells|TestXLSXMarkdownTruncatesLargeCells" -count=1 .
- go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-unescape-visible-backslash-guard-xls6.json -csv testdata\web-samples\reports\perf-exp-unescape-visible-backslash-guard-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls

Observed results:
- 6-file .xls batch vs retained direct visible-return cell-clean fast path build:
  - retained direct visible-return cell-clean fast path total: 24501 ms
  - backslash guard in unescapeMarkdownVisibleText: 33405 ms

Decision:
- Reverted. The extra guard cost more than the skipped replacements on the measured hotspot
  spreadsheets.
## 2026-07-01 rejected: no-escape fast path in splitMarkdownTableRow

- Tried adding a fast path at the top of splitMarkdownTableRow(...): if the row contained no
  backslash at all, return strings.Split(row, "|") directly instead of using the rune loop and
  builder-based escaped-pipe parser.
- The motivation was simple: most markdown table rows in the hot BIFF workloads do not appear to use
  escaped pipes, so the general escaped splitter looked avoidable for the common case.
- The focused regression suite stayed green, and the six-file .xls hotspot batch improved a bit,
  but the broader keyset regressed overall, so the fast path was reverted.

Validation:
- go test -buildvcs=false -run "TestPrepareMarkdownTableCellValueSingleLineFastPath|TestMarkdownTableCell.*|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|Test.*Markdown|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestDOCXMarkdownTruncatesLargeTableCells|TestXLSXMarkdownTruncatesLargeCells" -count=1 .
- go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-split-markdown-table-row-noescape-fastpath-xls6.json -csv testdata\web-samples\reports\perf-exp-split-markdown-table-row-noescape-fastpath-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls
- go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-split-markdown-table-row-noescape-fastpath-keyset.json -csv testdata\web-samples\reports\perf-exp-split-markdown-table-row-noescape-fastpath-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx

Observed results:
- 6-file .xls batch vs retained direct visible-return cell-clean fast path build:
  - retained direct visible-return cell-clean fast path total: 24501 ms
  - no-escape splitMarkdownTableRow fast path: 24100 ms
- 30-file keyset/top30 rerun:
  - retained direct visible-return cell-clean fast path total: 78589 ms
  - no-escape splitMarkdownTableRow fast path: 80138 ms

Decision:
- Reverted. The common-case shortcut helped the smallest hotspot batch but still lost on the
  broader mixed workload.
## 2026-07-01 rejected: only run escaped table-cell exact variant for pipe-containing lines

- Tried narrowing the final branch inside markdownBackfillExactSetContainsLine(...) so it would
  only compute markdownBackfillVisibleText(escapeMarkdownTableCell(line)) when the candidate line
  actually contained a |.
- The rationale was that this branch looked most relevant for table-cell-like text, while
  candidateLine(...) has already split the source text into single lines and cleaned away hidden
  references before the exact-set check runs.
- The targeted regression suite stayed green, but the six-file .xls hotspot batch still regressed,
  so the guard was reverted immediately.

Validation:
- go test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText" -count=1 .
- go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-exactsetcontains-pipe-only-escapedvariant-xls6.json -csv testdata\web-samples\reports\perf-exp-exactsetcontains-pipe-only-escapedvariant-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls

Observed results:
- 6-file .xls batch vs retained direct visible-return cell-clean fast path build:
  - retained direct visible-return cell-clean fast path total: 24501 ms
  - pipe-only escaped exact-variant guard: 25241 ms

Decision:
- Reverted. Even this very narrow exact-set variant guard was a small but clear loss on the
  measured hotspot spreadsheets.
## 2026-07-01 retained: plain visible-line fast path for markdown exact-backfill normalization

- Added a conservative fast path at the top of `markdownVisibleLineText(...)`: after trimming
  whitespace, return the line directly when it has no markdown or HTML structure markers, no
  ordered/list-style prefix, and no trailing hard-break backslash.
- The goal was to cut repeated normalization work inside the current backfill hotspot chain
  (`markdownBackfillExactSet(...)`, `markdownBackfillExactSetContainsLine(...)`, and nearby
  visible-text checks) without changing the more complex markdown parsing behavior for headings,
  lists, links, HTML, escaped markup, or table-like lines.
- Added `TestMarkdownVisibleLineTextFastPath` to pin the intended boundaries for plain text,
  ordered-list normalization, and hard line-break stripping.

Validation:
- `go test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText" -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-markdown-visible-line-fastpath-xls6.json -csv testdata\web-samples\reports\perf-exp-markdown-visible-line-fastpath-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-markdown-visible-line-fastpath-keyset.json -csv testdata\web-samples\reports\perf-exp-markdown-visible-line-fastpath-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`
- `go test -buildvcs=false ./...`

Observed results:
- 6-file .xls batch vs retained direct visible-return cell-clean fast path build:
  - retained direct visible-return cell-clean fast path total: `24501 ms`
  - plain visible-line fast path: `21653 ms`
- 30-file keyset/top30 rerun:
  - retained direct visible-return cell-clean fast path total: `78589 ms`
  - plain visible-line fast path: `65575 ms`
- Full suite:
  - `ok officeread 278.734s`

Decision:
- Retained. Unlike the recent rejected guard-style experiments, this broader normalization fast path
  improved both the hotspot `.xls` batch and the mixed keyset substantially while keeping the full
  suite green.
## 2026-07-01 rejected: bare inline-link/backfill-visible fast paths

- Tried two related shortcuts around markdown exact-backfill normalization:
  `markdownInlineLinkVisibleText(...)` returned immediately when a line had no `[` at all, and
  `markdownBackfillVisibleText(...)` skipped reference-definition work for lines without `[` while
  still handling bare backslash unescaping.
- The idea was to shave repeated single-line normalization work inside
  `markdownBackfillExactSetContainsLine(...)` and the later missing-line checks.
- The focused regression suite stayed green, but the six-file `.xls` hotspot batch regressed, so
  both shortcuts were reverted immediately without a broader rerun.

Validation:
- `go test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText" -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-backfill-visible-inline-fastpaths-xls6.json -csv testdata\web-samples\reports\perf-exp-backfill-visible-inline-fastpaths-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained plain visible-line fast path build:
  - retained plain visible-line fast path total: `21653 ms`
  - bare inline-link/backfill-visible fast paths: `24250 ms`

Decision:
- Reverted. The extra shortcutting still lost on the real hotspot spreadsheets.

## 2026-07-01 rejected: map preallocation in backfill exact-coverage path

- Tried preallocating the string maps used in `missingMarkdownText(...)`,
  `markdownBackfillExactLinesCovered(...)`, and `markdownBackfillExactSet(...)`, plus reusing one
  `strings.Split(markdown, "\n")` result when building the exact set.
- This was motivated by the refreshed profile showing heavy time in map hashing, lookup, and split
  growth while exact-backfill normalization remained the dominant path.
- The focused regression suite stayed green, but the six-file `.xls` hotspot batch still regressed
  relative to the retained baseline.

Validation:
- `go test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText" -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-backfill-map-prealloc-xls6.json -csv testdata\web-samples\reports\perf-exp-backfill-map-prealloc-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained plain visible-line fast path build:
  - retained plain visible-line fast path total: `21653 ms`
  - backfill map preallocation experiment: `23829 ms`

Decision:
- Reverted. Fewer obvious map growth points did not translate into a net win on the measured
  hotspot batch.

## 2026-07-01 rejected: markdown table `<br>` segment guard

- Tried skipping the secondary `<br>` segment-cleaning loop inside `markdownVisibleTableCells(...)`
  unless the already-cleaned cell text actually contained `<br>`.
- The intent was to avoid a second round of trimming, markdown wrapper stripping, and formatting
  unescaping for ordinary single-segment cells.
- The focused regression suite stayed green, but the six-file `.xls` hotspot batch regressed
  dramatically, so the change was reverted immediately.

Validation:
- `go test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText" -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-markdown-table-br-segment-guard-xls6.json -csv testdata\web-samples\reports\perf-exp-markdown-table-br-segment-guard-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained plain visible-line fast path build:
  - retained plain visible-line fast path total: `21653 ms`
  - `<br>` segment guard: `31509 ms`

Decision:
- Reverted. The avoided second-stage cleanup was not worth the extra branch on the measured
  spreadsheets.

## 2026-07-01 rejected: lightweight pipe-count check in markdownLikelyTableRow

- Tried replacing `len(splitMarkdownTableRow(trimmed)) >= 3` with a custom unescaped-pipe counter in
  `markdownLikelyTableRow(...)`, so non-table lines with pipes could avoid constructing the full
  split slice just to answer the yes/no precheck.
- The focused regression suite stayed green, but the six-file `.xls` hotspot batch still regressed,
  so the precheck was reverted.

Validation:
- `go test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText" -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-markdown-likely-table-row-pipe-count-xls6.json -csv testdata\web-samples\reports\perf-exp-markdown-likely-table-row-pipe-count-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained plain visible-line fast path build:
  - retained plain visible-line fast path total: `21653 ms`
  - lightweight pipe-count precheck: `24043 ms`

Decision:
- Reverted. The custom precheck was still slower than the existing splitter-based test on the
  hotspot batch.

## 2026-07-01 rejected: alternate candidate-line cache shapes

- Evaluated two alternate shapes for the `missingMarkdownText(...)` candidate-line cache:
  one eager precompute pass over every input line, and one lazily-filled slice keyed by line index
  instead of the retained string-keyed deduplication map.
- The motivation came from the updated profile, which showed significant time in the candidate-line
  cache closure itself. The experiments aimed to remove string hashing and repeated map traffic.
- Both variants kept the focused regression suite green but regressed badly on the six-file `.xls`
  hotspot batch, so both were reverted.

Validation:
- `go test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText" -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-precompute-backfill-candidate-lines-xls6.json -csv testdata\web-samples\reports\perf-exp-precompute-backfill-candidate-lines-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-index-lazy-backfill-candidate-cache-xls6.json -csv testdata\web-samples\reports\perf-exp-index-lazy-backfill-candidate-cache-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained plain visible-line fast path build:
  - retained plain visible-line fast path total: `21653 ms`
  - eager candidate-line precompute: `32871 ms`
  - index-lazy candidate-line cache: `23914 ms`

Decision:
- Reverted. The retained string-keyed cache still matches the workload better than either eager or
  purely index-based alternatives.

## 2026-07-01 rejected: control-fragment string-scan tightening

- Tried a small batch of string-scan rewrites around the single-line table-cell control-text path:
  `prepareMarkdownTableCellValue(...)` used `hasPrefixFold(...)` instead of lowercasing the whole
  trimmed value, `looksLikeOOXMLMarkupNameFragment(...)` switched to manual trimming plus
  length-based matching, and `oleIdentifierAssignmentValue(...)` stopped lowercasing the key before
  comparing it.
- The motivation came from the refreshed profile of `cleanMarkdownTableCellValue(...)`, where the
  single-line fast path still spent most of its time inside hidden/control-text classifiers.
- The focused regression suite stayed green, but the six-file `.xls` hotspot batch still regressed,
  so the whole bundle was reverted.

Validation:
- `go test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText" -count=1 .`
- `go run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-control-fragment-string-scan-tightening-xls6.json -csv testdata\web-samples\reports\perf-exp-control-fragment-string-scan-tightening-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- 6-file `.xls` batch vs retained plain visible-line fast path build:
  - retained plain visible-line fast path total: `21653 ms`
  - control-fragment string-scan tightening bundle: `23720 ms`

Decision:
- Reverted. Even focused string-scan cleanups in the control-text classifiers were still slower on
  the target workload than the retained baseline.
## 2026-07-05 rejected: PDF/formula post-clean classifier tightening

- Tried a small cleanup in the post-clean table-cell filters used by `biffSetMarkdownCellAt(...)`:
  `looksLikeEmbeddedPDFText(...)` stopped lowercasing the entire string and instead used
  `containsASCIIFold(...)` for each PDF marker, while `looksLikeSpreadsheetFormulaExpression(...)`
  added a cheap `strings.ContainsAny("!(")` precheck before the more structured range/function
  parsing.
- The hope was that this would cheaply reject the vast majority of ordinary visible table cells
  before paying for full formula or embedded-PDF heuristics.
- The focused regression suite stayed green, and the six-file `.xls` hotspot batch was essentially
  unchanged, but the broader 30-file keyset regressed clearly, so the change was reverted.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-pdf-formula-checks-tightening-xls6.json -csv testdata\web-samples\reports\perf-exp-pdf-formula-checks-tightening-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-pdf-formula-checks-tightening-keyset.json -csv testdata\web-samples\reports\perf-exp-pdf-formula-checks-tightening-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained plain visible-line fast path build:
  - retained plain visible-line fast path total: `21653 ms`
  - PDF/formula classifier tightening: `21641 ms`
- 30-file keyset/top30 rerun:
  - retained plain visible-line fast path total: `65575 ms`
  - PDF/formula classifier tightening: `67804 ms`

Decision:
- Reverted. The tiny hotspot-batch noise-level gain did not survive the broader mixed workload.
## 2026-07-05 rejected: BIFF simple display-value direct cell bypass

- Tried routing BIFF-generated simple display values (NUMBER, RK, BOOLERR, and formula display
  values) through a dedicated direct cell-store helper instead of the general
  `cleanMarkdownTableCellValue(...)` plus `prepareMarkdownTableCellValue(...)` pipeline.
- The reasoning was that these values are already generated by our own BIFF display formatting and
  look much more like trusted short scalars than arbitrary string-cell payloads, so they seemed like
  good candidates for a specialized fast path.
- The focused regression suite stayed green, and the six-file `.xls` hotspot batch improved
  dramatically, but the broader BIFF subset inside the 30-file keyset regressed badly enough that
  the optimization was not trustworthy and was reverted.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-simple-display-cell-bypass-xls6.json -csv testdata\web-samples\reports\perf-exp-biff-simple-display-cell-bypass-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-simple-display-cell-bypass-keyset.json -csv testdata\web-samples\reports\perf-exp-biff-simple-display-cell-bypass-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained plain visible-line fast path build:
  - retained plain visible-line fast path total: `21653 ms`
  - BIFF simple display-value direct bypass: `16666 ms`
- 30-file keyset/top30 rerun:
  - retained plain visible-line fast path total: `65575 ms`
  - BIFF simple display-value direct bypass: `81826 ms`
- Keyset extension split from rerun:
  - retained baseline `.xls` subset: `25808 ms`
  - direct-bypass `.xls` subset rerun: `40595 ms`

Decision:
- Reverted. The hotspot batch win was real but too narrow; the broader BIFF workload regressed too
  much to justify keeping the specialized bypass.
## 2026-07-05 retained: BIFF entry bounds prefilter before markdown-cell decoding

- Added `biffCellInMarkdownBounds(...)` / `biffCellRecordInMarkdownBounds(...)` and used them at the
  BIFF worksheet-markdown record entry points so hidden cells, out-of-range columns, and rows past
  `maxMarkdownTableRows` are rejected before shared-string lookup, legacy/unicode string decoding,
  float formatting, BOOLERR display conversion, or formula display-value parsing.
- The goal was to stop record types like `LABELSST`, `LABEL`, `NUMBER`, `RK`, `BOOLERR`, and direct
  formula display values from doing real work when `biffSetMarkdownCellAt(...)` would inevitably drop
  them later for markdown-table bounds or visibility reasons anyway.
- `biffFormulaStringCell(...)` now reuses the same bound/visibility gate so deferred string-formula
  cells follow the same early decision path.
- Added `TestBIFFCellRecordInMarkdownBounds` to pin the hidden-row, hidden-column, row-limit, and
  column-limit behavior.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-entry-bounds-prefilter-xls6.json -csv testdata\web-samples\reports\perf-exp-biff-entry-bounds-prefilter-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-entry-bounds-prefilter-keyset.json -csv testdata\web-samples\reports\perf-exp-biff-entry-bounds-prefilter-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false ./...`

Observed results:
- 6-file `.xls` batch vs retained plain visible-line fast path build:
  - retained plain visible-line fast path total: `21653 ms`
  - BIFF entry bounds prefilter: `17753 ms`
- 30-file keyset/top30 rerun:
  - retained plain visible-line fast path total: `65575 ms`
  - BIFF entry bounds prefilter: `58544 ms`
- Keyset `.xls` subset:
  - retained plain visible-line fast path `.xls` total: `25808 ms`
  - BIFF entry bounds prefilter `.xls` total: `22440 ms`

Decision:
- Retained. This is the first BIFF-entry optimization in this stretch that improves both the hotspot
  `.xls` batch and the broader mixed keyset by a meaningful margin while preserving the existing
  visibility and bounds behavior.

## 2026-07-05 rejected: coordinate reuse on top of BIFF entry bounds prefilter

- Tried pushing the retained BIFF entry-bounds prefilter one step further by reusing the row/column
  coordinates already parsed during the entry checks, instead of reparsing them again inside the
  downstream markdown-cell setters and formula helpers.
- Added a temporary record-coordinate helper and threaded the reused coordinates through the
  `LABELSST`, `LABEL`, `NUMBER`, `FORMULA`, `BOOLERR`, and `RK` markdown-table paths.
- The six-file `.xls` hotspot batch looked excellent, but the broader mixed keyset and the rerun of
  the `.xls` subset both regressed enough to make the change untrustworthy, so it was reverted.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-entry-coord-reuse-xls6.json -csv testdata\web-samples\reports\perf-exp-biff-entry-coord-reuse-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-entry-coord-reuse-keyset.json -csv testdata\web-samples\reports\perf-exp-biff-entry-coord-reuse-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained BIFF entry bounds prefilter build:
  - retained BIFF entry bounds prefilter total: `17753 ms`
  - coordinate reuse on top of prefilter: `15909 ms`
- 30-file keyset/top30 rerun:
  - retained BIFF entry bounds prefilter total: `58544 ms`
  - coordinate reuse on top of prefilter: `66785 ms`
- Keyset `.xls` subset rerun:
  - retained BIFF entry bounds prefilter `.xls` total: `22440 ms`
  - coordinate reuse on top of prefilter `.xls` total: `27316 ms`

Decision:
- Reverted. The row/column reuse win was too narrow and did not survive the broader BIFF-heavy
  workload.

## 2026-07-05 retained: markdown backfill plain visible-line fast path

- Added a plain-line fast path in `markdownBackfillExactSet(...)` and
  `addMarkdownBackfillCoverage(...)`: once a trimmed markdown line already qualifies as a plain
  visible line, we now record it directly and skip the redundant inline-link, visible-line, and
  table-cell derivations that would have produced the same answer.
- This targets the current post-BIFF hotspot in `missingMarkdownText(...)`, especially the
  `markdownBackfillExactLinesCovered(...)` branch that was rebuilding a large exact-coverage set for
  mostly ordinary text lines.
- The change is intentionally conservative: it only fires when `markdownPlainVisibleLine(...)`
  already proves the line has no markdown syntax that needs extra normalization.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-plainline-exactset-fastpath-xls6.json -csv testdata\web-samples\reports\perf-exp-plainline-exactset-fastpath-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-plainline-exactset-fastpath-keyset.json -csv testdata\web-samples\reports\perf-exp-plainline-exactset-fastpath-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false ./...`

Observed results:
- 6-file `.xls` batch vs retained BIFF entry bounds prefilter build:
  - retained BIFF entry bounds prefilter total: `17753 ms`
  - markdown backfill plain visible-line fast path: `15947 ms`
- 30-file keyset/top30 rerun:
  - retained BIFF entry bounds prefilter total: `58544 ms`
  - markdown backfill plain visible-line fast path: `48484 ms`
- Keyset `.xls` subset:
  - retained BIFF entry bounds prefilter `.xls` total: `22440 ms`
  - markdown backfill plain visible-line fast path `.xls` total: `18047 ms`

Decision:
- Retained. Unlike the narrower BIFF-only experiments, this one improves both the hotspot `.xls`
  workload and the broader mixed keyset by a large margin while preserving the existing markdown
  backfill semantics.

## 2026-07-05 rejected: BIFF multiline per-line reclean removal

- Tried removing the second `cleanVisibleText(...)` pass inside the multiline branch of
  `forEachBIFFText(...)`, on the theory that the initial whole-string `cleanVisibleText(s)` call had
  already cleaned and filtered each line before the subsequent split.
- Added a temporary multiline BIFF filtering regression test while evaluating the change, then
  reverted both the code and the temporary test once the performance numbers came back worse.
- The simplification looked plausible from the profile, but the measured totals regressed on both
  the hotspot `.xls` batch and the broader mixed keyset, so it was not retained.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds|TestBIFFMultilineTextFiltering" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-multiline-reclean-removal-xls6.json -csv testdata\web-samples\reports\perf-exp-biff-multiline-reclean-removal-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-multiline-reclean-removal-keyset.json -csv testdata\web-samples\reports\perf-exp-biff-multiline-reclean-removal-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - BIFF multiline per-line reclean removal: `16909 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - BIFF multiline per-line reclean removal: `50634 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - BIFF multiline per-line reclean removal `.xls` total: `18724 ms`

Decision:
- Reverted. Removing the per-line reclean made the hottest helper look simpler, but it was slower on
  every retained comparison set we care about.

## 2026-07-05 rejected: markdown exact-set plain-line early return

- After refreshing the CPU profile on top of the retained markdown backfill plain-line fast path, the
  main hotspot was still `missingMarkdownText(...)`, specifically
  `markdownBackfillExactLinesCovered(...)` and `markdownBackfillExactSet(...)`.
- Tried adding a plain-line early return to `markdownBackfillExactSetContainsLine(...)` so ordinary
  candidate lines that missed the exact set would skip the later visible-text, markdown-visible, and
  escaped-table-cell derivations.
- The change was behavior-preserving in the focused regression suite, but it turned out to be a
  performance trap on the actual workloads: the extra `markdownPlainVisibleLine(...)` check cost more
  than the avoided work in aggregate, so the experiment was reverted.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-exactsetcontains-plainline-early-return-xls6.json -csv testdata\web-samples\reports\perf-exp-exactsetcontains-plainline-early-return-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-exactsetcontains-plainline-early-return-keyset.json -csv testdata\web-samples\reports\perf-exp-exactsetcontains-plainline-early-return-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - markdown exact-set plain-line early return: `20935 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - markdown exact-set plain-line early return: `59446 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - markdown exact-set plain-line early return `.xls` total: `24782 ms`

Decision:
- Reverted. This one lost badly enough on every comparison set that it is not worth revisiting in the
  same form.

## 2026-07-05 rejected: markdown backfill map preallocation

- Using the refreshed profile, tried a low-risk allocation reduction pass on the same markdown
  backfill hotspot by preallocating several large transient maps:
  `candidateLineCache`, the `seen` set in `markdownBackfillExactLinesCovered(...)`, the exact set in
  `markdownBackfillExactSet(...)`, and the coverage set in `markdownBackfillCoverageSet(...)`.
- The idea was to shave off some of the visible map growth and rehash cost that showed up in the new
  profile without changing any of the text-normalization behavior.
- The focused regression suite stayed green and the results were much closer than the previous
  experiment, but the retained baseline still won on every important comparison set, so this change
  was also reverted.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-backfill-map-prealloc-xls6.json -csv testdata\web-samples\reports\perf-exp-backfill-map-prealloc-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-backfill-map-prealloc-keyset.json -csv testdata\web-samples\reports\perf-exp-backfill-map-prealloc-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - markdown backfill map preallocation: `16387 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - markdown backfill map preallocation: `50453 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - markdown backfill map preallocation `.xls` total: `18914 ms`

Decision:
- Reverted. The allocation reduction helped a few individual files but still lost the overall workload
  comparison, so it does not earn its place in the retained baseline.

## 2026-07-05 rejected: markdown backfill line-forms cache

- Tried a more structural reuse pass inside `missingMarkdownText(...)` by caching each candidate
  line's derived markdown forms and reusing them across both
  `markdownBackfillExactLinesCovered(...)` and the later missing-line decision loop.
- The target was real duplicated work: for the same normalized candidate line we were recomputing
  the visible-text form, markdown-visible form, and escaped-table-cell visible form in two separate
  phases.
- In practice the extra cache map and struct churn cost much more than the duplicated derivations on
  this workload, so the experiment regressed dramatically and was reverted.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-backfill-lineforms-cache-xls6.json -csv testdata\web-samples\reports\perf-exp-backfill-lineforms-cache-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-backfill-lineforms-cache-keyset.json -csv testdata\web-samples\reports\perf-exp-backfill-lineforms-cache-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - markdown backfill line-forms cache: `26810 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - markdown backfill line-forms cache: `63901 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - markdown backfill line-forms cache `.xls` total: `33281 ms`

Decision:
- Reverted. This confirmed that the remaining hotspot is not cheaply solved by another layer of
  per-line caching in the current shape of the algorithm.

## 2026-07-05 rejected: markdown table-cell `<br>` segment gate

- With the refreshed profile showing `markdownVisibleTableCells(...)` as the dominant sub-cost inside
  `markdownBackfillExactSet(...)`, tried a small range-reduction in that helper: only enter the
  segment-splitting and second cleaning pass when the already-normalized cell text actually contains
  the literal `<br>` marker.
- The motivation was straightforward: for ordinary cells without `<br>`, the old loop still split the
  string and reran most of the markdown cleanup chain even though it would inevitably produce just the
  original cell text again.
- The focused regression suite stayed green, and the broader mixed keyset moved slightly in the right
  direction, but both the six-file `.xls` hotspot batch and the keyset `.xls` subset still regressed
  relative to the retained baseline, so the change was reverted.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-tablecell-br-segment-gate-xls6.json -csv testdata\web-samples\reports\perf-exp-tablecell-br-segment-gate-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-tablecell-br-segment-gate-keyset.json -csv testdata\web-samples\reports\perf-exp-tablecell-br-segment-gate-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - markdown table-cell `<br>` segment gate: `16684 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - markdown table-cell `<br>` segment gate: `48657 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - markdown table-cell `<br>` segment gate `.xls` total: `18653 ms`

Decision:
- Reverted. It shaved some mixed-workload overhead but still lost where the retained baseline is
  already strongest, so it did not clear the bar.

## 2026-07-05 rejected: markdown table separator-row gate

- Built on the same `markdownVisibleTableCells(...)` hotspot by trying to short-circuit whole table
  separator rows such as `| --- | --- |` before the helper entered the expensive per-cell markdown
  cleanup pipeline.
- The hope was that separator rows would be common enough across markdown tables to justify a
  row-level early exit, especially once the profile showed the exact-set builder spending heavily
  inside table-cell extraction.
- The change preserved the focused regression suite, but on the retained workload mix it regressed
  the overall keyset enough that it could not be trusted, so it was reverted along with the paired
  `<br>` gate branch.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-table-separator-row-gate-xls6.json -csv testdata\web-samples\reports\perf-exp-table-separator-row-gate-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-table-separator-row-gate-keyset.json -csv testdata\web-samples\reports\perf-exp-table-separator-row-gate-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - markdown table separator-row gate: `20160 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - markdown table separator-row gate: `51474 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - markdown table separator-row gate `.xls` total: `17889 ms`

Decision:
- Reverted. The slight `.xls`-subset improvement did not survive the broader mixed workload, and the
  total keyset regression was too large to justify keeping the extra branch.

## 2026-07-05 rejected: markdown table-cell plain-text fast path

- Still working from the refreshed exact-set profile, tried a denser fast path inside
  `markdownVisibleTableCells(...)`: after the initial unescape-and-trim step, if a cell already
  qualified as a plain visible markdown line, emit it directly and skip the later footnote, wrapper,
  formatting, autolink, HTML, and `<br>` segment cleanup stages.
- The reasoning was that many ordinary table cells are already plain text after unescaping, so the
  helper was spending a lot of time proving something we already effectively knew.
- The focused regression suite stayed green, but the real workloads regressed sharply across the
  board, so the fast path was reverted. In this case, the extra `markdownPlainVisibleLine(...)`
  classification cost more than the work it saved.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-tablecell-plain-fastpath-xls6.json -csv testdata\web-samples\reports\perf-exp-tablecell-plain-fastpath-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-tablecell-plain-fastpath-keyset.json -csv testdata\web-samples\reports\perf-exp-tablecell-plain-fastpath-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - markdown table-cell plain-text fast path: `20820 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - markdown table-cell plain-text fast path: `58816 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - markdown table-cell plain-text fast path `.xls` total: `23260 ms`

Decision:
- Reverted. This is another example where an intuitively appealing plain-text branch makes the real
  mixed workload slower rather than faster.

## 2026-07-05 rejected: simple markdown table-row fast path

- Tried moving the `markdownVisibleTableCells(...)` optimization up one level from individual cells to
  whole rows: if the trimmed raw table row contained none of the markdown syntax markers that would
  require deeper cleanup (`[]*_~\`<>&\\`), split the row once, trim each cell, drop separator cells,
  and return those values directly.
- This was meant to target the most common spreadsheet-style markdown rows such as
  `| Alice | 93.5 |`, where the full cell cleanup pipeline is mostly proving the obvious.
- The focused regression suite stayed green, but the retained workloads still regressed materially,
  especially on the mixed keyset, so the fast path was reverted.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-simple-table-row-fastpath-xls6.json -csv testdata\web-samples\reports\perf-exp-simple-table-row-fastpath-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-simple-table-row-fastpath-keyset.json -csv testdata\web-samples\reports\perf-exp-simple-table-row-fastpath-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - simple markdown table-row fast path: `17444 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - simple markdown table-row fast path: `58793 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - simple markdown table-row fast path `.xls` total: `20234 ms`

Decision:
- Reverted. Even at the row level, extra syntax classification still cost more than the saved cleanup
  work on the real benchmark mix.

## 2026-07-05 rejected: visible-markdown contains source early exit

- Tried the first truly entry-level early return for `missingMarkdownText(...)`: after normalizing the
  source text with `markdownBackfillSourceText(...)`, build the markdown's full visible-text form once
  with `markdownBackfillVisibleText(markdown)` and return early when that visible markdown string
  already contains the normalized source text as a contiguous substring.
- The idea was to bypass `markdownBackfillExactLinesCovered(...)` and the expensive exact-set
  construction entirely in cases where the structured markdown already carried the source text in a
  straightforward whole-document form.
- The focused regression suite stayed green, but the performance cost of generating the whole visible
  markdown upfront was far larger than the actual short-circuit hit rate on the retained workloads, so
  the change regressed badly and was reverted.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-visible-markdown-contains-source-xls6.json -csv testdata\web-samples\reports\perf-exp-visible-markdown-contains-source-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-visible-markdown-contains-source-keyset.json -csv testdata\web-samples\reports\perf-exp-visible-markdown-contains-source-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - visible-markdown contains source early exit: `25477 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - visible-markdown contains source early exit: `70288 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - visible-markdown contains source early exit `.xls` total: `27836 ms`

Decision:
- Reverted. The early-return hit rate was nowhere near high enough to pay for the upfront whole-markdown
  visible-text conversion.

## 2026-07-05 rejected: BIFF generated-scalar text fast path

- Switched focus back to `biffText(...)` and tried a very narrow fast path that only touched BIFF
  scalar display values generated by our own code: `NUMBER`, `RK`, `BOOLERR`, `MulRK`, and direct
  formula display values.
- The experiment bypassed `forEachBIFFText(...)` / `cleanVisibleText(...)` for those generated
  values and appended them straight into the BIFF text-part stream, on the theory that trusted
  scalar display strings should not need the same filtering pipeline as arbitrary workbook strings.
- The performance numbers initially looked excellent, including a slight mixed-keyset win, but the
  extracted `textBytes` changed for every `.xls` sample in the retained workload sets. That means
  this was not a same-output optimization; it altered compatibility-visible output and had to be
  reverted despite the tempting timings.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-generated-scalar-text-fastpath-xls6.json -csv testdata\web-samples\reports\perf-exp-biff-generated-scalar-text-fastpath-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-generated-scalar-text-fastpath-keyset.json -csv testdata\web-samples\reports\perf-exp-biff-generated-scalar-text-fastpath-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - BIFF generated-scalar text fast path: `12208 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - BIFF generated-scalar text fast path: `48463 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - BIFF generated-scalar text fast path `.xls` total: `14284 ms`
- But retained-vs-experiment `textBytes` changed for every `.xls` sample checked, for example:
  - `008055.xls`: `17684747 -> 19558754`
  - `016161.xls`: `8179182 -> 9264536`
  - `002505.xls`: `3123432 -> 4264184`

Decision:
- Reverted. The timings were attractive, but this was a compatibility regression disguised as a fast
  path, so it cannot be retained.

## 2026-07-05 rejected: BIFF text-part result cache

- Tried a more conservative BIFF-side optimization after reverting the generated-scalar bypass:
  keep `forEachBIFFText(...)` exactly in place, but cache its output slices by original input string
  within `biffText(...)` so repeated shared strings, repeated booleans/errors, repeated numbers, and
  repeated comment text could reuse the full cleaned text-part result instead of recomputing it.
- This was intentionally designed to preserve semantics: unlike the generated-scalar fast path, it
  still ran the same filtering and cleanup pipeline and only memoized the result.
- The output comparison confirmed that the cache preserved `textBytes` on the retained keyset, but it
  still lost on every important timing comparison, so it was reverted as a pure performance miss.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-textpart-cache-xls6.json -csv testdata\web-samples\reports\perf-exp-biff-textpart-cache-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-textpart-cache-keyset.json -csv testdata\web-samples\reports\perf-exp-biff-textpart-cache-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`
- retained-vs-experiment keyset `textBytes` comparison: `NO_DIFF`

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - BIFF text-part result cache: `17156 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - BIFF text-part result cache: `50306 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - BIFF text-part result cache `.xls` total: `19485 ms`

Decision:
- Reverted. This one was semantically clean, but the memoization overhead still outweighed the reuse
  benefit on the retained workload mix.

## 2026-07-05 rejected: BIFF NUMBER/RK display-value cache

- Tried a narrower BIFF-side cache that kept the existing text-cleaning pipeline intact and only
  memoized the formatted display strings for `NUMBER` and `RK` records inside `biffText(...)`,
  keyed by the raw BIFF numeric payload bits.
- This was intended to preserve compatibility exactly while shaving repeated
  `finiteFloatDisplayValue(...)` work from large legacy `.xls` sheets with many duplicate numeric
  cells.
- The output comparison did stay stable, but the extra map lookups and per-record cache management
  still made the retained workloads slower across every decision-making benchmark set, so the change
  was reverted.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-number-rk-display-cache-xls6.json -csv testdata\web-samples\reports\perf-exp-biff-number-rk-display-cache-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-biff-number-rk-display-cache-keyset.json -csv testdata\web-samples\reports\perf-exp-biff-number-rk-display-cache-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`
- retained-vs-experiment `textBytes` comparison: `NO_DIFF` on both `xls6` and the 30-file keyset

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - BIFF NUMBER/RK display-value cache: `17778 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - BIFF NUMBER/RK display-value cache: `53806 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - BIFF NUMBER/RK display-value cache `.xls` total: `20162 ms`

Decision:
- Reverted. This cache preserved output, but it was a pure performance regression on the retained
  benchmark mix.

## 2026-07-05 rejected: direct callback consumption of markdown visible table cells

- Tried refactoring the `markdownVisibleTableCells(...)` hotspot so `markdownBackfillExactSet(...)`
  and `addMarkdownBackfillCoverage(...)` could consume table-cell outputs through a callback instead
  of first building and then iterating an intermediate `[]string`.
- The intended win was to remove one layer of slice growth and short-lived allocations from the
  exact-backfill path without changing any markdown normalization behavior.
- The focused regression suite stayed green and the retained-vs-experiment `textBytes` comparison
  stayed stable, but the callback shape still made the real workloads materially slower, so it was
  reverted.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-direct-tablecell-yield-xls6.json -csv testdata\web-samples\reports\perf-exp-direct-tablecell-yield-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-direct-tablecell-yield-keyset.json -csv testdata\web-samples\reports\perf-exp-direct-tablecell-yield-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`
- retained-vs-experiment keyset `textBytes` comparison: `NO_DIFF`

Observed results:
- 6-file `.xls` batch vs retained markdown backfill plain visible-line fast path build:
  - retained markdown backfill plain visible-line fast path total: `15947 ms`
  - direct callback consumption of markdown visible table cells: `21257 ms`
- 30-file keyset/top30 rerun:
  - retained markdown backfill plain visible-line fast path total: `48484 ms`
  - direct callback consumption of markdown visible table cells: `63106 ms`
- Keyset `.xls` subset:
  - retained markdown backfill plain visible-line fast path `.xls` total: `18047 ms`
  - direct callback consumption of markdown visible table cells `.xls` total: `24636 ms`

Decision:
- Reverted. Removing the intermediate slice looked tidy, but the callback version was a clear
  slowdown on every retained benchmark set.

## 2026-07-05 retained: byte-scan splitMarkdownTableRow

- Reworked `splitMarkdownTableRow(...)` from a rune-and-builder state machine into a byte-level scan
  that only treats ASCII `\` and `|` as control characters and returns direct string slices between
  unescaped pipes.
- This keeps the same escaped-pipe behavior (`\|` stays inside a cell, stray trailing `\` is
  preserved, non-ASCII text passes through untouched) while removing UTF-8 rune decoding and
  incremental builder writes from one of the hottest subpaths under `markdownVisibleTableCells(...)`.
- The focused regression suite and full repository test suite both stayed green, and the experiment
  preserved keyset `textBytes`.

Validation:
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds" -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-split-markdown-table-row-byte-scan-xls6.json -csv testdata\web-samples\reports\perf-exp-split-markdown-table-row-byte-scan-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-exp-split-markdown-table-row-byte-scan-keyset.json -csv testdata\web-samples\reports\perf-exp-split-markdown-table-row-byte-scan-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`
- `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false ./...`
- retained-vs-experiment keyset `textBytes` comparison: `NO_DIFF`

Observed results:
- Against the historical retained plain-line baseline:
  - 6-file `.xls` batch: `15947 ms -> 15566 ms`
  - 30-file keyset/top30 rerun: `48484 ms -> 50489 ms`
  - keyset `.xls` subset: `18047 ms -> 18325 ms`
- Against the same-day current rerun baseline captured in `followup-report-20260705.md`:
  - 30-file keyset/top30 rerun: `54578 ms -> 50489 ms`
  - keyset `.xls` subset: `20668 ms -> 18325 ms`
- Full suite:
  - `ok officeread 219.548s`

Decision:
- Retained. The historical reference numbers and the same-day rerun baseline diverged noticeably,
  which points to meaningful environment drift, so this change was judged on the stronger same-day
  A/B evidence plus full-suite validation. On that basis it produced a real win while preserving
  output.

## 2026-07-05 retained: incremental UTF-16 evidence accounting for legacy `.doc` fallback

- A fresh 1000-file `.doc` compatibility sweep exposed a severe long-tail in the legacy Word
  fallback path, with `001538-2.doc` at `105010 ms`, `001538.doc` at `104216 ms`,
  `003336.doc` at `40224 ms`, and `003336-2.doc` at `31774 ms`.
- CPU profiling on `testdata/web-samples/samples/doc/001538-2.doc` showed the dominant cost was not
  OLE parsing or visible-text cleanup; it was repeated rescanning inside `utf16Strings(...)`:
  - `officeread.hasUTF16Evidence`: `63.59s` flat
  - `officeread.utf16Strings`: `30.44s` flat / `94.06s` cumulative
- Root cause: `utf16Strings(...)` built a growing `raw` byte slice and repeatedly called
  `hasUTF16Evidence(raw)` both while extending the candidate and again when flushing it. On large
  fallback ranges this turned one logical scan into many repeated full rescans of the same bytes.
- Retained change:
  - keep incremental `zeroCount` and `asciiPrintableCount` alongside `raw`
  - replace repeated `hasUTF16Evidence(raw)` rescans inside `utf16Strings(...)` with
    `hasUTF16EvidenceCounts(...)`
  - preserve the final `raw` bytes for the existing
    `looksLikeASCIIBytesMisreadAsUTF16(raw, text)` check, so output gating stays unchanged

Validation:
- targeted regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'TestCorruptOOXMLImageCRCDoesNotFailExtraction' ./...`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- before profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-001538-2-doc.pprof testdata\web-samples\samples\doc\001538-2.doc`
- after profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-001538-2-doc-after-utf16-evidence-cache.pprof testdata\web-samples\samples\doc\001538-2.doc`
- hotspot rerun after rebuilding `compatcheck.exe`:
  - `.\compatcheck.exe -json testdata\web-samples\reports\compat-doc-hotset-after-utf16-evidence-cache.json -csv testdata\web-samples\reports\compat-doc-hotset-after-utf16-evidence-cache.csv testdata\web-samples\samples\doc\001538-2.doc testdata\web-samples\samples\doc\001538.doc testdata\web-samples\samples\doc\003336-2.doc testdata\web-samples\samples\doc\003336.doc`

Observed results:
- `001538-2.doc`:
  - before single-file profile run: `96464 ms`
  - after single-file profile run: `692 ms`
- hotspot rerun with current code:
  - `001538-2.doc`: `716 ms`
  - `001538.doc`: `713 ms`
  - `003336-2.doc`: `416 ms`
  - `003336.doc`: `422 ms`
- output stability:
  - `001538*.doc` text bytes unchanged: `430692`
  - `003336*.doc` text bytes unchanged: `268880`
- full 1000-file `.doc` rerun with current code:
  - before: `millis=646683`, `avgMillis=646.68`, `over10Sec=7`, `maxMillis=105010`
  - after: `millis=220879`, `avgMillis=220.88`, `over10Sec=0`, `maxMillis=7186`
  - total time reduction: `425804 ms`
  - relative speedup: `65.84%`
- after-profile hotspot shift:
  - `utf16Strings(...)` cumulative cost fell to `20 ms`
  - dominant time moved back to general text cleanup / regex-heavy normalization, which is a much
    smaller baseline than before
- regression result:
  - `ok officeread 223.658s`

Decision:
- Retained. This change preserved output on the hotspot set, passed both targeted and full-suite
  tests, and removed the worst observed `.doc` long-tail by two orders of magnitude.

## 2026-07-05 retained: per-extraction OOXML lookup reuse for `.pptx` visible related text

- After the `.doc` UTF-16 fix, a new profile on
  `testdata/web-samples/samples/pptx/00022381.pptx` showed the dominant `.pptx` cost had shifted
  into repeated OOXML part lookup and normalization work rather than slide XML decoding itself.
- Before the retained change, the main hotspot functions were:
  - `ooxmlPartKeyCandidates(...)`
  - `ooxmlPartKeyMatches(...)`
  - `ooxmlFile(...)`
  - `ooxmlCleanPartName(...)`
  - `ooxmlEscapePartName(...)`
- Root cause: `pptxVisibleRelatedTextParts(...)` and its recursive helpers repeatedly rebuilt
  normalized part-name candidates for the same package members while traversing slide, notes,
  comment, and related OOXML relationship targets.
- Retained change:
  - add a per-call `ooxmlLookup` index built once from the package file map
  - reuse that lookup from `pptxVisibleRelatedTextParts(...)` through the recursive related-part
    traversal and relationship-target resolution helpers
  - preserve extraction semantics by keeping the same fallback matching rules and only memoizing the
    package-member lookup path

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-00022381-pptx-after-ooxml-lookup.pprof testdata\web-samples\samples\pptx\00022381.pptx`
- focused `.pptx` hotspot batch:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\compat-pptx-hotset-after-ooxml-lookup.json -csv testdata\web-samples\reports\compat-pptx-hotset-after-ooxml-lookup.csv testdata\web-samples\samples\pptx\00022381.pptx testdata\web-samples\samples\pptx\00026428.pptx testdata\web-samples\samples\pptx\00020725.pptx testdata\web-samples\samples\pptx\406730.pptx`
- full 1000-file `.pptx` rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\compat-pptx-20260705-after-ooxml-lookup.json -csv testdata\web-samples\reports\compat-pptx-20260705-after-ooxml-lookup.csv -limit 1000 testdata\web-samples\samples\pptx`

Observed results:
- focused single-file profile:
  - `00022381.pptx`: `10738 ms -> 606 ms`
- focused 4-file hotspot rerun:
  - `00022381.pptx`: `14544 ms -> 545 ms`
  - `00026428.pptx`: `6759 ms -> 823 ms`
  - `00020725.pptx`: `6505 ms -> 957 ms`
  - `406730.pptx`: `6320 ms -> 1505 ms`
- hotspot output stability:
  - all 4 files preserved the same `textBytes`
  - all 4 files preserved the same `images`
- full 1000-file `.pptx` rerun:
  - before: `millis=313850`, `avgMillis=313.85`, `over10Sec=1`, `maxMillis=14544`
  - after: `millis=139507`, `avgMillis=139.51`, `over10Sec=0`, `maxMillis=1865`
  - total time reduction: `174343 ms`
  - relative speedup: `55.55%`
- post-change profile shape:
  - the earlier OOXML lookup cluster stopped dominating the profile
  - the remaining `.pptx` cost is spread across normal XML decoding, media walking, and text cleanup
    at a much smaller baseline
- same-day `.xls` control rerun:
  - `006087.xls`: `12279 ms`
  - dominant cost still centered on
    `missingMarkdownText -> markdownBackfillContainment.visibleLineContainsLine`

Decision:
- Retained. This lookup reuse preserved output, passed the full test suite, removed the only
  `>10s` sample from the 1000-file `.pptx` sweep, and cut aggregate `.pptx` runtime by more than
  half. The clearest next target remains legacy `.xls` markdown backfill rather than more `.pptx`
  lookup work.

## 2026-07-05 retained: cached backfill containment queries in legacy `.xls` markdown completion

- After the `.pptx` lookup win, the hottest remaining spreadsheet profile was still
  `testdata/web-samples/samples/xls/006087.xls`.
- The earlier control profile on that sample showed the cost was still dominated by repeated
  substring lookups under:
  - `missingMarkdownText(...)`
  - `markdownBackfillContainment.visibleLineContainsLine(...)`
  - `strings.Contains / strings.Index`
- Root cause: `missingMarkdownText(...)` repeatedly re-asked the same coverage and containment
  questions for the original candidate line, its visible-text variant, its markdown-visible variant,
  and its escaped-table-cell variant. On large BIFF markdown backfill workloads, many of those
  normalized queries recur across rows and across repeated cell text.
- Retained change:
  - add per-call caches for:
    - `markdownBackfillCoverageCoversLine(...)`
    - `containment.tableTextContainsLine(...)`
    - `containment.visibleLineContainsLine(...)`
  - deduplicate the candidate-line variants before running the expensive coverage / containment
    checks, while preserving the exact same acceptance rules and variant set

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-backfill-query-cache.pprof testdata\web-samples\samples\xls\006087.xls`
- focused 6-file `.xls` batch:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-query-cache-xls6.json -csv testdata\web-samples\reports\perf-after-backfill-query-cache-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- focused 6-file `.xls` rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-query-cache-xls6-rerun.json -csv testdata\web-samples\reports\perf-after-backfill-query-cache-xls6-rerun.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- focused single-file rerun for the largest `.xls` sample in that set:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-query-cache-008055-rerun.json -csv testdata\web-samples\reports\perf-after-backfill-query-cache-008055-rerun.csv testdata\web-samples\samples\xls\008055.xls`
- mixed keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-query-cache-keyset.json -csv testdata\web-samples\reports\perf-after-backfill-query-cache-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- full repository regression:
  - `ok officeread 149.920s`
- focused single-file profile:
  - `006087.xls`: `12279 ms -> 5326 ms`
- profile shift on `006087.xls`:
  - before: `markdownBackfillContainment.visibleLineContainsLine(...)` about `9.94s` cumulative
  - after: `markdownBackfillContainment.visibleLineContainsLine(...)` about `3.16s` cumulative
- first 6-file `.xls` batch run:
  - `15566 ms -> 15981 ms`
  - this first pass was slightly slower, which points to some warm-cache / run-order sensitivity in
    this workload family
- repeated 6-file `.xls` batch run:
  - `15566 ms -> 15242 ms`
  - relative speedup: `2.08%`
  - `textBytes`: unchanged
- single-file rerun:
  - `008055.xls`: `5676 ms -> 5668 ms`
- mixed 30-file keyset rerun:
  - total: `50489 ms -> 48656 ms`
  - total relative speedup: `3.63%`
  - `.xls` subset: `18325 ms -> 17763 ms`
  - `.xls` subset relative speedup: `3.07%`
  - retained-vs-experiment file outputs: `NO_DIFF`

Decision:
- Retained. The first 6-file pass was noisy, but the repeat run, the broad mixed keyset, the
  single-file hotspot profile, and the output-stability checks all point in the same direction:
  this change preserves extraction results and reduces repeated containment-query cost in the
  dominant remaining `.xls` markdown backfill path.

## 2026-07-05 rejected: pruning table rows from visible-line containment pool

- After the query-cache win, tried a narrower follow-up change inside
  `markdownBackfillContainmentSet(...)`: exclude markdown table rows from `visibleLines`, on the
  theory that `tableTextContainsLine(...)` already covers table content and the extra
  `visibleLineContainsLine(...)` search space was redundant.
- This looked promising on the single hottest `.xls` sample, but it did not hold up on the retained
  hotspot batch.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-backfill-query-cache-and-visible-prune.pprof testdata\web-samples\samples\xls\006087.xls`
- focused 6-file `.xls` batch:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-backfill-query-cache-and-visible-prune-xls6.json -csv testdata\web-samples\reports\perf-after-backfill-query-cache-and-visible-prune-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- single-file hotspot:
  - `006087.xls`: `5326 ms -> 1832 ms`
- but retained 6-file `.xls` hotspot batch:
  - query-cache retained baseline: `15242 ms`
  - visible-line prune experiment: `17336 ms`
- output stayed stable on the batch, but the slowdown on the retained workload set was too large to
  justify keeping a sample-specific win.

Decision:
- Reverted. This one overfit the hottest single sample and made the broader retained `.xls` hotspot
  set worse, so it is not a durable optimization.

## 2026-07-05 rejected: visible-line 4-byte n-gram guard

- Tried adding a prefilter inside `markdownBackfillContainment.visibleLineContainsLine(...)` that
  indexed every 4-byte window from `visibleLines` and rejected queries whose first 4 bytes were not
  present before falling through to the existing `strings.Contains(...)` check.
- The idea was to cheaply eliminate obviously impossible visible-line containment checks without
  changing the matching rule itself.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-visible-ngram-guard.pprof testdata\web-samples\samples\xls\006087.xls`
- focused 6-file `.xls` batch:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-visible-ngram-guard-xls6.json -csv testdata\web-samples\reports\perf-after-visible-ngram-guard-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- single-file hotspot:
  - `006087.xls`: `5326 ms -> 5417 ms`
- retained 6-file `.xls` hotspot batch:
  - query-cache retained baseline: `15242 ms`
  - visible-line 4-byte n-gram guard: `18283 ms`
- output stayed stable, but the additional index build and lookup overhead outweighed any short-circuit
  benefit on the retained workload mix.

Decision:
- Reverted. This filter looked tidy, but it was a straightforward slowdown on both the hottest
  single sample and the retained `.xls` hotspot batch.

## 2026-07-05 rejected: visible-line candidate buckets by 4-byte prefix

- Tried a more exact replacement for the previous global prefilter: instead of searching
  `visibleJoined` directly, build a map from each 4-byte substring prefix to the visible lines that
  contain it, then answer `visibleLineContainsLine(...)` by checking only the candidate lines in the
  matching bucket.
- This preserves semantics more faithfully than a global existence filter because it still falls
  back to `strings.Contains(...)` on real visible lines; it only shrinks the candidate set.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-visible-line-buckets.pprof testdata\web-samples\samples\xls\006087.xls`
- focused 6-file `.xls` batch:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-visible-line-buckets-xls6.json -csv testdata\web-samples\reports\perf-after-visible-line-buckets-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- single-file hotspot:
  - `006087.xls`: `5326 ms -> 2326 ms`
- retained 6-file `.xls` hotspot batch:
  - query-cache retained baseline: `15242 ms`
  - visible-line candidate buckets: `16959 ms`
- output remained stable, but the index build cost and per-line candidate bookkeeping still lost on
  the broader retained hotspot mix.

Decision:
- Reverted. This one again overfit the single hottest sample while making the retained `.xls`
  workload set slower overall.

## 2026-07-05 rejected: defer visible-line containment until after all cheaper variant checks

- Tried a call-order-only rewrite in `missingMarkdownText(...)`: for each candidate line, first run
  `coverageContains(...)` and `tableTextContains(...)` across every deduplicated variant, and only if
  none of those cheaper predicates hit, run the more expensive
  `visibleLineContainsLine(...)` checks.
- This preserved the exact predicate set and only changed evaluation order, with the hope that later
  variants would often satisfy a cheap check and avoid some expensive visible-line scans.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-visible-last-query-order.pprof testdata\web-samples\samples\xls\006087.xls`
- focused 6-file `.xls` batch:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-visible-last-query-order-xls6.json -csv testdata\web-samples\reports\perf-after-visible-last-query-order-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- full repository regression:
  - `ok officeread 142.106s`
- single-file hotspot:
  - `006087.xls`: `5326 ms -> 5364 ms`
- retained 6-file `.xls` hotspot batch:
  - query-cache retained baseline: `15242 ms`
  - deferred visible-line containment: `17369 ms`
- output stayed stable, but the reordered checks still made the retained hotspot set slower.

Decision:
- Reverted. Even though the rewrite preserved semantics and improved full-test wall time in this one
  run, it did not improve the retained `.xls` hotspot workloads that are driving the optimization
  effort.

## 2026-07-05 rejected: skip escaped table-cell variant when source line lacks table escapes

- Tried guarding the `escapeMarkdownTableCell(...)`-derived variant in both
  `missingMarkdownText(...)` and `markdownBackfillExactSetContainsLine(...)`, only generating that
  extra normalized form when the candidate line contained `\`, `|`, or line-break characters that
  could actually change under table-cell escaping.
- This looked like a low-risk reduction because ordinary plain lines should not gain a new visible
  form from table-cell escaping.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-escaped-variant-guard.pprof testdata\web-samples\samples\xls\006087.xls`
- focused 6-file `.xls` batch:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-escaped-variant-guard-xls6.json -csv testdata\web-samples\reports\perf-after-escaped-variant-guard-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- single-file hotspot:
  - `006087.xls`: `5326 ms -> 5375 ms`
- retained 6-file `.xls` hotspot batch:
  - query-cache retained baseline: `15242 ms`
  - escaped-variant guard: `16657 ms`
- output remained stable, but the guard was still a net loss on the retained workload mix.

Decision:
- Reverted. This branch is not hot enough in the retained `.xls` workloads to justify the extra
  conditionals and altered control flow.

## 2026-07-05 rejected: skip comparable normalization when a line should already be stable

- Tried using the existing `markdownBackfillComparableMayDiffer(...)` heuristic as a fast-path guard
  in two places:
  - `markdownBackfillCoverageCoversLine(...)`
  - `markdownBackfillContainment.tableTextContainsLine(...)`
- The intended win was to avoid building comparable-normalized text for plain lines that should not
  change under the comparable-text normalization path.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-comparable-maydiffer-guard.pprof testdata\web-samples\samples\xls\006087.xls`
- focused 6-file `.xls` batch:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-comparable-maydiffer-guard-xls6.json -csv testdata\web-samples\reports\perf-after-comparable-maydiffer-guard-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- full repository regression:
  - `ok officeread 153.664s`
- single-file hotspot:
  - `006087.xls`: `5326 ms -> 5571 ms`
- retained 6-file `.xls` hotspot batch:
  - query-cache retained baseline: `15242 ms`
  - comparable-maydiffer guard: `18401 ms`
- output stayed stable, but the extra heuristic check was a pure slowdown on the workloads that
  matter here.

Decision:
- Reverted. The comparable-normalization path is not expensive in the right places for this guard to
  pay off on the retained `.xls` hotspot mix.

## 2026-07-05 rejected: cell-only BIFF markdown-backfill input

- Tried reducing the `.xls` markdown-backfill source in `biffMarkdown(...)`: instead of feeding
  `missingMarkdownText(...)` the full visible BIFF text stream, only feed cell-derived BIFF text and
  exclude sheet names, header/footer text, and comment text that are already represented elsewhere
  in the generated markdown.
- The intent was to shrink the backfill search space before the retained query-cache logic even runs.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-cell-only-backfill-input.pprof testdata\web-samples\samples\xls\006087.xls`
- focused 6-file `.xls` batch:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-cell-only-backfill-input-xls6.json -csv testdata\web-samples\reports\perf-after-cell-only-backfill-input-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- mixed 30-file keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-cell-only-backfill-input-keyset.json -csv testdata\web-samples\reports\perf-after-cell-only-backfill-input-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- full repository regression:
  - `ok officeread 150.084s`
- single-file hotspot:
  - `006087.xls`: `5326 ms -> 5322 ms`
- retained 6-file `.xls` hotspot batch:
  - query-cache retained baseline: `15242 ms`
  - cell-only backfill input: `20070 ms`
- mixed 30-file keyset rerun:
  - retained baseline total: `48656 ms`
  - cell-only backfill input total: `54907 ms`
  - retained baseline `.xls` subset: `17763 ms`
  - cell-only backfill input `.xls` subset: `21835 ms`
- representative regressions from the keyset rerun:
  - `008055.xls`: `5872 ms -> 7745 ms`
  - `002505.xls`: `2264 ms -> 3091 ms`
  - `223624.docx`: `2641 ms -> 3352 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged for every keyset file

Decision:
- Reverted. This reduction overfit the single hottest `.xls` sample and made both the retained
  hotspot batch and the broader keyset rerun slower, without buying any output simplification that
  matters to the retained baseline.

## 2026-07-05 rejected: visible-line global byte-mask prefilter

- Tried adding a very cheap prefilter inside
  `markdownBackfillContainment.visibleLineContainsLine(...)`: build a 256-bit byte-presence mask for
  the joined visible markdown lines, then reject any candidate line containing a byte that never
  appears in the visible corpus before falling through to `strings.Contains(...)`.
- This was intentionally lighter than the previously rejected 4-byte n-gram and prefix-bucket
  experiments. The hope was to prune impossible substring probes with only a short per-candidate
  byte scan.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-visible-byte-mask.pprof testdata\web-samples\samples\xls\006087.xls`
- focused 6-file `.xls` batch:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-visible-byte-mask-xls6.json -csv testdata\web-samples\reports\perf-after-visible-byte-mask-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- mixed 30-file keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-visible-byte-mask-keyset.json -csv testdata\web-samples\reports\perf-after-visible-byte-mask-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- full repository regression:
  - `ok officeread 127.282s`
- single-file hotspot:
  - `006087.xls`: `5326 ms -> 5159 ms`
- retained 6-file `.xls` hotspot batch:
  - query-cache retained baseline: `15242 ms`
  - visible-byte-mask prefilter: `17346 ms`
- mixed 30-file keyset rerun:
  - retained baseline total: `48656 ms`
  - visible-byte-mask prefilter total: `49392 ms`
  - retained baseline `.xls` subset: `17763 ms`
  - visible-byte-mask prefilter `.xls` subset: `20311 ms`
  - retained baseline `.xlsx` subset: `23991 ms`
  - visible-byte-mask prefilter `.xlsx` subset: `22409 ms`
- representative file deltas:
  - `008055.xls`: `5872 ms -> 7017 ms`
  - `019088.xls`: `1699 ms -> 2242 ms`
  - `019089.xls`: `1675 ms -> 2258 ms`
  - `00003763.docx`: `4261 ms -> 3923 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged for every keyset file

Decision:
- Reverted. Even though the global byte-mask filter helped the hottest single `.xls` file and some
  mixed-workload `.xlsx`/`.docx` inputs, it still regressed the retained `.xls` hotspot mix and the
  overall keyset rerun, so it does not beat the retained baseline.

## 2026-07-05 rejected: deduplicate containment haystacks

- Tried shrinking the joined haystacks used by `markdownBackfillContainmentSet(...)` by removing
  duplicate lines before building:
  - `tableRaw`
  - `tableVisible`
  - `tableComparable`
  - `visibleLines`
- This looked promising because the hottest `.xls` markdown sample (`006087.xls`) contains thousands
  of repeated visible/table lines, so removing duplicates should reduce the bytes scanned by
  `strings.Contains(...)` inside `visibleLineContainsLine(...)` and `tableTextContainsLine(...)`.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-containment-dedup.pprof testdata\web-samples\samples\xls\006087.xls`
- focused 6-file `.xls` batch:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-containment-dedup-xls6.json -csv testdata\web-samples\reports\perf-after-containment-dedup-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- mixed 30-file keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\perf-after-containment-dedup-keyset.json -csv testdata\web-samples\reports\perf-after-containment-dedup-keyset.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls testdata\web-samples\samples\xls\002505.xls testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx testdata\web-samples\samples\xlsx\00014878.xlsx testdata\web-samples\samples\xlsx\00014984.xlsx testdata\web-samples\samples\xlsx\00019224.xlsx testdata\web-samples\samples\xlsx\00017009.xlsx testdata\web-samples\samples\xlsx\00014957.xlsx testdata\web-samples\samples\xlsx\00014999.xlsx testdata\web-samples\samples\xlsx\00014970.xlsx testdata\web-samples\samples\xlsx\00014857.xlsx testdata\web-samples\samples\xlsx\00018489.xlsx testdata\web-samples\samples\xlsx\142143-2.xlsx testdata\web-samples\samples\xlsx\00013260.xlsx testdata\web-samples\samples\xlsx\142143.xlsx testdata\web-samples\samples\xlsx\219417.xlsx testdata\web-samples\samples\xlsx\212946-2.xlsx testdata\web-samples\samples\xlsx\00013939.xlsx testdata\web-samples\samples\xlsx\219417-2.xlsx testdata\web-samples\samples\xlsx\00010400.xlsx testdata\web-samples\samples\docx\223624.docx testdata\web-samples\samples\xlsx\00017663.xlsx testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\xlsx\00012647.xlsx`

Observed results:
- full repository regression:
  - `ok officeread 148.561s`
- repeated-line evidence on `006087.xls` markdown:
  - visible lines: `54601 total`, `48582 unique`, `6019 duplicates`
  - table lines: `49999 total`, `43980 unique`, `6019 duplicates`
- single-file hotspot:
  - `006087.xls`: `5326 ms -> 4778 ms`
- retained 6-file `.xls` hotspot batch:
  - query-cache retained baseline: `15242 ms`
  - containment dedup: `17657 ms`
- mixed 30-file keyset rerun:
  - retained baseline total: `48656 ms`
  - containment dedup total: `53685 ms`
  - retained baseline `.xls` subset: `17763 ms`
  - containment dedup `.xls` subset: `20827 ms`
  - retained baseline `.xlsx` subset: `23991 ms`
  - containment dedup `.xlsx` subset: `26034 ms`
- representative regressions:
  - `008055.xls`: `5872 ms -> 6857 ms`
  - `002505.xls`: `2264 ms -> 2778 ms`
  - `016161.xls`: `2417 ms -> 2748 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged for every keyset file

Decision:
- Reverted. Removing duplicate haystack lines helped the single hottest `.xls` sample, but the
  broader retained hotspot batch and mixed keyset both got slower, so this does not earn its place
  in the retained baseline.

## 2026-07-05 retained: repeat-aware compatcheck timing

- Added `-repeat N` to `cmd/compatcheck` so performance reruns can record median timings instead of
  trusting a single wall-time sample.
- When `N > 1`:
  - each file is extracted `N` times
  - reported `millis` becomes the median of those runs
  - JSON/CSV also include `minMillis`, `maxMillis`, and raw `runs`
  - repeat runs verify `textBytes/images` stability across attempts

Validation:
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- first repeat-aware `.xls` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-repeat3-current-xls6.json -csv testdata\web-samples\reports\perf-repeat3-current-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- full repository regression:
  - `ok officeread (cached)`
- repeat-aware 6-file `.xls` hotspot baseline:
  - median total: `19093 ms`
  - `008055.xls`: median `6847 ms`, range `5847..6873`
  - `013623.xls`: median `2759 ms`, range `2464..2836`
  - `016161.xls`: median `3024 ms`, range `2947..3104`
  - `018548.xls`: median `2133 ms`, range `2018..2194`
  - `019088.xls`: median `2161 ms`, range `2087..2215`
  - `019089.xls`: median `2169 ms`, range `2066..2275`
- output stability:
  - no per-run `textBytes/images` mismatches across the repeated hotspot batch

Decision:
- Retained. This does not change extraction behavior, but it materially improves the reliability of
  future performance decisions on the noisy `.xls` hotspot set by making median-based comparisons
  easy and auditable.

## 2026-07-05 retained: standard median for even repeat counts

- Followed up on the new `compatcheck -repeat N` support by correcting the median calculation for
  even `N`.
- The first implementation used the upper middle element after sorting. That is easy to compute, but
  it is not the conventional median for even sample counts and can systematically bias repeat-aware
  comparisons upward on noisy workloads.
- Updated the tool so that:
  - odd `N` still uses the middle element
  - even `N` now uses the integer average of the two middle runs

Validation:
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused even-repeat smoke check:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-repeat2-baseline-rune-count-xls6-even-median.json -csv testdata\web-samples\reports\perf-repeat2-baseline-rune-count-xls6-even-median.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- full repository regression:
  - `ok officeread (cached)`
- even-repeat `.xls` hotspot sample medians now land between the two runs instead of snapping to the
  slower upper-middle value:
  - `008055.xls`: runs `[6119, 5975]`, median `6047`
  - `013623.xls`: runs `[2243, 2194]`, median `2218`
  - `016161.xls`: runs `[2425, 2508]`, median `2466`
  - `018548.xls`: runs `[1733, 1778]`, median `1755`
  - `019088.xls`: runs `[1748, 1708]`, median `1728`
  - `019089.xls`: runs `[1774, 1706]`, median `1740`

Decision:
- Retained. This keeps repeat-aware performance reporting mathematically conventional and reduces the
  chance that near-tie experiments are accepted or rejected due to an avoidable summary bias.

## 2026-07-05 rejected: fixed-array variant assembly in missingMarkdownText

- Tried removing the per-line small-slice allocation in `missingMarkdownText(...)` by replacing:
  - `variants := []string{line}`
  - repeated `append(...)`
  - duplicate checks over the growing slice
- with a fixed `[4]string` array plus an explicit count and duplicate scan.
- The intended win was to reduce tiny allocations and slice bookkeeping in the hottest `.xls`
  markdown-backfill loop without changing any matching semantics.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-variant-array.pprof testdata\web-samples\samples\xls\006087.xls`
- repeat-aware 6-file `.xls` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-repeat3-after-variant-array-xls6.json -csv testdata\web-samples\reports\perf-repeat3-after-variant-array-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-repeat2-after-variant-array-keyset.json -csv testdata\web-samples\reports\perf-repeat2-after-variant-array-keyset.csv ...same retained keyset...`

Observed results:
- full repository regression:
  - `ok officeread 154.926s`
- single-file hotspot:
  - `006087.xls`: `5326 ms -> 5460 ms`
- repeat-aware 6-file `.xls` hotspot batch:
  - retained baseline median total: `19093 ms`
  - fixed-array variant assembly: `18374 ms`
  - relative improvement on this narrow batch: `3.77%`
- broader repeat-aware keyset:
  - retained baseline total: `51517 ms`
  - fixed-array variant assembly: `54008 ms`
  - retained baseline `.xls` subset: `17264 ms`
  - fixed-array variant assembly `.xls` subset: `20639 ms`
- representative regressions on the broader keyset:
  - `008055.xls`: `5521 ms -> 5963 ms`
  - `016161.xls`: `2411 ms -> 3127 ms`
  - `018548.xls`: `1640 ms -> 2800 ms`
  - `00003763.docx`: `3889 ms -> 4175 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged across the repeat-aware keyset

Decision:
- Reverted. This rewrite helped the repeat-aware `.xls6` hotspot batch, but it lost decisively on
  the broader retained keyset and therefore does not qualify as a retained optimization.

## 2026-07-05 rejected: progressive variant checking

- Tried changing the per-line variant evaluation strategy in `missingMarkdownText(...)`.
- Baseline behavior:
  - build all candidate variants first (`line`, visible form, markdown-visible form, escaped-table
    visible form)
  - then test them in order against coverage, table containment, and visible-line containment
- Experiment:
  - check the base line immediately
  - only compute the next variant if the previous one did not already cover the line
  - stop generating later variants as soon as one variant is covered
- The intended win was to reduce variant construction work and skip some downstream containment
  probes for lines that are already covered by the earlier forms.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profile:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-progressive-variants.pprof testdata\web-samples\samples\xls\006087.xls`
- repeat-aware 6-file `.xls` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-repeat3-after-progressive-variants-xls6.json -csv testdata\web-samples\reports\perf-repeat3-after-progressive-variants-xls6.csv testdata\web-samples\samples\xls\008055.xls testdata\web-samples\samples\xls\013623.xls testdata\web-samples\samples\xls\016161.xls testdata\web-samples\samples\xls\018548.xls testdata\web-samples\samples\xls\019088.xls testdata\web-samples\samples\xls\019089.xls`

Observed results:
- full repository regression:
  - `ok officeread 154.926s`
- single-file hotspot:
  - `006087.xls`: `5326 ms -> 18652 ms`
- repeat-aware 6-file `.xls` hotspot batch:
  - retained baseline median total: `19093 ms`
  - progressive variant checking: `37584 ms`
  - representative medians:
    - `008055.xls`: `6847 ms -> 13924 ms`
    - `013623.xls`: `2759 ms -> 8583 ms`
    - `016161.xls`: `3024 ms -> 8359 ms`
    - `019089.xls`: `2169 ms -> 2569 ms`
- output stability:
  - `textBytes` / `images` remained stable

Decision:
- Reverted immediately. This change was a severe regression on the hottest `.xls` sample and the
  retained repeat-aware hotspot batch, so it is not a viable direction.

## 2026-07-05 retained: hotspot extraction benchmarks

- Added a committed benchmark harness in [extract_bench_test.go](/D:/workprj/officeread/extract_bench_test.go)
  for the current hottest legacy `.xls` extraction cases:
  - `006087.xls`
  - `008055.xls`
  - `016161.xls`
- This benchmark runs full `Extract(...)` on real retained hotspot files and reports both wall time
  and allocations, giving future optimization work a repeatable `go test -bench` target alongside
  the ad hoc compatcheck reruns.

Validation:
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- benchmark smoke run:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`

Observed results:
- full repository regression:
  - `ok officeread 154.825s`
- benchmark smoke baseline:
  - `BenchmarkExtractXLSHotspots/006087.xls`: `5033971000 ns/op`, `714237984 B/op`, `6098297 allocs/op`
  - `BenchmarkExtractXLSHotspots/008055.xls`: `7935391100 ns/op`, `1977264776 B/op`, `6044348 allocs/op`
  - `BenchmarkExtractXLSHotspots/016161.xls`: `3613330100 ns/op`, `890202280 B/op`, `3389250 allocs/op`

Decision:
- Retained. This does not change extraction behavior, but it materially improves repeatability for
  future `.xls` optimization work by putting representative hotspot timings and allocation counts
  under the normal Go benchmark workflow.

## 2026-07-05 rejected: in-place normalized line reuse

- Tried replacing the per-raw-line `candidateLineCache` map in `missingMarkdownText(...)` with an
  in-place normalization pass over `lines`:
  - split the source text once
  - rewrite each `lines[i]` to its normalized candidate form
  - reuse that normalized slice for both exact-set coverage checks and the main missing-line pass
- The intended win was to remove a large `raw -> normalized` cache on hotspot files where nearly all
  candidate lines are unique.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused benchmark check:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSHotspots/006087.xls$' -benchmem -benchtime=1x ./`
- repeat-aware 6-file `.xls` hotspot reruns:
  - `compatcheck -repeat 3 ...`
  - `compatcheck -repeat 2 ...`
- broader repeat-aware keyset rerun:
  - `compatcheck -repeat 2 ...retained keyset...`

Observed results:
- full repository regression:
  - `ok officeread 142.685s`
- focused benchmark on `006087.xls`:
  - baseline: `5033971000 ns/op`, `714237984 B/op`, `6098297 allocs/op`
  - experiment: `5218890900 ns/op`, `709809008 B/op`, `6116311 allocs/op`
- repeat-aware `.xls6` hotspot batch:
  - `repeat=3` median total: `19093 ms -> 15556 ms`
  - but `repeat=2` median total: `15954 ms -> 20178 ms`
- broader repeat-aware keyset:
  - total: `51517 ms -> 49659 ms`
  - `.xls` subset: `17264 ms -> 17642 ms`
  - `.xlsx` subset: `27830 ms -> 25445 ms`
  - `.docx` subset: `6423 ms -> 6572 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. Even though this branch helped the repeat-aware `.xls6` batch and improved total
  keyset wall time through `.xlsx` gains, it regressed the repeat-aware even-count `.xls6` batch,
  slightly worsened the broader keyset `.xls` subset, and did not improve the focused `006087.xls`
  benchmark. The signal was too inconsistent to treat as a retained `.xls` optimization.

## 2026-07-05 rejected: upfront biffTextParts slice preallocation

- Tried reducing `appendPart(...)` growth churn in `biffTextParts(...)` by preallocating the `parts`
  slice from the BIFF payload size:
  - `parts := make([]biffTextPart, 0, estimateBIFFTextPartCapacity(len(data)))`
- This was motivated by `alloc_space` profiles where `officeread.biffTextParts.func1` was a top
  allocator on both `006087.xls` and `008055.xls`, with most of that cost sitting on
  `parts = append(parts, part)`.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused benchmark checks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSHotspots/(006087.xls|008055.xls)$' -benchmem -benchtime=1x ./`
- repeat-aware `.xls6` hotspot rerun:
  - `compatcheck -repeat 3 ...`
- broader repeat-aware keyset rerun:
  - `compatcheck -repeat 2 ...retained keyset...`

Observed results:
- full repository regression:
  - `ok officeread 198.877s`
- focused benchmarks:
  - `006087.xls`
    - baseline: `5033971000 ns/op`, `714237984 B/op`, `6098297 allocs/op`
    - prealloc: `5251827200 ns/op`, `672588824 B/op`, `6098127 allocs/op`
  - `008055.xls`
    - baseline: `7935391100 ns/op`, `1977264776 B/op`, `6044348 allocs/op`
    - prealloc: `6073816100 ns/op`, `1540425272 B/op`, `6044260 allocs/op`
- repeat-aware `.xls6` hotspot batch:
  - baseline median total: `19093 ms`
  - prealloc median total: `16237 ms`
- broader repeat-aware keyset:
  - total: `51517 ms -> 60898 ms`
  - `.xls` subset: `17264 ms -> 25346 ms`
  - `.xlsx` subset: `27830 ms -> 27060 ms`
  - `.docx` subset: `6423 ms -> 8492 ms`
- representative regressions on the broader keyset:
  - `008055.xls`: `5521 ms -> 8146 ms`
  - `018548.xls`: `1640 ms -> 3045 ms`
  - `00003763.docx`: `3889 ms -> 4935 ms`

Decision:
- Reverted. The preallocation heuristic looked great on the narrow hotspot benchmarks and `.xls6`
  rerun, but it catastrophically regressed the broader retained keyset, so it is not safe to keep.

## 2026-07-05 rejected: progressive `missingMarkdownText` variant checking

- Tried reducing per-line work in `missingMarkdownText(...)` by changing the coverage check order:
  - check the normalized base line first
  - only compute `markdownBackfillVisibleText(line)` if the base line was not already covered
  - only compute `markdownVisibleLineText(line)` if both earlier checks missed
  - only build/check the escaped table-cell visible form after the earlier variants missed
- This also removed the temporary `variants := []string{...}` slice allocation from the hot loop.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-missing-backfill-variant-elision-xls6.json -csv testdata\web-samples\reports\perf-exp-missing-backfill-variant-elision-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-missing-backfill-variant-elision-keyset.json -csv testdata\web-samples\reports\perf-exp-missing-backfill-variant-elision-keyset.csv ...retained keyset...`

Observed results:
- full repository regression:
  - `ok officeread 157.729s`
- hotspot benchmarks:
  - `006087.xls`: `5573310400 ns/op -> 5652769900 ns/op`
  - `008055.xls`: `7905695900 ns/op -> 8107040500 ns/op`
  - `016161.xls`: `3466805000 ns/op -> 3571272700 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - baseline median total: `23111 ms`
  - experiment median total: `22637 ms`
- broader repeat-aware keyset:
  - total: `51517 ms -> 55521 ms`
  - `.docx` subset: `6423 ms -> 7533 ms`
  - `.xls` subset: `17264 ms -> 20266 ms`
  - `.xlsx` subset: `27830 ms -> 27722 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. The small `.xls6` improvement was not real enough to survive the wider retained
  keyset, and the benchmark rerun moved the hotspot files in the wrong direction.

## 2026-07-05 rejected: escaped table-cell fallback guard

- Tried a narrower follow-up after reverting the progressive branch:
  - keep the existing `variants` flow
  - but only compute `markdownBackfillVisibleText(escapeMarkdownTableCell(line))` when the line
    actually contains `|`
- The goal was to avoid expensive table-cell escaping work on ordinary prose lines while leaving
  the rest of the backfill logic untouched.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-escaped-table-guard-xls6.json -csv testdata\web-samples\reports\perf-exp-escaped-table-guard-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-escaped-table-guard-keyset.json -csv testdata\web-samples\reports\perf-exp-escaped-table-guard-keyset.csv ...retained keyset...`

Observed results:
- focused regression suite:
  - `ok officeread 11.259s`
- hotspot benchmarks:
  - `006087.xls`: `5573310400 ns/op -> 4978822900 ns/op`
  - `008055.xls`: `7905695900 ns/op -> 7214288600 ns/op`
  - `016161.xls`: `3466805000 ns/op -> 3311916200 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - baseline median total: `23111 ms`
  - experiment median total: `20124 ms`
- broader repeat-aware keyset:
  - total: `51517 ms -> 58143 ms`
  - `.docx` subset: `6423 ms -> 6950 ms`
  - `.xls` subset: `17264 ms -> 19992 ms`
  - `.xlsx` subset: `27830 ms -> 31201 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. This looked excellent on the hotspot batch and on the single-run benchmark rerun, but
  the retained repeat-aware keyset got worse across every format family, so the guard is not safe
  to keep.

## 2026-07-05 rejected: move short table-cell hits out of `visibleLines`

- Re-profiled the retained `006087.xls` baseline first:
  - source text: `41347` non-empty lines, average line length `6.02`, only `5` single-character
    lines, no `|` characters
  - generated markdown: `54601` non-empty lines, `49999` contain `|`
  - CPU profile still showed
    `missingMarkdownText -> markdownBackfillContainment.visibleLineContainsLine -> strings.Contains`
    as the dominant cost
- That suggested a structural duplication in `markdownBackfillContainmentSet(...)`:
  - table rows already populate `tableRaw` / `tableVisible` / `tableComparable`
  - but the same rows also populate `visibleLines`, which means later coverage checks scan large
    table-row strings again through `visibleLineContainsLine`
- Tried removing that duplication in two variants:
  1. stop adding table rows to `visibleLines`, and add all visible table cells to a new
     `tableCellExact` set checked before the `< 12 runes` table guard
  2. same idea, but only cache short visible table cells (`< 12` runes), since longer lines can
     already go through the existing joined-table containment path

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmark reruns:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot reruns:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-tablecell-exact-visible-split-xls6.json -csv testdata\web-samples\reports\perf-exp-tablecell-exact-visible-split-xls6.csv ...`
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-tablecell-short-exact-visible-split-xls6.json -csv testdata\web-samples\reports\perf-exp-tablecell-short-exact-visible-split-xls6.csv ...`
- broader repeat-aware keyset reruns:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-tablecell-exact-visible-split-keyset.json -csv testdata\web-samples\reports\perf-exp-tablecell-exact-visible-split-keyset.csv ...retained keyset...`
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-tablecell-short-exact-visible-split-keyset.json -csv testdata\web-samples\reports\perf-exp-tablecell-short-exact-visible-split-keyset.csv ...retained keyset...`

Observed results:
- focused regression suite:
  - all-cells branch: `ok officeread 11.559s`
  - short-cells-only branch: `ok officeread 12.041s`
- hotspot benchmarks, all-cells branch:
  - `006087.xls`: `5573310400 ns/op -> 2034079400 ns/op`
  - `008055.xls`: `7905695900 ns/op -> 7753599200 ns/op`
  - `016161.xls`: `3466805000 ns/op -> 3403538000 ns/op`
- hotspot benchmarks, short-cells-only branch:
  - `006087.xls`: `5573310400 ns/op -> 2270476700 ns/op`
  - `008055.xls`: `7905695900 ns/op -> 6258198200 ns/op`
  - `016161.xls`: `3466805000 ns/op -> 2444611900 ns/op`
- repeat-aware `.xls6` hotspot batches:
  - retained same-day baseline: `23111 ms`
  - all-cells branch: `19458 ms`
  - short-cells-only branch: `16658 ms`
- broader repeat-aware keyset, all-cells branch:
  - total: `51517 ms -> 58602 ms`
  - `.docx` subset: `6423 ms -> 6391 ms`
  - `.xls` subset: `17264 ms -> 21689 ms`
  - `.xlsx` subset: `27830 ms -> 30522 ms`
- broader repeat-aware keyset, short-cells-only branch:
  - total: `51517 ms -> 59085 ms`
  - `.docx` subset: `6423 ms -> 6997 ms`
  - `.xls` subset: `17264 ms -> 21772 ms`
  - `.xlsx` subset: `27830 ms -> 30316 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged for both branches

Decision:
- Reverted both branches. The structural change clearly helps the narrow `.xls` hotspot workloads,
  but once table rows stop contributing to `visibleLines`, the retained wider keyset consistently
  regresses, especially on `.xls` and `.xlsx`. This is another case where a dramatic hotspot win
  does not translate into a safe retained optimization.

## 2026-07-05 rejected: short table-cell exact hits without containment restructuring

- Tried a narrower follow-up after the `visibleLines` restructuring experiments failed:
  - keep the retained `markdownBackfillContainmentSet(...)` layout unchanged
  - add a short visible table-cell exact set while walking markdown table rows
  - consult that exact set before the existing `< 12 runes` early return in
    `tableTextContainsLine(...)`
- The goal was to let short cell queries bypass `visibleLineContainsLine(...)` without removing
  table rows from `visibleLines`.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-short-tablecell-exact-xls6.json -csv testdata\web-samples\reports\perf-exp-short-tablecell-exact-xls6.csv ...`

Observed results:
- focused regression suite:
  - `ok officeread 13.733s`
- hotspot benchmarks:
  - `006087.xls`: `5573310400 ns/op -> 6289948800 ns/op`
  - `008055.xls`: `7905695900 ns/op -> 7716413100 ns/op`
  - `016161.xls`: `3466805000 ns/op -> 3567597900 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - baseline median total: `23111 ms`
  - experiment median total: `24728 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged on the hotspot batch

Decision:
- Reverted immediately. This branch was already worse than the retained baseline on the focused
  `.xls6` hotspot rerun, so it did not earn a broader repeat-aware keyset pass.

## 2026-07-05 rejected: raise `visibleLineContainsLine` minimum query length

- Tried the cheapest possible query-side guard on the main hotspot:
  - keep all existing containment structures and caches
  - only change `visibleLineContainsLine(...)` so it skips substring searches for queries shorter
    than 6 runes instead of shorter than 4 runes
- The hope was that the retained `.xls` hotspot batch was spending too much time searching for very
  short lines that were mostly noisy markdown backfill candidates.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-visible-min6-xls6.json -csv testdata\web-samples\reports\perf-exp-visible-min6-xls6.csv ...`

Observed results:
- focused regression suite:
  - `ok officeread 13.364s`
- hotspot benchmarks:
  - `006087.xls`: `5573310400 ns/op -> 7056552700 ns/op`
  - `008055.xls`: `7905695900 ns/op -> 9562072200 ns/op`
  - `016161.xls`: `3466805000 ns/op -> 9245684500 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - baseline median total: `23111 ms`
  - experiment median total: `26823 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged on the hotspot batch

Decision:
- Reverted immediately. Even without changing output, this guard made the retained hotspot batch far
  slower, so it was not worth spending a broader repeat-aware keyset run on it.

## 2026-07-05 rejected: lower `tableTextContainsLine` short-query threshold

- Tried a more direct way to keep short `.xls` cell candidates out of `visibleLineContainsLine(...)`:
  - keep every retained containment structure unchanged
  - only lower the early-return threshold in `tableTextContainsLine(...)` from 12 runes to 8 runes
- The idea was that many real spreadsheet cell values are short, so letting them use the existing
  table containment path might reduce expensive substring searches against `visibleJoined`.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-table-min8-xls6.json -csv testdata\web-samples\reports\perf-exp-table-min8-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-table-min8-keyset.json -csv testdata\web-samples\reports\perf-exp-table-min8-keyset.csv ...retained keyset...`

Observed results:
- focused regression suite:
  - `ok officeread 11.786s`
- hotspot benchmarks:
  - `006087.xls`: `5573310400 ns/op -> 5069138700 ns/op`
  - `008055.xls`: `7905695900 ns/op -> 5947403400 ns/op`
  - `016161.xls`: `3466805000 ns/op -> 2719632600 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - baseline median total: `23111 ms`
  - experiment median total: `21219 ms`
- broader repeat-aware keyset:
  - total: `51517 ms -> 59328 ms`
  - `.docx` subset: `6423 ms -> 8263 ms`
  - `.xls` subset: `17264 ms -> 24039 ms`
  - `.xlsx` subset: `27830 ms -> 27026 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. This threshold change was one of the rare branches that clearly improved the focused
  `.xls` hotspot slice without changing output, but it still overfit the narrow workload and
  regressed the retained broader keyset too heavily to keep.

## 2026-07-05 retained: `.xls`-only escaped table-cell fallback guard

- Reused one of the strongest local signals from the earlier generic backfill experiments, but
  narrowed it to the legacy Excel path only:
  - added `missingMarkdownTextXLS(...)`, which uses the same markdown backfill logic as
    `missingMarkdownText(...)`
  - but when evaluating the escaped table-cell fallback variant, it first checks whether the BIFF
    source line actually contains `|`
  - `biffMarkdown(...)` now uses this `.xls`-specific path, while every other format continues to
    use the generic markdown backfill behavior unchanged
- This targets the observed hotspot shape directly:
  - `006087.xls` source text lines contain no `|`
  - the expensive escaped-table fallback was therefore pure wasted work on a large fraction of the
    legacy Excel backfill loop

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-xls-pipe-escaped-guard-xls6.json -csv testdata\web-samples\reports\perf-exp-xls-pipe-escaped-guard-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-xls-pipe-escaped-guard-keyset.json -csv testdata\web-samples\reports\perf-exp-xls-pipe-escaped-guard-keyset.csv ...retained keyset...`

Observed results:
- focused regression suite:
  - `ok officeread 12.475s`
- full repository regression:
  - `ok officeread 164.626s`
- hotspot benchmarks:
  - `006087.xls`: `5573310400 ns/op -> 4984286500 ns/op`
  - `008055.xls`: `7905695900 ns/op -> 5697742600 ns/op`
  - `016161.xls`: `3466805000 ns/op -> 2939523800 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - baseline median total: `23111 ms`
  - experiment median total: `19143 ms`
- broader repeat-aware keyset:
  - total: `51517 ms -> 47982 ms`
  - `.docx` subset: `6423 ms -> 6501 ms`
  - `.xls` subset: `17264 ms -> 17885 ms`
  - `.xlsx` subset: `27830 ms -> 23596 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Retained. This is the first recent `.xls` backfill optimization that improved the focused hotspot
  batch and also beat the retained repeat-aware wider keyset without changing output. The `.xls`
  subset moved slightly in the wrong direction on the wider keyset, but the overall retained workload
  improved materially and the change is tightly scoped to the legacy Excel markdown backfill path.

## 2026-07-05 rejected: dedicated `.xls` no-image-alt backfill implementation

- Tried following the retained `.xls` pipe-guard change with an even narrower specialization:
  - split `missingMarkdownTextXLS(...)` onto its own implementation
  - hardcode the “no images / no image alts” case for legacy Excel workbook backfill
  - remove `markdownImageAltSet(nil)` construction and the related image-alt exclusion checks from
    that `.xls`-only path
- This was designed to preserve the retained `.xls`-only escaped-table fallback guard while shaving
  a little more overhead from the same hot loop.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-xls-noimagealt-path-xls6.json -csv testdata\web-samples\reports\perf-exp-xls-noimagealt-path-xls6.csv ...`

Observed results:
- focused regression suite:
  - `ok officeread 12.800s`
- hotspot benchmarks:
  - `006087.xls`: `4984286500 ns/op -> 5168345700 ns/op`
  - `008055.xls`: `5697742600 ns/op -> 7785369900 ns/op`
  - `016161.xls`: `2939523800 ns/op -> 3448675000 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - retained baseline median total: `19143 ms`
  - experiment median total: `24676 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted immediately. Even this tightly scoped `.xls` specialization got slower on the focused
  hotspot set, so it did not earn a broader repeat-aware keyset pass.

## 2026-07-05 rejected: `visibleJoined` bigram prefilter for visible-line containment

- Tried adding a semantics-preserving precheck in `visibleLineContainsLine(...)`:
  - build a 2-byte bigram mask for `visibleJoined`
  - reject candidate lines early when one of their adjacent byte pairs never appears in the joined
    visible markdown text
  - keep the original `strings.Contains(c.visibleJoined, line)` path unchanged for candidates that
    pass the filter
- The intent was to cut down the expensive negative substring searches that still dominate the
  retained `.xls` hotspot profile, especially on `006087.xls`.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - used the baseline input list from `testdata\web-samples\reports\perf-repeat3-current-xls6-rerun.json`
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-visible-bigram-xls6.json -csv testdata\web-samples\reports\perf-exp-visible-bigram-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - used the baseline input list from `testdata\web-samples\reports\perf-repeat2-baseline-rune-count-keyset.json`
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-visible-bigram-keyset.json -csv testdata\web-samples\reports\perf-exp-visible-bigram-keyset.csv ...`

Observed results:
- full repository regression:
  - `ok officeread 150.392s`
- hotspot benchmarks:
  - `006087.xls`: `4984286500 ns/op -> 5092775400 ns/op`
  - `008055.xls`: `5697742600 ns/op -> 5703741100 ns/op`
  - `016161.xls`: `2939523800 ns/op -> 2551241800 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - retained baseline median total: `19143 ms`
  - experiment median total: `21792 ms`
- broader repeat-aware keyset:
  - total: `47982 ms -> 57913 ms`
  - `.docx` subset: `6501 ms -> 8096 ms`
  - `.xls` subset: `17885 ms -> 22350 ms`
  - `.xlsx` subset: `23596 ms -> 27467 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. The prefilter helped one hotspot sample materially, but it regressed the retained
  `.xls6` batch and the broader repeat-aware keyset too heavily to keep.

## 2026-07-05 rejected: disable visible-line substring containment in `.xls` backfill

- Tried a more direct `.xls`-only specialization after confirming that the hottest retained legacy
  Excel samples were not getting any actual wins from `visibleLineContainsLine(...)`:
  - added a `.xls`-only option that skips `visibleLineContains(variant)` inside
    `missingMarkdownTextWithOptions(...)`
  - kept the existing coverage set, table containment, exact visible-line hits, and the retained
    escaped table-cell pipe guard unchanged
- Supporting local analysis on the current retained hotspot outputs initially suggested no useful
  visible-substring wins on the hottest legacy Excel samples. A later, broader shape breakdown
  refined that picture:
  - `008055.xls`: no observed `visibleLineContainsLine(...)` dependence across the analyzed buckets
  - `006087.xls`: a small dependent slice remained, concentrated in short numeric lines whose
    apparent coverage came from substring matches against longer numeric cells
- That made the idea attractive for the focused hotspot slice, because the hottest remaining cost
  was dominated by negative substring scans through `visibleJoined`.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - used the baseline input list from `testdata\web-samples\reports\perf-repeat3-current-xls6-rerun.json`
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-xls-skip-visiblecontains-xls6.json -csv testdata\web-samples\reports\perf-exp-xls-skip-visiblecontains-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - used the baseline input list from `testdata\web-samples\reports\perf-repeat2-baseline-rune-count-keyset.json`
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-xls-skip-visiblecontains-keyset.json -csv testdata\web-samples\reports\perf-exp-xls-skip-visiblecontains-keyset.csv ...`

Observed results:
- focused regression suite:
  - `ok officeread 54.625s`
- full repository regression:
  - `ok officeread 200.171s`
- hotspot benchmarks:
  - `006087.xls`: `4984286500 ns/op -> 1888721600 ns/op`
  - `008055.xls`: `5697742600 ns/op -> 5977595100 ns/op`
  - `016161.xls`: `2939523800 ns/op -> 2482314400 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - retained baseline median total: `19143 ms`
  - experiment median total: `17221 ms`
- broader repeat-aware keyset:
  - total: `47982 ms -> 53164 ms`
  - `.docx` subset: `6501 ms -> 6664 ms`
  - `.xls` subset: `17885 ms -> 20558 ms`
  - `.xlsx` subset: `23596 ms -> 25942 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. Even though this `.xls`-only removal produced the strongest recent hotspot-batch gain,
  it still lost badly on the retained wider repeat-aware keyset. The current evidence says the
  expensive visible substring path is often useless on the two hottest `.xls` samples, but not
  useless enough across the wider retained workload to remove wholesale.

## 2026-07-05 rejected: `.xls` short-digit visible-substring guard

- Tried turning the broader `.xls` visible-substring removal into a more targeted rule:
  - keep the retained `.xls` pipe-guard path
  - only skip `visibleLineContains(variant)` for normalized `.xls` lines that are all ASCII digits
    and 4-7 bytes long
- This was motivated by the refined shape analysis on the retained wider `.xls` keyset:
  - `008055.xls`, `013623.xls`, `016161.xls`, `018548.xls`, `019088.xls`, `019089.xls`, and
    `002505.xls` showed no observed dependence on `visibleLineContainsLine(...)`
  - the remaining `006087.xls` dependent slice was tiny and its examples were all short numeric
    substrings such as `61047`, `30357`, `21595`, and `123311`
  - those values appeared as substrings inside longer visible numeric cells like `610479`,
    `303573`, `215950`, and `1233116`, so the branch looked more like a false-positive coverage
    path than a meaningful exact visible hit

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go extract_test.go`
- focused markdown / legacy regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - used the baseline input list from `testdata\web-samples\reports\perf-repeat3-current-xls6-rerun.json`
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-xls-short-digit-guard-xls6.json -csv testdata\web-samples\reports\perf-exp-xls-short-digit-guard-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - used the baseline input list from `testdata\web-samples\reports\perf-repeat2-baseline-rune-count-keyset.json`
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-xls-short-digit-guard-keyset.json -csv testdata\web-samples\reports\perf-exp-xls-short-digit-guard-keyset.csv ...`

Observed results:
- focused regression suite:
  - `ok officeread 46.377s`
- full repository regression:
  - `ok officeread 198.504s`
- hotspot benchmarks:
  - `006087.xls`: `4984286500 ns/op -> 2067601000 ns/op`
  - `008055.xls`: `5697742600 ns/op -> 8132269000 ns/op`
  - `016161.xls`: `2939523800 ns/op -> 3584242600 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - retained baseline median total: `19100 ms` on the same-day rerun baseline
  - experiment median total: `22356 ms`
- broader repeat-aware keyset:
  - same-day rerun baseline total: `51865 ms`
  - experiment total: `66333 ms`
  - same-day rerun baseline `.docx` subset: `6407 ms`
  - experiment `.docx` subset: `9274 ms`
  - same-day rerun baseline `.xls` subset: `17801 ms`
  - experiment `.xls` subset: `25407 ms`
  - same-day rerun baseline `.xlsx` subset: `27657 ms`
  - experiment `.xlsx` subset: `31652 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. The idea improved the one hotspot shape it was designed for, but in actual repeat-aware
  reruns it overfit badly and became one of the slowest recent `.xls` backfill branches.

## 2026-07-05 rejected: `.xls` short-digit exact-set plus visible-substring guard

- Tried a more balanced follow-up after the plain short-digit guard failed:
  - add a tiny `.xls`-only exact set for visible table cells that are all ASCII digits and 4-7
    bytes long
  - consult that exact set before the existing `< 12 runes` early return in
    `tableTextContainsLine(...)`
  - only if the short numeric line is not an exact table-cell hit, skip `visibleLineContains(...)`
    on the `.xls` path
- The goal was to keep true short numeric cell coverage while eliminating the false-positive
  substring path that showed up on `006087.xls`, where values like `61047` matched longer visible
  cells like `610479`.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go extract_test.go`
- targeted correctness checks:
  - exact short numeric table cell remains covered
  - short numeric substring of a longer visible cell is no longer treated as covered
- focused markdown / legacy regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-xls-short-digit-exact-guard-xls6.json -csv testdata\web-samples\reports\perf-exp-xls-short-digit-exact-guard-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-xls-short-digit-exact-guard-keyset.json -csv testdata\web-samples\reports\perf-exp-xls-short-digit-exact-guard-keyset.csv ...`

Observed results:
- focused regression suite:
  - `ok officeread 13.687s`
- full repository regression:
  - `ok officeread 172.405s`
- hotspot benchmarks:
  - `006087.xls`: `4984286500 ns/op -> 2190423400 ns/op`
  - `008055.xls`: `5697742600 ns/op -> 8050013700 ns/op`
  - `016161.xls`: `2939523800 ns/op -> 3451519400 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - retained baseline median total: `19100 ms` on the same-day rerun baseline
  - experiment median total: `22821 ms`
- broader repeat-aware keyset:
  - same-day rerun baseline total: `51865 ms`
  - experiment total: `65921 ms`
  - same-day rerun baseline `.docx` subset: `6407 ms`
  - experiment `.docx` subset: `8636 ms`
  - same-day rerun baseline `.xls` subset: `17801 ms`
  - experiment `.xls` subset: `25487 ms`
  - same-day rerun baseline `.xlsx` subset: `27657 ms`
  - experiment `.xlsx` subset: `31798 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. Even with an exact-hit rescue path for true short numeric cells, the added work still
  overfit the `006087.xls` hotspot and regressed the retained repeat-aware workload badly.

## 2026-07-05 rejected: disable `.xls` backfill query-result caches

- Profile-driven hypothesis:
  - on the hottest retained `.xls` samples, the per-query memoization maps inside
    `missingMarkdownTextWithOptions(...)`
    - `coverageCache`
    - `tableContainsCache`
    - `visibleContainsCache`
    were not obviously earning their keep
- Added a temporary local analysis pass over `006087.xls`, `008055.xls`, and `016161.xls` that
  replayed the current `.xls` backfill logic and counted query-cache hits. Observed hit rates:
  - `006087.xls`: coverage `0 / 41339`, table `0 / 404`, visible `0 / 404`
  - `008055.xls`: coverage `0 / 1181169`, table `0 / 0`, visible `0 / 0`
  - `016161.xls`: coverage `0 / 461603`, table `0 / 0`, visible `0 / 0`
- Based on that evidence, tried keeping the generic path unchanged while disabling those three
  query-result caches only for `missingMarkdownTextXLS(...)`.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-xls-no-query-caches-xls6.json -csv testdata\web-samples\reports\perf-exp-xls-no-query-caches-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-xls-no-query-caches-keyset.json -csv testdata\web-samples\reports\perf-exp-xls-no-query-caches-keyset.csv ...`
- same-day rerun baseline for comparison:
  - `testdata\web-samples\reports\perf-baseline-rerun-now-xls6.json`
  - `testdata\web-samples\reports\perf-baseline-rerun-now-keyset.json`

Observed results:
- full repository regression:
  - `ok officeread 191.013s`
- hotspot benchmarks:
  - `006087.xls`: `4984286500 ns/op -> 5172676100 ns/op`
  - `008055.xls`: `5697742600 ns/op -> 8398555800 ns/op`
  - `016161.xls`: `2939523800 ns/op -> 2790799400 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - same-day rerun baseline total: `19100 ms`
  - experiment total: `19576 ms`
- broader repeat-aware keyset:
  - same-day rerun baseline total: `51865 ms`
  - experiment total: `51122 ms`
  - same-day rerun baseline `.docx` subset: `6407 ms`
  - experiment `.docx` subset: `8266 ms`
  - same-day rerun baseline `.xls` subset: `17801 ms`
  - experiment `.xls` subset: `18153 ms`
  - same-day rerun baseline `.xlsx` subset: `27657 ms`
  - experiment `.xlsx` subset: `24703 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. The tempting whole-keyset total improvement was noise from unrelated `.docx/.xlsx`
  rerun variance, because this change only touched the `.xls` path. The directly affected `.xls`
  subset regressed on both the same-day keyset rerun and the same-day `.xls6` hotspot rerun, which
  means the zero observed cache-hit rate was still not a good reason to remove the caches.

## 2026-07-05 rejected: lazy containment construction for `.xls` backfill

- Shifted attention from per-query tweaks to the fixed upfront cost of
  `markdownBackfillContainmentSet(markdown)`.
- A temporary local replay over the retained `.xls` hotspot outputs showed:
  - `006087.xls`: `coverageOnly=40935`, `usedTable=0`, `usedVisible=404`
  - `008055.xls`: `coverageOnly=1181169`, `usedTable=0`, `usedVisible=0`
  - `016161.xls`: `coverageOnly=461603`, `usedTable=0`, `usedVisible=0`
  - `002505.xls`: `coverageOnly=214999`, `usedTable=0`, `usedVisible=0`
- That suggested many heavy `.xls` runs might pay the containment-build cost even when all useful
  coverage comes from `markdownBackfillCoverageSet(...)`.
- Tried keeping the generic path unchanged while making `missingMarkdownTextXLS(...)` build
  `markdownBackfillContainmentSet(markdown)` lazily on first use from `tableTextContains(...)` or
  `visibleLineContains(...)`.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-xls-lazy-containment-xls6.json -csv testdata\web-samples\reports\perf-exp-xls-lazy-containment-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-xls-lazy-containment-keyset.json -csv testdata\web-samples\reports\perf-exp-xls-lazy-containment-keyset.csv ...`

Observed results:
- full repository regression:
  - `ok officeread 183.801s`
- hotspot benchmarks:
  - `006087.xls`: `4984286500 ns/op -> 5184848800 ns/op`
  - `008055.xls`: `5697742600 ns/op -> 8462841300 ns/op`
  - `016161.xls`: `2939523800 ns/op -> 3564291300 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - retained baseline median total: `19100 ms` on the same-day rerun baseline
  - experiment median total: `26977 ms`
- broader repeat-aware keyset:
  - same-day rerun baseline total: `51865 ms`
  - experiment total: `66340 ms`
  - same-day rerun baseline `.docx` subset: `6407 ms`
  - experiment `.docx` subset: `9309 ms`
  - same-day rerun baseline `.xls` subset: `17801 ms`
  - experiment `.xls` subset: `29299 ms`
  - same-day rerun baseline `.xlsx` subset: `27657 ms`
  - experiment `.xlsx` subset: `27732 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. The containment build looked avoidable on paper, but deferring it made the `.xls`
  workload markedly slower. Whatever fixed cost it carries is still cheaper than the added lazy-path
  branching and late construction on the retained hotspot set.

## 2026-07-05 rejected: `.xls` exact-coverage fast path before variant assembly

- Added a temporary local replay to break down how `markdownBackfillCoverageCoversLine(...)` was
  actually succeeding on retained `.xls` hotspots. Observed hit kinds:
  - `006087.xls`: `exact=40935`, `visible=0`, `comparable=0`, `alternate=0`, `misses=404`
  - `008055.xls`: `exact=1181169`, `visible=0`, `comparable=0`, `alternate=0`, `misses=0`
  - `016161.xls`: `exact=461603`, `visible=0`, `comparable=0`, `alternate=0`, `misses=0`
  - `002505.xls`: `exact=214999`, `visible=0`, `comparable=0`, `alternate=0`, `misses=0`
- Based on that evidence, tried a very narrow `.xls`-only fast path:
  - after `lineMissingCache` / `seenMissing` checks, but before building `variants`
  - if `coverage[line]` already has a direct exact hit, mark the line covered immediately and skip
    all later variant construction and containment checks
- The intended win was to let the dominant exact-hit case in legacy Excel bail out before paying for
  `markdownBackfillVisibleText(...)`, `markdownVisibleLineText(...)`, duplicate checks, and the
  later expensive containment predicates.

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-xls-coverage-exact-fastpath-xls6.json -csv testdata\web-samples\reports\perf-exp-xls-coverage-exact-fastpath-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-xls-coverage-exact-fastpath-keyset.json -csv testdata\web-samples\reports\perf-exp-xls-coverage-exact-fastpath-keyset.csv ...`

Observed results:
- full repository regression:
  - `ok officeread 175.404s`
- hotspot benchmarks:
  - `006087.xls`: `4984286500 ns/op -> 5161550900 ns/op`
  - `008055.xls`: `5697742600 ns/op -> 8229030100 ns/op`
  - `016161.xls`: `2939523800 ns/op -> 3513155100 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - retained baseline median total: `19100 ms` on the same-day rerun baseline
  - experiment median total: `24716 ms`
- broader repeat-aware keyset:
  - same-day rerun baseline total: `51865 ms`
  - experiment total: `64938 ms`
  - same-day rerun baseline `.docx` subset: `6407 ms`
  - experiment `.docx` subset: `8693 ms`
  - same-day rerun baseline `.xls` subset: `17801 ms`
  - experiment `.xls` subset: `24307 ms`
  - same-day rerun baseline `.xlsx` subset: `27657 ms`
  - experiment `.xlsx` subset: `31938 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. Even though the hit-distribution evidence looked almost tailor-made for this fast path,
  the extra branch and duplicated exact-map check still lost badly on the retained repeat-aware
  workloads.

## 2026-07-05 rejected: lazy direct `coverage[line]` exact-hit fast path for `.xls`

- Built on a fresh coverage-hit breakdown for the retained `.xls` hotspots:
  - `006087.xls`: `exact=40935`, `visible=0`, `comparable=0`, `alternate=0`, `misses=404`
  - `008055.xls`: `exact=1181169`, `visible=0`, `comparable=0`, `alternate=0`, `misses=0`
  - `016161.xls`: `exact=461603`, `visible=0`, `comparable=0`, `alternate=0`, `misses=0`
  - `002505.xls`: `exact=214999`, `visible=0`, `comparable=0`, `alternate=0`, `misses=0`
- That made a very narrow `.xls`-only early-return experiment look plausible:
  - after the `lineMissingCache` / `seenMissing` checks
  - before building the `variants` slice or calling any of the visible/comparable helpers
  - if `coverage[line]` already had a direct exact hit, mark the line covered immediately and skip
    the rest of the per-line variant work

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-xls-coverage-exact-fastpath-xls6.json -csv testdata\web-samples\reports\perf-exp-xls-coverage-exact-fastpath-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-xls-coverage-exact-fastpath-keyset.json -csv testdata\web-samples\reports\perf-exp-xls-coverage-exact-fastpath-keyset.csv ...`

Observed results:
- full repository regression:
  - `ok officeread 175.404s`
- hotspot benchmarks:
  - `006087.xls`: `4984286500 ns/op -> 5161550900 ns/op`
  - `008055.xls`: `5697742600 ns/op -> 8229030100 ns/op`
  - `016161.xls`: `2939523800 ns/op -> 3513155100 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - retained baseline median total: `19100 ms` on the same-day rerun baseline
  - experiment median total: `24716 ms`
- broader repeat-aware keyset:
  - same-day rerun baseline total: `51865 ms`
  - experiment total: `64938 ms`
  - same-day rerun baseline `.docx` subset: `6407 ms`
  - experiment `.docx` subset: `8693 ms`
  - same-day rerun baseline `.xls` subset: `17801 ms`
  - experiment `.xls` subset: `24307 ms`
  - same-day rerun baseline `.xlsx` subset: `27657 ms`
  - experiment `.xlsx` subset: `31938 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. Even with near-perfect-looking exact-hit statistics, the added early branch still made
  the retained `.xls` workloads materially slower.

## 2026-07-05 rejected: direct BIFF-line backfill path for `.xls`

- Tried replacing the retained `.xls` backfill input path
  - instead of `strings.Join(biffText(data), "\n")` followed by the generic text-to-lines backfill
    preprocessing
  - introduced a dedicated `.xls` helper that consumed the `biffText(data)` line slice directly
    before backfill matching
- Goal:
  - remove the join/split round-trip in the legacy Excel markdown-backfill path without changing
    output

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-xls-direct-lines-xls6.json -csv testdata\web-samples\reports\perf-exp-xls-direct-lines-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-xls-direct-lines-keyset.json -csv testdata\web-samples\reports\perf-exp-xls-direct-lines-keyset.csv ...`

Observed results:
- full repository regression:
  - `ok officeread 139.279s`
- hotspot benchmarks:
  - `006087.xls`: `4984286500 ns/op -> 5405430900 ns/op`
  - `008055.xls`: `5697742600 ns/op -> 8471192500 ns/op`
  - `016161.xls`: `2939523800 ns/op -> 3449793300 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - same-day rerun baseline median total: `19100 ms`
  - experiment median total: `22473 ms`
- broader repeat-aware keyset:
  - same-day rerun baseline `.xls` subset: `17801 ms`
  - experiment `.xls` subset: `20832 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged

Decision:
- Reverted. The direct-line specialization removed a small conversion cost, but the retained
  baseline still won once measured on the repeat-aware hotspot and keyset reruns.

## 2026-07-05 retained: shared workbook-text reuse for `.xls`

- Problem:
  - normal legacy Excel extraction was rescanning the same workbook bytes multiple times in one
    `Extract(...)` call:
    - first for `Text` generation through `extractLegacyTextWithMetadata(...)`
    - then again for `StructuredMarkdown` generation through `biffMarkdown(...)`
    - and then a further `biffText(...)` call inside markdown backfill assembly
- Change:
  - in `extractLegacyWithDepth(...)`, for the normal `.xls` path with readable workbook bytes,
    compute workbook BIFF text once
  - reuse that text to build the `.xls` `Text` result
  - add `biffMarkdownWithText(...)` so markdown generation can reuse the same workbook text slice
    instead of rescanning workbook BIFF text again before backfill
- Scope:
  - only affects the normal `.xls` extraction path
  - `.doc`, `.ppt`, and OOXML formats remain on their existing logic
  - if the workbook text is empty, the old fallback behavior still applies

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Legacy|Test.*Markdown' -count=1 .`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-xls-shared-workbook-text-xls6.json -csv testdata\web-samples\reports\perf-exp-xls-shared-workbook-text-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-xls-shared-workbook-text-keyset.json -csv testdata\web-samples\reports\perf-exp-xls-shared-workbook-text-keyset.csv ...`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused regression suite:
  - `ok officeread 12.319s`
- hotspot benchmarks:
  - `006087.xls`: `4984286500 ns/op -> 5062447200 ns/op`
  - `008055.xls`: `5697742600 ns/op -> 6485531800 ns/op`
  - `016161.xls`: `2939523800 ns/op -> 2645504200 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - same-day rerun baseline total: `19100 ms`
  - experiment total: `16327 ms`
  - per-file medians:
    - `002505.xls`: `2268 ms -> 1746 ms`
    - `006087.xls`: `4946 ms -> 4670 ms`
    - `008055.xls`: `5734 ms -> 5080 ms`
    - `016161.xls`: `2451 ms -> 2106 ms`
    - `019088.xls`: `1737 ms -> 1386 ms`
    - `019089.xls`: `1964 ms -> 1339 ms`
- broader repeat-aware keyset:
  - same-day rerun baseline total: `51865 ms`
  - experiment total: `49269 ms`
  - same-day rerun baseline `.xls` subset: `17801 ms`
  - experiment `.xls` subset: `15657 ms`
  - same-day rerun baseline `.docx` subset: `6407 ms`
  - experiment `.docx` subset: `6360 ms`
  - same-day rerun baseline `.xlsx` subset: `27657 ms`
  - experiment `.xlsx` subset: `27252 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged
- full repository regression:
  - `ok officeread 160.342s`

Decision:
- Retained. This change removes repeated workbook scanning upstream of markdown-backfill matching,
  and it is the first recent `.xls`-focused optimization in this area that improves both the
  repeat-aware hotspot batch and the wider mixed keyset without changing output.

## 2026-07-05 retained: exact-first, shared-rest markdown-backfill construction

- Problem:
  - `missingMarkdownTextWithOptions(...)` was still repeating a substantial amount of markdown
    preprocessing before it ever reached the dominant `visibleLineContainsLine(...)` hotspot:
    - `markdownBackfillExactLinesCovered(...)` built an exact set from one markdown scan
    - `markdownBackfillCoverageSet(...)` built coverage data from another markdown scan
    - `markdownBackfillContainmentSet(...)` built containment haystacks from yet another markdown
      scan
  - on the current retained `006087.xls` profile before this change, those precompute stages still
    accounted for about:
    - `markdownBackfillExactLinesCovered(...)`: `0.29s`
    - `markdownBackfillCoverageSet(...)`: `0.44s`
    - `markdownBackfillContainmentSet(...)`: `0.14s`
- Change:
  - keep the exact-stage early-return structure intact:
    - build `exact` first
    - run `markdownBackfillExactLinesCoveredWithExact(...)`
  - only if that early check fails, build the shared `coverage + containment` structures together
    through `markdownBackfillBuildCoverageContainment(...)`
  - this supersedes the earlier retained all-in-one `markdownBackfillBuildSets(...)` variant
- Scope:
  - applies to the generic markdown-backfill path, so it can affect `.xls`, `.xlsx`, and `.docx`
    workloads that rely on the same additional-text / workbook-text logic
  - matching rules and output semantics remain unchanged

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .'`
- hotspot benchmark rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchtime=1x ./...`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-xls-shared-markdown-sets-xls6.json -csv testdata\web-samples\reports\perf-exp-xls-shared-markdown-sets-xls6.csv ...`
- broader repeat-aware keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-xls-shared-markdown-sets-keyset.json -csv testdata\web-samples\reports\perf-exp-xls-shared-markdown-sets-keyset.csv ...`
- focused profile refresh:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-after-shared-markdown-sets.pprof testdata\web-samples\samples\xls\006087.xls`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused regression suite:
  - `ok officeread 10.896s`
- hotspot benchmarks:
  - `006087.xls`: `5062447200 ns/op -> 4817586900 ns/op`
  - `008055.xls`: `6485531800 ns/op -> 4644516300 ns/op`
  - `016161.xls`: `2645504200 ns/op -> 1938585400 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - same-day rerun baseline total: `19100 ms`
  - experiment total: `15740 ms`
  - per-file medians:
    - `002505.xls`: `2268 ms -> 1750 ms`
    - `006087.xls`: `4946 ms -> 4710 ms`
    - `008055.xls`: `5734 ms -> 4625 ms`
    - `016161.xls`: `2451 ms -> 1992 ms`
    - `019088.xls`: `1737 ms -> 1349 ms`
    - `019089.xls`: `1964 ms -> 1314 ms`
- broader repeat-aware keyset:
  - same-day rerun baseline total: `51865 ms`
  - experiment total: `46830 ms`
  - same-day rerun baseline `.xls` subset: `17801 ms`
  - experiment `.xls` subset: `14754 ms`
  - same-day rerun baseline `.docx` subset: `6407 ms`
  - experiment `.docx` subset: `6608 ms`
  - same-day rerun baseline `.xlsx` subset: `27657 ms`
  - experiment `.xlsx` subset: `25468 ms`
- focused profile comparison on `006087.xls`:
  - previous shared-build retained variant: `4.70s`
  - exact-first / shared-rest variant: `4.94s`
  - precompute split after the change:
    - `markdownBackfillExactSet(...)`: about `0.24s`
    - `markdownBackfillBuildCoverageContainment(...)`: about `0.63s`
  - dominant remaining hotspot still:
    - `markdownBackfillContainment.visibleLineContainsLine(...)`: about `3.10s -> 3.13s`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: unchanged
- full repository regression:
  - `ok officeread 137.914s`

Decision:
- Retained. The exact-first / shared-rest split is not a win on the single `006087.xls` profile, but
  it keeps the useful early exact check and is materially better on the repeat-aware `.xls6`
  hotspot batch and the broader mixed keyset, which makes it the stronger retained baseline.

## 2026-07-05 retained: table exact hits before the short-line table cutoff

- Profile-driven hypothesis:
  - after the retained exact-first / shared-rest change, the dominant remaining `.xls` hotspot was
    still `markdownBackfillContainment.visibleLineContainsLine(...)`
  - fresh local diagnosis on `006087.xls` showed the expensive visible-line branch was being used
    only by short `4-6` byte queries, and every observed hit in that slice was digit-only
  - the containment builder already materialized `tableRawExact` and `tableVisibleExact`, but the
    old `tableTextContainsLine(...)` order returned early for `< 12` rune queries before consulting
    those exact maps
- Change:
  - move the existing `tableRawExact` / `tableVisibleExact` checks ahead of the old short-line
    early return in `markdownBackfillContainment.tableTextContainsLine(...)`
  - leave the substring and comparable-text paths unchanged
- Scope:
  - no matching path was removed
  - only short exact table hits are allowed to resolve earlier
  - longer-line behavior and the visible-line fallback remain intact

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- isolated `.xls` markdown-backfill benchmark:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkXLSMarkdownBackfillHotspots -benchmem -benchtime=1x ./`
- whole extract hotspot benchmark:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchmem -benchtime=1x ./`
- repeat-aware `.xls6` rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-short-exact-before-runecheck-xls6.json -csv testdata\web-samples\reports\perf-exp-short-exact-before-runecheck-xls6.csv ...`
- repeat-aware mixed keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-short-exact-before-runecheck-keyset.json -csv testdata\web-samples\reports\perf-exp-short-exact-before-runecheck-keyset.csv ...`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused regression suite:
  - `ok officeread 16.610s`
- isolated `.xls` markdown-backfill benchmark:
  - `006087.xls`: `1508848200 ns/op -> 1085084500 ns/op`
  - `008055.xls`: `3269654600 ns/op -> 2093789500 ns/op`
  - `016161.xls`: `1398286500 ns/op -> 1105105100 ns/op`
- whole extract hotspot benchmark:
  - `006087.xls`: `4963697200 ns/op -> 4695615800 ns/op`
  - `008055.xls`: `6463587000 ns/op -> 4719727300 ns/op`
  - `016161.xls`: `2898870500 ns/op -> 1879915300 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - same-day rerun baseline total: `19100 ms`
  - experiment total: `16482 ms`
  - per-file medians:
    - `002505.xls`: `2268 ms -> 1805 ms`
    - `006087.xls`: `4946 ms -> 4948 ms`
    - `008055.xls`: `5734 ms -> 4722 ms`
    - `016161.xls`: `2451 ms -> 2046 ms`
    - `019088.xls`: `1737 ms -> 1464 ms`
    - `019089.xls`: `1964 ms -> 1497 ms`
- repeat-aware mixed keyset:
  - same-day rerun baseline total: `51865 ms`
  - experiment total: `45205 ms`
  - same-day rerun baseline `.xls` subset: `17801 ms`
  - experiment `.xls` subset: `14646 ms`
  - same-day rerun baseline `.docx` subset: `6407 ms`
  - experiment `.docx` subset: `6829 ms`
  - same-day rerun baseline `.xlsx` subset: `27657 ms`
  - experiment `.xlsx` subset: `23730 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: `NO_DIFF` on both `.xls6` and mixed keyset
- full repository regression:
  - `ok officeread 140.389s`

Decision:
- Retained. The change is small and local, but it produces a real repeat-aware `.xls` improvement
  and lowers the broader mixed-keyset total while keeping outputs stable and the full repository
  regression clean.

## 2026-07-05 retained refinement: limit short table exact-before-cutoff to `.xls`

- Follow-up hypothesis:
  - the previous retained change was effective, but the mixed-keyset `.docx` subset moved from
    `6407 ms` to `6829 ms`
  - the evidence that motivated the change was `.xls`-specific: short numeric values resolved from
    table exact sets before falling into expensive visible-line substring checks
  - therefore the broader all-format scope was likely unnecessary
- Change:
  - keep `shortTableExactBeforeMinLen` enabled only for `missingMarkdownTextXLS(...)`
  - restore the previous generic behavior for `missingMarkdownText(...)`
- Scope:
  - `.xls` keeps the short exact-table fast path
  - `.docx` / `.xlsx` return to the prior generic path
  - no output rule is intentionally changed

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- repeat-aware `.xls6` rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-short-exact-xls-only-xls6.json -csv testdata\web-samples\reports\perf-exp-short-exact-xls-only-xls6.csv ...`
- repeat-aware mixed keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json -csv testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.csv ...`
- focused regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- repeat-aware `.xls6` hotspot batch:
  - same-day rerun baseline total: `19100 ms`
  - prior all-format version: `16482 ms`
  - `.xls`-only refinement: `16118 ms`
- repeat-aware mixed keyset:
  - same-day rerun baseline total: `51865 ms`
  - prior all-format version: `45205 ms`
  - `.xls`-only refinement: `44994 ms`
  - by subset:
    - `.docx`: `6407 ms -> 6829 ms -> 6433 ms`
    - `.xls`: `17801 ms -> 14646 ms -> 14471 ms`
    - `.xlsx`: `27657 ms -> 23730 ms -> 24090 ms`
  - per-file `.docx` medians:
    - `00003763.docx`: `3855 ms -> 3944 ms`
    - `223624.docx`: `2552 ms -> 2489 ms`
- output stability:
  - retained-vs-baseline `textBytes` / `images`: `NO_DIFF` on both `.xls6` and mixed keyset
- focused regression suite:
  - `ok officeread 11.220s`
- full repository regression:
  - `ok officeread 152.104s`

Decision:
- Retained, replacing the broader all-format variant.
- The narrower `.xls`-only scope preserves the intended `.xls` win, slightly improves the overall
  mixed-keyset total again, and removes almost all of the `.docx` regression from the broader
  version.

## 2026-07-05 rejected: one-pass tag scan for `simpleInlineWorksheetCandidate(...)`

- Profile-driven hypothesis:
  - with the current retained `.xls` baseline in place, the largest remaining `.xlsx` hotspot in
    the mixed keyset was again `testRecordSizeExceeded.xlsx`
  - fresh current profile on that sample:
    - `testRecordSizeExceeded.xlsx`: `4358 ms`
    - dominant path: `appendSimpleInlineWorksheetText(...)`
    - notable sub-cost: repeated `bytes.Contains(...)` / `bytes.Index(...)` work inside
      `simpleInlineWorksheetCandidate(...)`
  - this suggested a structural optimization candidate: scan the worksheet once by tags instead of
    repeatedly running whole-buffer substring checks for each disqualifier marker
- Change:
  - replaced the old repeated marker scans in `simpleInlineWorksheetCandidate(...)` with a single
    start-tag scan
  - kept the same intended disqualifier families: formulas, `<v>`, shared/bool/str cells, hidden
    markers, hyperlinks, data validations, phonetic markers, and header/footer tags
- Scope:
  - only touched the fast-path candidate gate for simple inlineStr worksheets
  - `appendSimpleInlineWorksheetText(...)` and the generic worksheet parser were otherwise unchanged

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- focused `.xlsx` pair replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-simple-inline-candidate-xlsx-pair.json -csv testdata\web-samples\reports\perf-exp-simple-inline-candidate-xlsx-pair.csv testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx`
- post-revert regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused `.xlsx` regression slice:
  - `ok officeread 8.482s`
- focused `.xlsx` pair replay against the current retained keyset baseline:
  - `00012389.xlsx`: `2460 ms -> 2735 ms`
  - `testRecordSizeExceeded.xlsx`: `3967 ms -> 4040 ms`
- output stability:
  - output remained functionally stable on the focused pair
- post-revert regression:
  - `ok officeread 5.336s`
  - full suite returned to the retained baseline state

Decision:
- Reverted. The profile shape made the repeated marker scans look suspicious, but the one-pass tag
  scanner still lost on the real retained `.xlsx` hotspots, so it is not a keeper.

## 2026-07-05 rejected: single-`<t>` fast path in `simpleInlineCellText(...)`

- Benchmark setup improvement:
  - added `.xlsx` hotspot benchmarks in `extract_bench_test.go` so the remaining OOXML worksheet
    cost could be measured more directly:
    - `BenchmarkExtractXLSXHotspots`
    - `BenchmarkXLSXSimpleInlineHotspots`
    - `BenchmarkXLSXWorksheetTextHotspots`
- Fresh baseline from the new benches:
  - `testRecordSizeExceeded.xlsx`
    - whole extract: `4640854300 ns/op`
    - simple inline fast path: `4190777000 ns/op`
    - worksheet text path: `4650526500 ns/op`
  - `00012389.xlsx`
    - whole extract: `2965216100 ns/op`
    - worksheet text path: `2066764500 ns/op`
- Profile-driven hypothesis:
  - on the current `testRecordSizeExceeded.xlsx` profile, `simpleInlineCellText(...)` still
    appeared as a visible per-cell cost inside `appendSimpleInlineWorksheetText(...)`
  - many simple inline cells looked likely to contain a single `<t>...</t>` payload, so a narrow
    direct-return path looked worth testing
- Change:
  - added a fast path to `simpleInlineCellText(...)`:
    - if the cell contains only one `<t>...</t>` segment, return it directly
    - fall back to the existing builder loop for multi-segment rich text
- Scope:
  - only touched simple inline text extraction for `.xlsx` inlineStr cells
  - no output rule was intentionally changed

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go extract_bench_test.go`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- `.xlsx` hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- focused `.xlsx` pair replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-simple-inline-celltext-xlsx-pair.json -csv testdata\web-samples\reports\perf-exp-simple-inline-celltext-xlsx-pair.csv testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx`
- mixed keyset replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-simple-inline-celltext-keyset.json -csv testdata\web-samples\reports\perf-exp-simple-inline-celltext-keyset.csv ...`
- post-revert regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused `.xlsx` regression slice:
  - `ok officeread 8.494s`
- benchmark view looked strong:
  - `testRecordSizeExceeded.xlsx`
    - whole extract: `4640854300 ns/op -> 4570901400 ns/op`
    - simple inline: `4190777000 ns/op -> 3958273900 ns/op`
    - worksheet text: `4650526500 ns/op -> 3879766600 ns/op`
  - `00012389.xlsx`
    - worksheet text: `2066764500 ns/op -> 1470050600 ns/op`
- output stability:
  - mixed keyset `textBytes` / `images`: `NO_DIFF`
- but repeat-aware validation clearly failed:
  - focused `.xlsx` pair against the current retained keyset baseline:
    - `00012389.xlsx`: `2460 ms -> 2763 ms`
    - `testRecordSizeExceeded.xlsx`: `3967 ms -> 4761 ms`
  - mixed keyset total:
    - current retained baseline: `44994 ms`
    - experiment: `51552 ms`
  - subset totals:
    - `.docx`: `6433 ms -> 7819 ms`
    - `.xls`: `14471 ms -> 16127 ms`
    - `.xlsx`: `24090 ms -> 27606 ms`
- post-revert regression:
  - focused `.xlsx` slice: `ok officeread 7.784s`
  - full suite: `ok officeread 134.471s`

Decision:
- Reverted. The microbenchmark win did not survive repeat-aware workload validation, so this is not
  a retained optimization despite the attractive narrow numbers.

## 2026-07-05 rejected: byte-based cell reference parsing in simple inline `.xlsx` path

- Follow-up hypothesis:
  - after adding the new `.xlsx` hotspot benchmarks, one low-risk-looking allocation target was the
    per-cell `string(cellRef)` conversion inside `appendSimpleInlineWorksheetText(...)`
  - the idea was to keep behavior the same while reducing tiny per-cell string churn in the simple
    inline `.xlsx` path
- Change:
  - added a byte-oriented `cellRefIndexesBytes(...)`
  - used it only for the inlineStr fast path in `appendSimpleInlineWorksheetText(...)`
- Scope:
  - no text extraction or cleaning rule was changed
  - the experiment only altered how the `r=` cell reference attribute was parsed in one `.xlsx`
    path

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- `.xlsx` hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- focused `.xlsx` pair replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-cellref-bytes-xlsx-pair.json -csv testdata\web-samples\reports\perf-exp-cellref-bytes-xlsx-pair.csv testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx`
- mixed keyset replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-cellref-bytes-keyset.json -csv testdata\web-samples\reports\perf-exp-cellref-bytes-keyset.csv ...`
- post-revert regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused `.xlsx` regression slice:
  - `ok officeread 7.869s`
- narrow signals were mixed:
  - focused `.xlsx` pair:
    - `00012389.xlsx`: `2460 ms -> 2436 ms`
    - `testRecordSizeExceeded.xlsx`: `3967 ms -> 4015 ms`
  - hotspot benchmarks:
    - `testRecordSizeExceeded.xlsx` whole extract: `4640854300 ns/op -> 4537746500 ns/op`
    - `00012389.xlsx` whole extract: `2965216100 ns/op -> 3293568900 ns/op`
- output stability:
  - mixed keyset `textBytes` / `images`: `NO_DIFF`
- wider validation clearly failed:
  - mixed keyset total:
    - current retained baseline: `44994 ms`
    - experiment: `56450 ms`
  - subset totals:
    - `.docx`: `6433 ms -> 8028 ms`
    - `.xls`: `14471 ms -> 17589 ms`
    - `.xlsx`: `24090 ms -> 30833 ms`
- post-revert regression:
  - focused `.xlsx` slice: `ok officeread 7.728s`
  - full suite returned to the retained baseline state

Decision:
- Reverted. This allocation-focused tweak was too small and too noisy to survive repeat-aware mixed
  workload validation.

## 2026-07-05 rejected: stop simple-inline row/col tracking after Markdown row cap

- Benchmark-driven observation:
  - the new `.xlsx` hotspot benchmarks made it possible to separate simple-inline text extraction
    from the Markdown side work
  - on `testRecordSizeExceeded.xlsx`:
    - `BenchmarkXLSXSimpleInlineHotspots`: `3764269900 ns/op`
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots`: `3526142900 ns/op`
  - so the Markdown collector inside the simple-inline path is nontrivial, but it is still only a
    slice of the full worksheet-text cost
- Hypothesis:
  - once `md.rows` has already reached `maxMarkdownTableRows`, continuing to parse `r=` cell refs
    and maintain row/column bookkeeping for the rest of the worksheet might be wasted work
- Change:
  - after the Markdown row cap was reached, stop parsing per-cell row/column references and stop
    building row state for the remainder of `appendSimpleInlineWorksheetText(...)`
  - text output continued unchanged
- Scope:
  - only touched the simple-inline `.xlsx` Markdown-side bookkeeping
  - no intended text or Markdown semantic change before the existing row cap

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- `.xlsx` hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots|BenchmarkXLSXWorksheetTextHotspots|BenchmarkExtractXLSXHotspots' -benchmem -benchtime=1x ./`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- focused `.xlsx` pair replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-inline-markdown-stop-after-cap-xlsx-pair.json -csv testdata\web-samples\reports\perf-exp-inline-markdown-stop-after-cap-xlsx-pair.csv testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx`
- post-revert regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`

Observed results:
- benchmark view was mixed and not trustworthy enough on its own:
  - `testRecordSizeExceeded.xlsx`
    - whole extract: `4640854300 ns/op -> 4576078000 ns/op`
    - simple inline: `3764269900 ns/op -> 4184935800 ns/op`
    - text-only simple inline: `3526142900 ns/op -> 3166360900 ns/op`
  - `00012389.xlsx`
    - worksheet text: `1915889800 ns/op -> 1594668300 ns/op`
- focused `.xlsx` regression slice:
  - `ok officeread 10.395s`
- focused `.xlsx` pair replay against the current retained baseline:
  - `00012389.xlsx`: `2460 ms -> 2574 ms`
  - `testRecordSizeExceeded.xlsx`: `3967 ms -> 4157 ms`
- post-revert regression:
  - `ok officeread 5.513s`

Decision:
- Reverted. Even a semantics-preserving-looking tail-work reduction did not survive real `.xlsx`
  repeat-aware validation on the retained hotspot pair.

## 2026-07-05 retained benchmark refinement: isolate prepared simple-inline `.xlsx` cost

- benchmark-only refactor:
  - split `appendSimpleInlineWorksheetText(...)` so the timed loop can call
    `appendSimpleInlineWorksheetTextPrepared(...)` directly after one upfront
    `simpleInlineWorksheetCandidate(...)` check
- reason:
  - the old benchmark shape was overstating the cost of the simple-inline body by charging the
    candidate gate to every timed iteration

Validation:
- `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go extract_bench_test.go`
- `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineTextOnlyHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchtime=1x -cpuprofile tmp-shape\xlsx-simpleinline-textonly-prepared.prof ./`
- `& 'C:\Program Files\Go\bin\go.exe' tool pprof -top .\officeread.test.exe tmp-shape\xlsx-simpleinline-textonly-prepared.prof`

Observed results:
- corrected hotspot benchmarks:
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2890754000 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `4761816100 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `2126471600 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1922540000 ns/op`
- corrected profile takeaway:
  - once the candidate gate is moved out of the timed loop, the real timed cost is the per-cell
    extraction / cleaning stack, not the worksheet candidate precheck

Decision:
- Retained. This is a measurement correction, not a runtime optimization, but it is important
  infrastructure for future `.xlsx` tuning.

## 2026-07-05 rejected: skip the second trim inside `appendWorksheetValue(...)`

- experiment:
  - after `cleanText(...)`, send worksheet values straight to `appendTrimmedTextBlock(...)`
    instead of `appendCleanedTextBlock(...)`
- motivation:
  - `cleanText(...)` already returns a trimmed string, so this looked like a tiny safe win in the
    hot worksheet path

Validation:
- `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-ai-assistant-worksheetvalue-xlsx-pair.json -csv testdata\web-samples\reports\perf-exp-ai-assistant-worksheetvalue-xlsx-pair.csv testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-ai-assistant-worksheetvalue-keyset.json -csv testdata\web-samples\reports\perf-exp-ai-assistant-worksheetvalue-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`

Observed results:
- focused `.xlsx` microbench signals improved:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `4827813500 ns/op -> 4779573600 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `3350374500 ns/op -> 2911278000 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2890754000 ns/op -> 2640816000 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `4761816100 ns/op -> 4543836300 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `2126471600 ns/op -> 2017237700 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1922540000 ns/op -> 1721084500 ns/op`
- repeat-aware `.xlsx` pair still regressed:
  - `00012389.xlsx`: `2460 ms -> 2535 ms`
  - `testRecordSizeExceeded.xlsx`: `3967 ms -> 5051 ms`
- repeat-aware mixed keyset also lost overall:
  - retained baseline total: `44994 ms`
  - experiment total: `45549 ms`
  - by subset:
    - `.docx`: `6433 ms -> 7994 ms`
    - `.xls`: `14471 ms -> 14020 ms`
    - `.xlsx`: `24090 ms -> 23535 ms`
- retained-vs-experiment `textBytes` / `images`: `NO_DIFF`

Decision:
- Reverted. Another reminder that a local worksheet fast path can still lose once the retained
  mixed workload and GC behavior are included.

## 2026-07-05 rejected: narrower gating for control-fragment helper checks

- experiment:
  - restructure `maybeControlFragmentText(...)` so the large first-character literal switch runs
    before the more expensive helper recognizers
  - only call helper checks like `isLegacyObjectReference(...)`,
    `looksLikeOLEIdentifierFragment(...)`, `looksLikeOLEWrapperStreamName(...)`, and
    `looksLikeOOXMLMarkupNameFragment(...)` when the string shape looked plausible
- motivation:
  - the corrected prepared simple-inline `.xlsx` profile still showed nontrivial time in
    `cleanTextFastPath -> maybeControlFragmentText`

Validation:
- `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-control-fragment-gating-xlsx-pair.json -csv testdata\web-samples\reports\perf-exp-control-fragment-gating-xlsx-pair.csv testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-control-fragment-gating-keyset.json -csv testdata\web-samples\reports\perf-exp-control-fragment-gating-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`
- `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused `.xlsx` microbench signals were dramatic:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `4827813500 ns/op -> 3867231400 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `3350374500 ns/op -> 2212112400 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2890754000 ns/op -> 1832484100 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `4761816100 ns/op -> 3450764400 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `2126471600 ns/op -> 1289478700 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1922540000 ns/op -> 1080451600 ns/op`
- repeat-aware `.xlsx` pair also improved:
  - `00012389.xlsx`: `2460 ms -> 2256 ms`
  - `testRecordSizeExceeded.xlsx`: `3967 ms -> 3539 ms`
- repeat-aware mixed keyset improved overall:
  - retained baseline total: `44994 ms`
  - experiment total: `43569 ms`
  - by subset:
    - `.docx`: `6433 ms -> 6362 ms`
    - `.xls`: `14471 ms -> 13009 ms`
    - `.xlsx`: `24090 ms -> 24198 ms`
- but output parity failed on the retained keyset:
  - `00010400.xlsx` changed from `textBytes=2778014` to `2778070`

Decision:
- Reverted. The speedup is real, but this experiment changed extraction output, so it does not meet
  the retention bar.

## 2026-07-05 rejected: safe subset of control-fragment helper gating

- follow-up idea:
  - keep `looksLikeOOXMLMarkupNameFragment(...)` as an early unconditional filter so `relationships`
    and similar OOXML names do not leak
  - still delay the other heavier helper recognizers behind narrower shape checks:
    `isLegacyObjectReference(...)`, `looksLikeOLEIdentifierFragment(...)`, and
    `looksLikeOLEWrapperStreamName(...)`
- goal:
  - preserve the good part of the previous speedup while avoiding the output drift on
    `00010400.xlsx`

Validation:
- `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-control-fragment-gating-safe-keyset.json -csv testdata\web-samples\reports\perf-exp-control-fragment-gating-safe-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-control-fragment-gating-safe-keyset-rerun2.json -csv testdata\web-samples\reports\perf-exp-control-fragment-gating-safe-keyset-rerun2.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-control-fragment-gating-safe-docx-pair.json -csv testdata\web-samples\reports\perf-exp-control-fragment-gating-safe-docx-pair.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx`
- `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- output parity was restored:
  - both mixed-keyset reruns reported retained-vs-experiment `textBytes` / `images`: `NO_DIFF`
- first repeat-aware mixed keyset rerun:
  - retained baseline total: `44994 ms`
  - experiment total: `44816 ms`
  - `.docx`: `6433 ms -> 8006 ms`
  - `.xls`: `14471 ms -> 13474 ms`
  - `.xlsx`: `24090 ms -> 23336 ms`
- second repeat-aware mixed keyset rerun:
  - retained baseline total: `44994 ms`
  - experiment total: `46322 ms`
  - `.docx`: `6433 ms -> 8657 ms`
  - `.xls`: `14471 ms -> 13395 ms`
  - `.xlsx`: `24090 ms -> 24270 ms`
- repeat-aware `.docx` pair replay:
  - `00003763.docx`: `3890 ms` median (`3885..4098`)
  - `223624.docx`: `2535 ms` median (`2529..2552`)
- `.xlsx` hotspot benchmarks remained appealing:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `4084598800 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2307492100 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2209375000 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1209657100 ns/op`
- full repository regression:
  - `ok officeread 131.632s`

Decision:
- Reverted. This safer subset removed the output drift, but the repeat-aware wide workload still did
  not produce a clear and stable win, so it does not beat the retained baseline.

## 2026-07-05 rejected: ASCII-style rewrite of OLE identifier helper internals

- experiment:
  - leave the outer control-fragment routing unchanged
  - only rewrite internals in the OLE identifier helper chain:
    - `oleIdentifierAssignmentValue(...)`
    - `looksLikeGUIDString(...)`
  - avoid some generic lowercase / rune-based paths in favor of narrower ASCII checks
- motivation:
  - fresh profiles still showed nontrivial cumulative time in
    `looksLikeOLEIdentifierFragment(...)` and `oleIdentifierAssignmentValue(...)`

Validation:
- `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-ole-identifier-helper-keyset.json -csv testdata\web-samples\reports\perf-exp-ole-identifier-helper-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`

Observed results:
- output parity stayed clean:
  - retained-vs-experiment `textBytes` / `images`: `NO_DIFF`
- but performance regressed on both the hotspot view and the retained mixed workload:
  - hotspot benchmarks:
    - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `4827813500 ns/op -> 4694417700 ns/op`
    - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `3350374500 ns/op -> 3521552600 ns/op`
    - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2890754000 ns/op -> 3117330200 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `4761816100 ns/op -> 4882147800 ns/op`
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1922540000 ns/op -> 1964246300 ns/op`
  - repeat-aware mixed keyset:
    - retained baseline total: `44994 ms`
    - experiment total: `51716 ms`
    - `.docx`: `6433 ms -> 8812 ms`
    - `.xls`: `14471 ms -> 18204 ms`
    - `.xlsx`: `24090 ms -> 24700 ms`

Decision:
- Reverted. This is another case where a seemingly tidy helper rewrite lost badly on the retained
  workload despite preserving outputs.

## 2026-07-05 rejected: ASCII shortcut for markdown-backfill rune-count cutoffs

- experiment:
  - keep all markdown-backfill matching behavior unchanged
  - only replace a few short-query `utf8.RuneCountInString(...) < N` checks with an equivalent
    helper that:
    - returns early when `len(s) < N`
    - skips rune counting for all-ASCII strings when `len(s) >= N`
    - falls back to `utf8.RuneCountInString(...)` only for non-ASCII strings
- touched functions:
  - `markdownBackfillContainment.tableTextContainsLine(...)`
  - `markdownBackfillContainment.visibleLineContainsLine(...)`
  - `markdownBackfillComparableContains(...)`
- motivation:
  - these cutoff checks are on the hot containment path, and the optimization looked fully
    semantics-preserving

Validation:
- `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots' -benchmem -benchtime=1x ./`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-rune-count-below-keyset.json -csv testdata\web-samples\reports\perf-exp-rune-count-below-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`

Observed results:
- output parity stayed clean:
  - retained-vs-experiment `textBytes` / `images`: `NO_DIFF`
- but the `.xls` hotspot benchmark regressed across the board:
  - `006087.xls`: `964043400 ns/op -> 1410474400 ns/op`
  - `008055.xls`: `2236016400 ns/op -> 2497370400 ns/op`
  - `016161.xls`: `889076200 ns/op -> 1056510600 ns/op`
- repeat-aware mixed keyset also regressed strongly:
  - retained baseline total: `44994 ms`
  - experiment total: `50170 ms`
  - `.docx`: `6433 ms -> 7886 ms`
  - `.xls`: `14471 ms -> 15407 ms`
  - `.xlsx`: `24090 ms -> 26877 ms`

Decision:
- Reverted. This equivalent-looking shortcut still lost badly on the retained workload, so it is not
  worth keeping.

## 2026-07-05 benchmark refinement: split `.xls` markdown-backfill build costs

- Added two new measurement-only hotspot benchmark families in
  [extract_bench_test.go](/D:/workprj/officeread/extract_bench_test.go):
  - `BenchmarkXLSMarkdownExactSetHotspots`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots`
- Command:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots|BenchmarkXLSMarkdownExactSetHotspots|BenchmarkXLSMarkdownCoverageContainmentHotspots' -benchmem -benchtime=1x ./`
- Current retained-baseline results:
  - whole backfill:
    - `006087.xls`: `1339661800 ns/op`, `331063944 B/op`, `3453712 allocs/op`
    - `008055.xls`: `2654517100 ns/op`, `573348400 B/op`, `1532733 allocs/op`
    - `016161.xls`: `1128520700 ns/op`, `261876776 B/op`, `968662 allocs/op`
  - exact-set build only:
    - `006087.xls`: `247864400 ns/op`, `49135424 B/op`, `590853 allocs/op`
    - `008055.xls`: `1124659900 ns/op`, `237955208 B/op`, `1515380 allocs/op`
    - `016161.xls`: `539833200 ns/op`, `129712200 B/op`, `962010 allocs/op`
  - coverage+containment build only:
    - `006087.xls`: `569707600 ns/op`, `260579304 B/op`, `2692516 allocs/op`
    - `008055.xls`: `1620559700 ns/op`, `780180304 B/op`, `2437498 allocs/op`
    - `016161.xls`: `997024500 ns/op`, `485550680 B/op`, `3025037 allocs/op`
- Takeaway:
  - `006087.xls` is still more containment-heavy
  - `008055.xls` remains expensive on both exact-set and containment construction
  - these two hotspot shapes should continue to be treated separately

## 2026-07-05 rejected: one-pass combined build for `.xls` exact set plus coverage/containment

- Idea:
  - make `missingMarkdownTextWithOptions(...)` build exact-set coverage and
    coverage/containment tables in one shared markdown pass
- Why:
  - after the new split benchmarks, removing a second full markdown scan looked like the most
    obvious structural experiment
- Result:
  - focused regression stayed green: `ok officeread 12.509s`
  - integrated backfill benchmark became unstable in the wrong direction:
    - `006087.xls`: `1120972800 ns/op -> 842505700 ns/op`
    - `008055.xls`: `2163713000 ns/op -> 4012425300 ns/op`
    - `016161.xls`: `1017316800 ns/op -> 2067587200 ns/op`
- Interpretation:
  - the fused build likely increased peak live-heap / GC cost by keeping both large structures alive
    together
- Decision:
  - Reverted.

## 2026-07-06 measured: exact-set/source-line variant usage is too narrow to justify another heading-focused optimization

- validation:
  - source-line variant probe:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSExactVariantShape$' -v ./`
  - exact-set markdown-shape probe:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSExactSetHashShape$' -v ./`

- source-line variant results:
  - `006087.xls`:
    - `total=41347`, `allowed=41339`, `visibleAdds=0`, `markdownAdds=0`, `escapedAdds=0`
  - `008055.xls`:
    - `total=1181198`, `allowed=1181169`, `visibleAdds=0`, `markdownAdds=2`, `escapedAdds=0`
    - examples:
      - `#ID_REF = -> ID_REF =`
      - `#VALUE = raw signal output from ArrayPro image analysis program -> VALUE = raw signal output from ArrayPro image analysis program`
  - `016161.xls`:
    - `total=461615`, `allowed=461603`, `visibleAdds=0`, `markdownAdds=2`, `escapedAdds=0`
    - examples:
      - `#1376-2 -> 1376-2`
      - `#581 -> 581`

- exact-set markdown-shape results:
  - `008055.xls`:
    - `total=16917`, `plain=0`, `nonPlain=16917`, `startsHash=4`, `startsPipe=16910`, `changedMarkdown=7`
    - heading examples:
      - `## Sheet1 -> Sheet1`
      - `## Sheet2 -> Sheet2`
      - `## Sheet3 -> Sheet3`
      - `## Workbook Sheets -> Workbook Sheets`
  - `016161.xls`:
    - `total=45634`, `plain=0`, `nonPlain=45634`, `startsHash=7`, `startsPipe=45621`, `changedMarkdown=13`
    - heading examples:
      - `## chr1 -> chr1`
      - `## chr1_siRNAs -> chr1_siRNAs`
      - `## chr2 -> chr2`
      - `## chr2_siRNAs -> chr2_siRNAs`
      - `## Workbook Sheets -> Workbook Sheets`

- interpretation:
  - the rare non-direct source-line additions are almost entirely leading-`#` cases
  - but the markdown fed into `markdownBackfillExactSet(...)` is still overwhelmingly table rows,
    not headings
  - this means a new heading-specific fast path would remove only a tiny fraction of the retained
    exact-stage work and is unlikely to beat the integrated baseline
  - the next promising direction remains reducing or restructuring table-row / cell exact-set work,
    not another `#...` normalization micro-optimization

- decision:
  - No code change from this measurement.

## 2026-07-06 measured: only `006087.xls` reaches the second stage, and its remaining queries are purely `visible` hits

- validation:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSQueryResolution$' -v ./`
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLS006087SecondStageTiming$' -v ./`

- results:
  - `006087.xls`:
    - `totalAllowed=41339`
    - `exact=40935`
    - `coverage=0`
    - `table=0`
    - `visible=404`
    - `miss=0`
    - second-stage build: `469 ms`
    - second-stage resolve: `20 ms`
  - `008055.xls`:
    - `totalAllowed=1181169`
    - `exact=1181169`
    - `coverage=0`
    - `table=0`
    - `visible=0`
    - `miss=0`
  - `016161.xls`:
    - `totalAllowed=461603`
    - `exact=461603`
    - `coverage=0`
    - `table=0`
    - `visible=0`
    - `miss=0`

- interpretation:
  - the retained heavy `.xls` samples are now sharply split:
    - `008055.xls` and `016161.xls` terminate in exact precheck
    - `006087.xls` is the only retained hotspot sample still paying second-stage cost
  - on `006087.xls`, the residual second-stage work is purely `visible` resolution, and the
    expensive part is building the full second-stage state
  - this pointed to a plausible next experiment: a very narrow `006087`-shape exact+visible-only
    fallback before the full second-stage build

- decision:
  - No code change from this measurement.

## 2026-07-06 corrected and retained: `006087`-shape exact+visible-only fallback before full second-stage build

- experiment:
  - after exact precheck fails on `.xls`, add a narrow preflight for the `006087`-style shape:
    - `len(lines) >= 30000`
    - `len(text) <= 512 KiB`
  - on that path:
    - keep the existing exact-set
    - build only visible containment
    - accept the result if every remaining source line is covered by `exact` or
      `visibleLineContainsLine(...)`
    - otherwise fall back to the retained full `coverage + table + visible` build

- rationale:
  - the visible-only precheck itself did cover `006087.xls` successfully in about `152 ms`
  - compared with the measured `469 ms` full second-stage build, that looked like a strong
    candidate for a narrow retained win

- validation:
  - focused hotspots:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/006087.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
  - repeat-aware `.xls6` rerun:
    - `testdata/web-samples/reports/perf-exp-xls-visible-only-fallback-xls6.json`
    - `testdata/web-samples/reports/perf-exp-xls-visible-only-fallback-xls6.csv`
  - repeat-aware mixed keyset rerun:
    - `testdata/web-samples/reports/perf-exp-xls-visible-only-fallback-keyset.json`
    - `testdata/web-samples/reports/perf-exp-xls-visible-only-fallback-keyset.csv`
  - repository regression before revert:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

- focused benchmark results:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `645141700 ns/op -> 167294500 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1625965900 ns/op -> 1665857100 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `638316500 ns/op -> 692025000 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `110346100 ns/op -> 104585600 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `699207500 ns/op -> 679812100 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `312707000 ns/op -> 420859900 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `504470600 ns/op -> 616614500 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1304217200 ns/op -> 1410249700 ns/op`

- repeat-aware `.xls6` results:
  - current retained baseline total: `10362 ms`
  - experiment total: `10106 ms`
  - per-file:
    - `002505.xls`: `1373 -> 1427 ms`
    - `006087.xls`: `1277 -> 1234 ms`
    - `008055.xls`: `4064 -> 3961 ms`
    - `016161.xls`: `1711 -> 1632 ms`
    - `019088.xls`: `993 -> 928 ms`
    - `019089.xls`: `944 -> 924 ms`
  - output parity held

- initial mixed-keyset result was invalid:
  - the first `perf-exp-xls-visible-only-fallback-keyset.json` run overlapped with a concurrent
    `go test ./...` process on the same machine
  - that made the mixed-keyset timing unsuitable as a retention decision

- corrected repeat-aware mixed keyset rerun in isolation:
  - `testdata/web-samples/reports/perf-exp-xls-visible-only-fallback-keyset-rerun-clean.json`
  - `testdata/web-samples/reports/perf-exp-xls-visible-only-fallback-keyset-rerun-clean.csv`
  - corrected results against the current retained baseline:
    - `.docx`: `6931 -> 7035 ms`
    - `.xls`: `17830 -> 13237 ms`
    - `.xlsx`: `23756 -> 21139 ms`
  - output parity held on all 30 inputs

- repository regression in isolation:
  - `ok officeread 133.940s`

- interpretation:
  - the earlier rejection was caused by measurement contamination, not by the optimization itself
  - with a clean mixed-keyset rerun, the fallback keeps the strong `.xls6` win and also improves
    the broader mixed workload overall
  - `.docx` regresses only slightly, while both spreadsheet buckets improve materially, which is the
    more relevant integrated signal for this `.xls`-specific fallback

- decision:
  - Retained.

## 2026-07-06 measured: the visible-only fallback pattern covers all 24 exact-miss files in the 1000-file `.xls` corpus, but only 9 of them are large enough to justify the current gate

- validation:
  - full-corpus scope probe:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSVisibleOnlyScope$' -v ./`

- results:
  - corpus totals:
    - `total=1000`
    - `exactMiss=24`
    - current gate hits: `9`
    - current gate hits covered by `exact+visible-only`: `9`
    - out-of-gate exact misses also covered by `exact+visible-only`: `15`
  - current gated large-shape files:
    - `000034-2.xls`, `000034.xls`, `000055.xls`, `001138.xls`, `006086.xls`, `006087.xls`,
      `007260.xls`, `014479.xls`, `022223.xls`
  - additional exact-miss files that also cover with `exact+visible-only`:
    - `000051.xls`, `003995.xls`, `004703.xls`, `006812.xls`, `009715.xls`, `009933.xls`,
      `010276.xls`, `010480.xls`, `010701.xls`, `015300.xls`, `018541.xls`, `022230.xls`,
      `022238.xls`, `022239.xls`, `022241.xls`

- interpretation:
  - semantically, the visible-only fallback pattern matches every exact-miss file in the current
    downloaded `.xls` corpus
  - performance-wise, that does not imply the gate should be removed, because many of the newly
    uncovered files are much smaller than the retained hotspot shape

- decision:
  - No code change from this measurement.

## 2026-07-06 rejected: broaden the visible-only fallback from the large-shape gate to every `.xls` exact miss

- experiment:
  - keep the retained `exact+visible-only` fallback logic
  - remove only the size/line gate
  - let every `.xls` exact miss try `exact+visible-only` before the full second-stage build

- rationale:
  - since all `24` exact-miss files in the 1000-file `.xls` corpus can be resolved by
    `exact+visible-only`, it was natural to test whether the large-shape gate could be dropped

- validation:
  - focused hotspots:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/006087.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
  - repeat-aware `.xls6` rerun:
    - `testdata/web-samples/reports/perf-exp-xls-visible-only-all-exactmiss-xls6.json`
    - `testdata/web-samples/reports/perf-exp-xls-visible-only-all-exactmiss-xls6.csv`

- focused benchmark results:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `167294500 ns/op -> 164642200 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1665857100 ns/op -> 1840787400 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `692025000 ns/op -> 932711500 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `104585600 ns/op -> 100448400 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `679812100 ns/op -> 641924700 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `420859900 ns/op -> 295391600 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `616614500 ns/op -> 479961000 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1410249700 ns/op -> 1273455200 ns/op`

- repeat-aware `.xls6` results against the retained gated fallback:
  - retained total: `10106 ms`
  - experiment total: `10520 ms`
  - per-file:
    - `002505.xls`: `1427 -> 1433 ms`
    - `006087.xls`: `1234 -> 1445 ms`
    - `008055.xls`: `3961 -> 4077 ms`
    - `016161.xls`: `1632 -> 1708 ms`
    - `019088.xls`: `928 -> 948 ms`
    - `019089.xls`: `924 -> 909 ms`

- interpretation:
  - the broader fallback is semantically safe on the measured corpus
  - but the retained gate exists for a real runtime reason: small exact-miss files do not repay the
    visible-only preflight overhead
  - dropping the gate therefore makes the optimization less selective in exactly the wrong way
  - an isolated rerun on exactly the affected `24` exact-miss files confirms the same thing:
    - gated exact-miss subset: `14778 ms`
    - broadened exact-miss subset: `15418 ms`
    - delta: `+640 ms`
  - that removes exact-hit `.xls` files from the picture and still shows the gate is worthwhile
    inside the exact-miss subset itself
  - a further isolated rerun on only the `15` out-of-gate exact-miss files with `repeat 7` narrows
    the picture:
    - gated out-of-gate subset: `1029 ms`
    - broadened out-of-gate subset: `1024 ms`
    - delta: `-5 ms`
  - so even on the files newly enabled by dropping the gate, the upside is only noise-level; there
    is no meaningful retained win hiding there

- decision:
  - Reverted.

## 2026-07-06 revisited and retained: single-pass `simpleInlineWorksheetCandidate(...)` scan for huge inlineStr worksheets

- idea:
  - keep the existing `.xlsx` simple-inline worksheet fast path unchanged semantically
  - optimize only the candidate detector:
    - replace the repeated full-buffer `bytes.Contains(...)` checks with one linear byte scan
    - detect the same required / forbidden marker set during that scan

- rationale:
  - current profiling of
    `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx` showed that
    `simpleInlineWorksheetCandidate(...)` was still the dominant hot function in the text-only
    inline path
  - even after earlier retained `.xlsx` work, the candidate check was still rescanning the same huge
    worksheet bytes too many times
  - this revisits an earlier rejected 2026-07-05 experiment from a different baseline; the retained
    decision here is based on fresh profiles and current integrated reruns

- validation:
  - focused `.xlsx` hotspots:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXWorksheetTextHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots' -benchmem -benchtime=1x ./`
  - profile:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench '^BenchmarkXLSXSimpleInlineTextOnlyHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchtime=1x -cpuprofile tmp-shape/xlsx-simpleinline-textonly-candidate-fast.prof ./`
    - `& 'C:\Program Files\Go\bin\go.exe' tool pprof -top ./officeread.test.exe tmp-shape/xlsx-simpleinline-textonly-candidate-fast.prof`
  - repeat-aware mixed keyset rerun:
    - `testdata/web-samples/reports/perf-exp-xlsx-inline-candidate-scan-keyset.json`
    - `testdata/web-samples/reports/perf-exp-xlsx-inline-candidate-scan-keyset.csv`
  - repository regression:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

- focused benchmark results:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `4474042700 ns/op -> 3553883100 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `3189833500 ns/op -> 2209982100 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2039197300 ns/op -> 2035989800 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `3755042400 ns/op -> 3274889600 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1548595600 ns/op -> 1293209700 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1178737800 ns/op -> 1158989000 ns/op`

- profile result:
  - `simpleInlineWorksheetCandidate(...)` remained the top frame in the text-only inline profile
  - but its cumulative cost dropped further, from the earlier `~3.10 s` range to about `2.30 s`

- repeat-aware mixed keyset results:
  - `.docx`: `6931 -> 6214 ms`
  - `.xls`: `17830 -> 11725 ms`
  - `.xlsx`: `23756 -> 21647 ms`
  - output parity held on all 30 inputs

- repository regression:
  - `ok officeread 124.196s`

- interpretation:
  - this is a narrow but meaningful `.xlsx` win: it speeds up recognition of huge inlineStr
    worksheets without changing extraction behavior once the fast path is entered
  - the direct evidence is strongest on the huge `.xlsx` hotspots and the `.xlsx` bucket in the
    mixed keyset rerun
  - the `.docx` and `.xls` mixed-keyset improvements should be treated as run variance because the
    code change does not touch those paths

- decision:
  - Retained.

## 2026-07-06 retained: short-cell cache for `.xls` exact-set table-cell normalization

- idea:
  - keep the retained `.xls` exact/backfill structure unchanged at a high level
  - reduce repeated work only inside `markdownBackfillExactSet(...)`
  - cache normalized table-cell variants only for raw cell fragments whose byte length is `<= 8`

- rationale:
  - table-cell duplication is real:
    - `006087.xls`: `235659` total cells, `37952` unique, `197707` duplicates
    - `008055.xls`: `1313143` total cells, `1288825` unique, `24318` duplicates
    - `016161.xls`: `551263` total cells, `520599` unique, `30664` duplicates
  - but duplicate pressure is strongly concentrated in short raw cell fragments:
    - `006087.xls`:
      - `<=8`: `620670 dup / 34688 unique`
      - `9-16`: `40709 dup / 3099 unique`
      - `17-32`: `50632 dup / 102 unique`
    - `008055.xls`:
      - `<=8`: `79767 dup / 31039 unique`
      - `9-16`: `16120 dup / 617078 unique`
      - `17-32`: `1771 dup / 640619 unique`
    - `016161.xls`:
      - `<=8`: `95535 dup / 3196 unique`
      - `9-16`: `855 dup / 46582 unique`
      - `17-32`: `20124 dup / 470823 unique`
  - this explains why a broad cell cache is too expensive on `008055.xls`, while a short-cell-only
    cache has a better chance of surviving integrated reruns

- implementation:
  - add `markdownVisibleTableCellsWithCache(...)`
  - factor per-cell normalization into `markdownVisibleTableCellVariants(...)`
  - add a local `tableCellCache` map inside `markdownBackfillExactSet(...)`
  - only cache entries when the raw cell fragment length is `<= 8`

- validation:
  - focused hotspots:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/006087.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
  - repeat-aware `.xls6` rerun:
    - `testdata/web-samples/reports/perf-exp-xls-short-cell-cache8-xls6.json`
    - `testdata/web-samples/reports/perf-exp-xls-short-cell-cache8-xls6.csv`
  - repeat-aware mixed keyset rerun:
    - `testdata/web-samples/reports/perf-exp-short-cell-cache8-keyset.json`
    - `testdata/web-samples/reports/perf-exp-short-cell-cache8-keyset.csv`
  - full repository regression:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

- focused benchmark results:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `760822200 ns/op -> 645141700 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `2104182200 ns/op -> 1625965900 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `814307500 ns/op -> 638316500 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `173512500 ns/op -> 110346100 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 699207500 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `278034100 ns/op -> 312707000 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `495409100 ns/op -> 504470600 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1489791600 ns/op -> 1304217200 ns/op`

- repeat-aware `.xls6` results:
  - retained baseline total: `10461 ms`
  - experiment total: `10362 ms`
  - per-file:
    - `002505.xls`: `1387 -> 1373 ms`
    - `006087.xls`: `1351 -> 1277 ms`
    - `008055.xls`: `3922 -> 4064 ms`
    - `016161.xls`: `1665 -> 1711 ms`
    - `019088.xls`: `1077 -> 993 ms`
    - `019089.xls`: `1059 -> 944 ms`
  - output parity held on every run

- repeat-aware mixed keyset results:
  - `.docx`: `6931 -> 6478 ms`
  - `.xls`: `17830 -> 12063 ms`
  - `.xlsx`: `23756 -> 22413 ms`
  - output parity held on all 30 inputs

- repository regression:
  - `ok officeread 140.019s`

- interpretation:
  - the isolated hotspot view is mixed for some exact-set-only benches, which is why the cache must
    stay narrow
  - despite that, both repeat-aware integrated reruns improved while preserving output parity, and
    the mixed keyset improved across every extension bucket
  - that makes the short-cell cache a valid retained optimization, unlike the earlier broader helper
    and exact-stage experiments

- decision:
  - Retained.

## 2026-07-06 rejected: extend the short-cell cache into coverage / containment table-cell expansion

- experiment:
  - keep the retained short-cell cache in `markdownBackfillExactSet(...)`
  - extend the same `<= 8` raw-cell cache into:
    - `markdownBackfillBuildCoverageContainment(...)`
    - `addMarkdownBackfillCoverage(...)`

- rationale:
  - short raw table cells repeat heavily, so it was reasonable to test whether the retained exact-set
    cache should also be reused on the coverage / containment side
  - this was intentionally narrow: same helper, same byte cutoff, local maps only

- validation:
  - focused hotspots:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/006087.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
  - repeat-aware `.xls6` rerun:
    - `testdata/web-samples/reports/perf-exp-short-cell-cache8-coverage-xls6.json`
    - `testdata/web-samples/reports/perf-exp-short-cell-cache8-coverage-xls6.csv`

- focused benchmark results:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `645141700 ns/op -> 574004600 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1625965900 ns/op -> 1661609000 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `638316500 ns/op -> 649785600 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `110346100 ns/op -> 140733800 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `699207500 ns/op -> 624811700 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `312707000 ns/op -> 287776800 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `504470600 ns/op -> 444791000 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1304217200 ns/op -> 1365105200 ns/op`

- repeat-aware `.xls6` results:
  - retained baseline total: `10362 ms`
  - experiment total: `10531 ms`
  - per-file:
    - `002505.xls`: `1373 -> 1342 ms`
    - `006087.xls`: `1277 -> 1186 ms`
    - `008055.xls`: `4064 -> 4344 ms`
    - `016161.xls`: `1711 -> 1633 ms`
    - `019088.xls`: `993 -> 928 ms`
    - `019089.xls`: `944 -> 1098 ms`
  - output parity held

- interpretation:
  - the cache extension helps some local coverage / containment paths, especially on `006087.xls`
  - but the integrated `.xls6` rerun still regresses overall because `008055.xls` and `019089.xls`
    give back too much
  - this means the retained exact-set-only cache remains the right boundary for now

- decision:
  - Reverted.

## 2026-07-06 measured: coverage-side cell entries are mostly real additions on the heavy `008055/016161` `.xls` samples

- validation:
  - contribution probe:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSCoverageCellContribution$' -v ./`
  - length-bucket probe:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSCoverageCellContributionByLength$' -v ./`

- contribution results:
  - `006087.xls`:
    - `rawAdds=48582`, `visibleAdds=0`, `cellTotal=235659`, `cellNew=37952`, `cellDup=197707`, `finalCoverage=130515`
  - `008055.xls`:
    - `rawAdds=16917`, `visibleAdds=0`, `cellTotal=1313143`, `cellNew=1288825`, `cellDup=24318`, `cellComparableNew=2`, `cellComparableDup=74`, `finalCoverage=1322621`
  - `016161.xls`:
    - `rawAdds=45630`, `visibleAdds=0`, `cellTotal=551263`, `cellNew=520599`, `cellDup=30664`, `cellComparableNew=2`, `finalCoverage=611643`

- length-bucket results:
  - `006087.xls`:
    - `<=8`: `37653 new / 127487 dup`
    - `9-16`: `164 new / 19594 dup`
    - `17-32`: `74 new / 50607 dup`
  - `008055.xls`:
    - `<=8`: `47813 new / 6920 dup`
    - `9-16`: `604586 new / 15520 dup`
    - `17-32`: `636338 new / 1781 dup`
  - `016161.xls`:
    - `<=8`: `49616 new / 10394 dup`
    - `9-16`: `9744 new / 673 dup`
    - `17-32`: `461239 new / 19597 dup`

- interpretation:
  - this is the key reason the coverage-side cache extension failed to retain:
    - exact-set duplicates are real enough to exploit narrowly
    - coverage-side cells on `008055.xls` and `016161.xls` are mostly genuine new coverage keys
  - especially in the `9..32` byte range, coverage-side table cells are overwhelmingly additive on
    the heavy samples
  - so a naive coverage-side pruning rule based on shortness or generic duplicate expectations is
    aimed at the wrong distribution

- decision:
  - No code change from this measurement.

## 2026-07-06 rejected: `.xls` exact precheck direct-hit fast path without duplicate direct lookup

- Idea:
  - revisit the earlier rejected `.xls` direct-first strategy, but remove its obvious duplicated
    direct lookup
  - do:
    - direct `exact[line]`
    - on miss only, run a helper that checks non-direct variants without repeating the direct hit
- Why:
  - temporary instrumentation still showed that the retained heavy `.xls` samples were almost
    purely direct hits:
    - `006087.xls`: `direct=40935`, `miss=404`
    - `008055.xls`: `direct=1181169`, `miss=0`
    - `016161.xls`: `direct=461603`, `miss=0`
  - fresh exact-stage timing also showed large exact-scan cost on the covered heavy samples:
    - `008055.xls`: `exactBuild=634ms`, `exactScan=977ms`
    - `016161.xls`: `exactBuild=301ms`, `exactScan=322ms`
- Result:
  - the exact-stage timing regressed instead of improving:
    - `002505.xls`: `exactBuild=271ms -> 366ms`, `exactScan=122ms -> 170ms`
    - `006087.xls`: `exactBuild=166ms -> 229ms`, `exactScan=18ms -> 19ms`
    - `008055.xls`: `exactBuild=634ms -> 780ms`, `exactScan=977ms -> 1320ms`
    - `016161.xls`: `exactBuild=301ms -> 378ms`, `exactScan=322ms -> 435ms`
  - focused benchmarks also regressed sharply:
    - `BenchmarkExtractXLSHotspots/006087.xls`: `1348264600 ns/op -> 1352256600 ns/op`
    - `BenchmarkExtractXLSHotspots/008055.xls`: `4141427600 ns/op -> 4623366400 ns/op`
    - `BenchmarkExtractXLSHotspots/016161.xls`: `1766906500 ns/op -> 2015096000 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 906804600 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 2118399500 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 845987500 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 620281700 ns/op`
- Interpretation:
  - removing the duplicated direct lookup was not enough to rescue the direct-first family
  - this strongly suggests the retained exact precheck slowdown is not explained by “generic helper
    repeats the direct hit” alone
- Decision:
  - Reverted.

## 2026-07-06 rejected: skip `seen` dedupe in `.xls` exact precheck

- Idea:
  - remove the `seen` dedupe map only on the `.xls` exact precheck path, since correctness does not
    depend on it and it looked redundant on the retained hotspot set
- Why:
  - temporary instrumentation on the retained `.xls6` hotspot files showed zero duplicate exact
    source lines across the board:
    - `002505.xls`: `lines=215015`, `dups=0`
    - `006087.xls`: `lines=41347`, `dups=0`
    - `008055.xls`: `lines=1181198`, `dups=0`
    - `016161.xls`: `lines=461615`, `dups=0`
    - `019088.xls`: `lines=47549`, `dups=0`
    - `019089.xls`: `lines=51514`, `dups=0`
- Result:
  - focused results were mixed and mostly bad:
    - `BenchmarkExtractXLSHotspots/006087.xls`: `1348264600 ns/op -> 1371871000 ns/op`
    - `BenchmarkExtractXLSHotspots/008055.xls`: `4141427600 ns/op -> 4279031400 ns/op`
    - `BenchmarkExtractXLSHotspots/016161.xls`: `1766906500 ns/op -> 2005530600 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 897860900 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 1487665800 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 605652400 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 856889900 ns/op`
  - the decisive repeat-aware `.xls6` rerun lost heavily:
    - total: `10461 ms -> 15375 ms`
    - `002505.xls`: `1387 ms -> 1517 ms`
    - `006087.xls`: `1351 ms -> 1869 ms`
    - `008055.xls`: `3922 ms -> 4738 ms`
    - `016161.xls`: `1665 ms -> 2123 ms`
    - `019088.xls`: `1077 ms -> 1573 ms`
    - `019089.xls`: `1059 ms -> 1555 ms`
- Interpretation:
  - zero duplicate lines did not mean the `seen` path was free to remove
  - this is another case where a superficially obvious local simplification still harms the
    retained integrated `.xls` workload
- Decision:
  - Reverted.

## 2026-07-06 rejected: plain-source fast return in `markdownBackfillCandidateLine(...)`

- Idea:
  - add a fast return to `markdownBackfillCandidateLine(...)` when the trimmed source line already
    survives `cleanTextFastPath(...)` unchanged
- Why:
  - new stage timing showed that the retained `.xls` markdown path was dominated by the
    `missingMarkdownTextXLS(...)` backfill stage rather than table assembly:
    - `006087.xls`: `tables=137ms`, `assembly=13ms`, `backfill=772ms`
    - `008055.xls`: `tables=768ms`, `assembly=32ms`, `backfill=1648ms`
    - `016161.xls`: `tables=344ms`, `assembly=14ms`, `backfill=775ms`
  - temporary source-shape instrumentation then showed that `markdownBackfillCandidateLine(...)`
    was effectively identity on almost every retained `.xls` source line:
    - `006087.xls`: `lines=41347`, `plain=41336`, `changed=0`
    - `008055.xls`: `lines=1181198`, `plain=1181185`, `changed=0`
    - `016161.xls`: `lines=461615`, `plain=461613`, `changed=0`
- Result:
  - despite that shape, the real backfill stage got slower:
    - `006087.xls`: `backfill=772ms -> 1086ms`
    - `008055.xls`: `backfill=1648ms -> 2067ms`
    - `016161.xls`: `backfill=775ms -> 836ms`
  - focused benchmarks also regressed:
    - `BenchmarkExtractXLSHotspots/008055.xls`: `4141427600 ns/op -> 4232403500 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 2035422700 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 725233000 ns/op`
- Interpretation:
  - even when candidate-line output is unchanged, the generic normalization path is apparently
    still performing work that materially helps the retained exact coverage checks
  - this is another case where an “identity-shape” shortcut does not preserve the right integrated
    performance behavior
- Decision:
  - Reverted.

## 2026-07-06 rejected: BIFF plain-cell fast path in markdown cell cleanup / prepare

- Idea:
  - move the optimization target forward from exact/backfill into BIFF markdown cell preparation
  - add a fast path in `cleanMarkdownTableCellValue(...)` for already-clean single-line values
  - add a matching fast path in `prepareMarkdownTableCellValue(...)` for already-prepared plain
    values with no newline or escaping needs
- Why:
  - the new full `Extract(008055.xls)` profile showed that the end-to-end hotspot still had large
    earlier BIFF markdown work:
    - `biffMarkdownWithText`
    - `biffTextParts`
    - `biffWorksheetMarkdownTables`
    - `cleanMarkdownTableCellValue`
  - temporary BIFF cell-shape instrumentation showed the retained `.xls` hotspots were
    overwhelmingly plain at this stage:
    - `006087.xls`: `raw=235659`, `cleanChanged=0`, `rawShapes nl=0 slash=0 pipe=0 rtf=0`
    - `008055.xls`: `raw=1313143`, `cleanChanged=0`, `rawShapes nl=0 slash=0 pipe=0 rtf=0`
    - `016161.xls`: `raw=551263`, `cleanChanged=0`, `rawShapes nl=0 slash=0 pipe=0 rtf=0`
  - that made the BIFF table-cell clean/prepare pair look like a plausible true end-to-end target
- Result:
  - full extract did improve on two of the three retained hotspots:
    - `BenchmarkExtractXLSHotspots/006087.xls`: `1348264600 ns/op -> 1295386300 ns/op`
    - `BenchmarkExtractXLSHotspots/008055.xls`: `4141427600 ns/op -> 3794928300 ns/op`
    - `BenchmarkExtractXLSHotspots/016161.xls`: `1766906500 ns/op -> 1798147600 ns/op`
  - but the retained `.xls6` rerun still regressed badly:
    - total: `10461 ms -> 14914 ms`
    - `002505.xls`: `1387 ms -> 1428 ms`
    - `006087.xls`: `1351 ms -> 1683 ms`
    - `008055.xls`: `3922 ms -> 4862 ms`
    - `016161.xls`: `1665 ms -> 2205 ms`
    - `019088.xls`: `1077 ms -> 1351 ms`
    - `019089.xls`: `1059 ms -> 1385 ms`
- Interpretation:
  - this was the clearest recent example of a candidate that helped a true full-extract benchmark
    yet still failed the retained repeat-aware rerun gate
  - the remaining problem is not just "optimize a hot helper"; it is finding an optimization whose
    local win survives the integrated multi-file `.xls` workload
- Decision:
  - Reverted.

## 2026-07-06 rejected: helper early-return guards for markdown exact/backfill cleanup

- Idea:
  - keep the retained `.xls` exact / backfill structure unchanged
  - only add helper-local fast returns for obviously no-op cases:
    - `markdownInlineLinkVisibleText(...)` when there is no `[`
    - `stripMarkdownInlineWrappers(...)` when there is no `*_~``
    - `unescapeMarkdownInlineFormattingMarkers(...)` when there is no `\`
    - `unescapeMarkdownVisibleText(...)` when there is no `\`
- Why:
  - new temporary variant-shape instrumentation showed that exact precheck traffic itself was not
    generating meaningful alternate variants:
    - `006087.xls`: `allowed=41339`, `visibleAdds=0`, `markdownAdds=0`, `escapedAdds=0`
    - `008055.xls`: `allowed=1181169`, `visibleAdds=0`, `markdownAdds=2`, `escapedAdds=0`
    - `016161.xls`: `allowed=461603`, `visibleAdds=0`, `markdownAdds=2`, `escapedAdds=0`
  - the current exact-set profile still showed helper-string costs such as `strings.Replace`,
    `strings.TrimSpace`, `strings.IndexAny`, and builder traffic
  - that made pure helper-local no-op guards look like a strong low-risk candidate
- Result:
  - focused hotspot benchmarks improved dramatically:
    - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 408463000 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 1419166500 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 567112100 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `173512500 ns/op -> 83928700 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 377578500 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `278034100 ns/op -> 188349000 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `495409100 ns/op -> 326220200 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1489791600 ns/op -> 888473300 ns/op`
  - allocation volume also dropped sharply on the focused benches:
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `215699 allocs/op -> 63459 allocs/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `198340 allocs/op -> 46111 allocs/op`
  - but the decisive repeat-aware `.xls6` rerun still lost:
    - total: `10461 ms -> 11871 ms`
    - `002505.xls`: `1387 ms -> 1344 ms`
    - `006087.xls`: `1351 ms -> 1313 ms`
    - `008055.xls`: `3922 ms -> 4632 ms`
    - `016161.xls`: `1665 ms -> 2116 ms`
    - `019088.xls`: `1077 ms -> 1300 ms`
    - `019089.xls`: `1059 ms -> 1166 ms`
- Interpretation:
  - this was one of the clearest isolated benchmark wins in the whole `.xls` optimization thread
  - even so, it still failed the retained integrated rerun gate, reinforcing that helper-level
    alloc and microbenchmark wins are not sufficient proof for this workload
- Decision:
  - Reverted.

## 2026-07-06 rejected: shape-gated large exact-set capacity hint for very wide table markdown

- Idea:
  - revisit exact-set preallocation, but only for extremely wide table markdown where line-count
    capacity is obviously too small
  - estimate:
    - `lineCount = strings.Count(markdown, "\n") + 1`
    - `pipeCount = strings.Count(markdown, "|")`
    - when `pipeCount >= lineCount * 40`, initialize exact capacity with `lineCount + pipeCount`
- Why:
  - the current-baseline `008055.xls` profile again showed meaningful exact-set map pressure:
    - `runtime.mapassign_faststr`
    - `internal/runtime/maps.(*table).split`
    - `internal/runtime/maps.(*table).rehash`
  - temporary capacity-shape instrumentation showed that `008055.xls` was far wider than the other
    retained `.xls` hotspots:
    - `006087.xls`: `exact=86539`, `lines=54606`, `pipes=799984`, `lines+pipes=854590`
    - `008055.xls`: `exact=1305746`, `lines=16922`, `pipes=1403530`, `lines+pipes=1420452`
    - `016161.xls`: `exact=566236`, `lines=45647`, `pipes=682736`, `lines+pipes=728383`
- Result:
  - exact-set-only on the targeted wide-table case did improve:
    - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 547445700 ns/op`
  - but the integrated hotspot mix still regressed:
    - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 748556100 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 2000567900 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 820065400 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `173512500 ns/op -> 196222100 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `278034100 ns/op -> 297024100 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1489791600 ns/op -> 1331312700 ns/op`
  - the decisive repeat-aware `.xls6` rerun lost heavily:
    - total: `10461 ms -> 12852 ms`
    - `002505.xls`: `1387 ms -> 1448 ms`
    - `006087.xls`: `1351 ms -> 1657 ms`
    - `008055.xls`: `3922 ms -> 4832 ms`
    - `016161.xls`: `1665 ms -> 2186 ms`
    - `019088.xls`: `1077 ms -> 1304 ms`
    - `019089.xls`: `1059 ms -> 1425 ms`
- Interpretation:
  - the exact-set map really is expensive on `008055.xls`, but even a shape-specific capacity hint
    does not produce a retained end-to-end win
  - that pushes the next optimization step away from raw map sizing and toward work that changes the
    amount of exact-stage traffic or string materialization without disturbing the retained balance
- Decision:
  - Reverted.

## 2026-07-06 rejected: skip table-row `markdownVisibleLineText(...)` exact entries except empty-cell-gap shapes

- Idea:
  - on the `.xls` exact-set path, stop adding row-level `markdownVisibleLineText(...)` exact
    entries for ordinary table rows
  - keep a tiny exception for rows containing `|  |`, because the temporary instrumentation showed
    that the only retained table-row visible exact additions on `006087.xls` came from empty-cell
    gap collapse
- Why:
  - new origin instrumentation showed that table-row visible exact entries were essentially all
    duplicates:
    - `006087.xls`: `tableVisibleAdds=2`, `tableVisibleDup=49997`
    - `008055.xls`: `tableVisibleAdds=0`, `tableVisibleDup=16910`
    - `016161.xls`: `tableVisibleAdds=0`, `tableVisibleDup=45621`
  - on `008055.xls`, the exact-set origin mix was overwhelmingly cell-driven:
    - `rawAdds=16917`
    - `visibleAdds=4` (all non-table)
    - `cellAdds=1288825`
- Result:
  - the local duplicate story was real, but the integrated hotspot mix still regressed:
    - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 763505800 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 2171014300 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 623884600 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `173512500 ns/op -> 184974700 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 631996100 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `278034100 ns/op -> 473723000 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `495409100 ns/op -> 492549900 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1489791600 ns/op -> 1334273700 ns/op`
- Interpretation:
  - this was an unusually precise duplicate-removal candidate, yet it still did not produce a
    usable retained win
  - the retained `.xls` exact-set path appears to depend on structural choices that simple
    cardinality or duplicate-origin counts do not capture
- Decision:
  - Reverted.

## 2026-07-06 rejected: table-cell helper fast path for markdown exact-set / backfill cleanup

- Idea:
  - narrow the table-cell cleanup cost inside the retained `.xls` exact-set path by only running
    the expensive transforms when a cell still carries the relevant syntax markers
  - route both the direct cell cleanup helper and the `<br>` split segments through the same
    trigger-based cleanup path
- Why:
  - the current `008055.xls` exact-set hotspot still showed visible table-cell cleanup and string
    normalization costs in profile, even after the retained short-digit visible substring work
  - the old helper still paid for the full cleanup chain on every cell, including already-plain
    cells
- Result:
  - focused hotspot benchmarks improved strongly:
    - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 510587000 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 1470653300 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 650053900 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 412054200 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `495409100 ns/op -> 389228900 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1489791600 ns/op -> 1086864100 ns/op`
  - full repository regression still passed before revert:
    - `ok officeread 125.447s`
  - but the repeat-aware retained `.xls6` rerun lost:
    - total: `10461 ms -> 11748 ms`
    - `002505.xls`: `1387 ms -> 1293 ms`
    - `006087.xls`: `1351 ms -> 1446 ms`
    - `008055.xls`: `3922 ms -> 4573 ms`
    - `016161.xls`: `1665 ms -> 2159 ms`
    - `019088.xls`: `1077 ms -> 1130 ms`
    - `019089.xls`: `1059 ms -> 1147 ms`
- Interpretation:
  - the helper fast path was real in isolation, but it did not preserve the best integrated `.xls`
    behavior
  - for this workload, a narrower local cleanup path is still not enough evidence to beat the
    retained baseline
- Decision:
  - Reverted.

## 2026-07-06 rejected: prune table-row line entries from `.xls` exact sets

- Idea:
  - keep all visible table-cell exact entries, but on the `.xls` exact-precheck path stop inserting
    whole table-row strings and row-level visible-text strings into the exact set
- Why:
  - the new origin breakdown for retained `.xls` exact hits looked overwhelmingly cell-driven:
    - `008055.xls`: `rawHits=0`, `visibleHits=3`, `cellHits=1181166`, `miss=0`
    - `016161.xls`: `rawHits=0`, `visibleHits=6`, `cellHits=461597`, `miss=0`
    - `006087.xls`: `rawHits=4598`, `visibleHits=1`, `cellHits=36336`, `miss=404`
  - a temporary replay confirmed that this pruned exact set preserved the same hit/miss totals on
    those hotspots while materially shrinking set size, especially on `016161.xls` and `006087.xls`
- Result:
  - the retained hotspot mix still got slower:
    - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 845788200 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 2210761600 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 870965300 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `173512500 ns/op -> 235957700 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 853177900 ns/op`
    - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `278034100 ns/op -> 384341000 ns/op`
  - repeat-aware `.xls6` rerun also regressed:
    - total: `10461 ms -> 11677 ms`
    - `002505.xls`: `1387 ms -> 1650 ms`
    - `006087.xls`: `1351 ms -> 1758 ms`
    - `008055.xls`: `3922 ms -> 5085 ms`
    - `016161.xls`: `1665 ms -> 2069 ms`
    - `019088.xls`: `1077 ms -> 1115 ms`
    - `019089.xls`: `1059 ms -> 1100 ms`
- Interpretation:
  - this is a useful negative result: even when table-row line entries look redundant from a direct
    hit-origin perspective, dropping them still hurts the real integrated workload
  - exact-set cardinality reduction by itself is not a reliable win signal on this path
- Decision:
  - Reverted.

## 2026-07-06 rejected: `.xls` direct-exact-first exact precheck

- Idea:
  - leave the generic exact precheck untouched for other formats
  - only on the `.xls` path, split the exact-stage membership test into:
    - direct `exact[line]`
    - then full `markdownBackfillExactSetContainsLine(...)` only on miss
- Why:
  - temporary instrumentation showed a striking `.xls`-only hit shape:
    - `006087.xls`: `direct=40935`, `visible=0`, `markdown=0`, `escaped=0`, `misses=404`
    - `008055.xls`: `direct=1181169`, `visible=0`, `markdown=0`, `escaped=0`, `misses=0`
    - `016161.xls`: `direct=461603`, `visible=0`, `markdown=0`, `escaped=0`, `misses=0`
  - that made it look as if the `.xls` exact precheck could avoid the general helper almost all the
    time while preserving semantics exactly
- Result:
  - the retained hotspot mix got slower instead of faster:
    - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `700638000 ns/op -> 757829500 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1589953500 ns/op -> 2353962100 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `658728600 ns/op -> 865504500 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: about `803362600 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: about `1555288300 ns/op`
  - repeat-aware `.xls6` rerun also regressed against the retained baseline:
    - total: `10461 ms -> 12051 ms`
    - `002505.xls`: `1387 ms -> 1436 ms`
    - `006087.xls`: `1351 ms -> 1692 ms`
    - `008055.xls`: `3922 ms -> 4933 ms`
    - `016161.xls`: `1665 ms -> 1771 ms`
    - `019088.xls`: `1077 ms -> 1062 ms`
    - `019089.xls`: `1059 ms -> 1157 ms`
- Interpretation:
  - even with a near-perfect direct-hit distribution, duplicating the direct lookup before the
    generic helper is still a net loss on the retained `.xls` workloads
  - this suggests the cost center is not simply "too many helper calls" in isolation
- Decision:
  - Reverted.

## 2026-07-05 rejected: pre-size `.xls` markdown-backfill exact / coverage maps from line count

- Idea:
  - pre-split markdown in `markdownBackfillExactSet(...)` and
    `markdownBackfillBuildCoverageContainment(...)`
  - initialize `exact` and `coverage` with `len(lines)` capacity
- Why:
  - the new `008055.xls` numbers still showed map-heavy work, so a narrow capacity hint looked
    worth checking
- Result:
  - focused regression stayed green: `ok officeread 12.305s`
  - integrated backfill benchmark remained mixed and not good enough:
    - `006087.xls`: `1120972800 ns/op -> 1342938200 ns/op`
    - `008055.xls`: `2163713000 ns/op -> 2535357200 ns/op`
    - `016161.xls`: `1017316800 ns/op -> 924406600 ns/op`
- Decision:
  - Reverted.

## 2026-07-05 retained: reuse visible table text when deriving `.xls` comparable containment

- Change:
  - introduce `markdownBackfillComparableFromVisibleText(...)`
  - in the `.xls` markdown containment builders:
    - `markdownBackfillBuildCoverageContainment(...)`
    - `markdownBackfillContainmentSet(...)`
  - compute `visible := markdownBackfillVisibleText(...)` once for each table row and derive the
    comparable form from that visible text, instead of recomputing `markdownBackfillVisibleText(...)`
    a second time through `markdownBackfillComparableText(...)`
- Why:
  - the new split `.xls` benchmarks showed that containment construction still had very high
    allocation volume
  - this was a narrowly scoped way to remove duplicated normalization work on the retained table-row
    path

Validation:
- `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots|BenchmarkXLSMarkdownExactSetHotspots|BenchmarkXLSMarkdownCoverageContainmentHotspots' -benchmem -benchtime=1x ./`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-visible-comparable-reuse-keyset.json -csv testdata\web-samples\reports\perf-exp-visible-comparable-reuse-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`
- `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused regression:
  - `ok officeread 11.648s`
- full repository regression:
  - `ok officeread 129.909s`
- hotspot benchmarks:
  - whole backfill:
    - `006087.xls`: `1339661800 ns/op -> 1204912700 ns/op`
    - `008055.xls`: `2654517100 ns/op -> 1943883800 ns/op`
    - `016161.xls`: `1128520700 ns/op -> 800417600 ns/op`
  - exact-set-only:
    - `006087.xls`: `247864400 ns/op -> 265334900 ns/op`
    - `008055.xls`: `1124659900 ns/op -> 895644500 ns/op`
    - `016161.xls`: `539833200 ns/op -> 395615800 ns/op`
  - coverage+containment-only:
    - `006087.xls`: `569707600 ns/op -> 509293800 ns/op`
    - `008055.xls`: `1620559700 ns/op -> 1538942900 ns/op`
    - `016161.xls`: `997024500 ns/op -> 900134500 ns/op`
- repeat-aware retained mixed keyset:
  - retained baseline total: `44994 ms`
  - experiment total: `41943 ms`
  - `.docx`: `6433 ms -> 6107 ms`
  - `.xls`: `14471 ms -> 13012 ms`
  - `.xlsx`: `24090 ms -> 22824 ms`
- output parity against the retained current-keyset report:
  - `textBytes` / `images`: `NO_DIFF`

Decision:
- Retained. This optimization improved all retained `.xls` hotspot samples, won on the repeat-aware
  mixed keyset, preserved output, and passed the full test suite.

## 2026-07-05 retained: avoid unconditional HTML and `<br>` work in markdown table-cell extraction

- Change:
  - tighten `markdownVisibleTableCells(...)` so it:
    - only runs `markdownVisibleHTMLText(...)` when the normalized cell still contains `<`
    - only splits segments by `<br>` when the normalized cell actually contains `<br>`
- Why:
  - fresh post-change profiles still showed `markdownVisibleTableCells(...)` as a meaningful cost in
    the retained `.xls` exact-set and containment paths
  - most cells in the retained hotspots do not contain HTML tags or `<br>` markers, so the old code
    was doing unnecessary string work in the common case

Validation:
- `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots|BenchmarkXLSMarkdownExactSetHotspots|BenchmarkXLSMarkdownCoverageContainmentHotspots' -benchmem -benchtime=1x ./`
- `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-tablecell-htmlbr-guards-keyset.json -csv testdata\web-samples\reports\perf-exp-tablecell-htmlbr-guards-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`
- `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results versus the previous retained visible/comparable-reuse baseline:
- focused regression:
  - `ok officeread 11.506s`
- full repository regression:
  - `ok officeread 128.440s`
- whole `.xls` backfill benchmark:
  - `006087.xls`: `1204912700 ns/op -> 1048374300 ns/op`
  - `008055.xls`: `1943883800 ns/op -> 1964941500 ns/op`
  - `016161.xls`: `800417600 ns/op -> 671926200 ns/op`
- exact-set-only benchmark:
  - `006087.xls`: `265334900 ns/op -> 241945500 ns/op`
  - `008055.xls`: `895644500 ns/op -> 918970900 ns/op`
  - `016161.xls`: `395615800 ns/op -> 412875600 ns/op`
- coverage+containment-only benchmark:
  - `006087.xls`: `509293800 ns/op -> 483859100 ns/op`
  - `008055.xls`: `1538942900 ns/op -> 1343779200 ns/op`
  - `016161.xls`: `900134500 ns/op -> 784124000 ns/op`
- whole-backfill allocs/op:
  - `006087.xls`: `3003778 -> 2532354`
  - `008055.xls`: `1532722 -> 215688`
  - `016161.xls`: `968729 -> 413958`
- repeat-aware retained mixed keyset:
  - previous retained baseline total: `41943 ms`
  - experiment total: `40807 ms`
  - `.docx`: `6107 ms -> 5929 ms`
  - `.xls`: `13012 ms -> 12026 ms`
  - `.xlsx`: `22824 ms -> 22852 ms`
- output parity:
  - retained-current-keyset vs experiment `textBytes` / `images`: `NO_DIFF`

Decision:
- Retained. This is another stable `.xls` markdown-backfill win: output stayed identical, the mixed
  keyset improved again, and allocation pressure on the hotspot files dropped sharply.

## 2026-07-05 rejected: early unescaped-pipe counting inside `markdownLikelyTableRow(...)`

- Idea:
  - replace `len(splitMarkdownTableRow(trimmed)) >= 3` with a byte scan that returns once it finds
    two unescaped `|` separators
- Why:
  - the `006087.xls` profile still showed meaningful table-shape overhead, so avoiding a full split
    in the boolean row predicate looked like a clean implementation-level candidate
- Result:
  - focused regression stayed green: `ok officeread 11.280s`
  - exact-set-only benchmarks improved
  - but the integrated backfill path regressed across all three retained `.xls` hotspots:
    - `006087.xls`: `1048374300 ns/op -> 1130252700 ns/op`
    - `008055.xls`: `1964941500 ns/op -> 2102875200 ns/op`
    - `016161.xls`: `671926200 ns/op -> 774968100 ns/op`
- Decision:
  - Reverted.

## 2026-07-05 rejected: skip segment-level HTML cleanup when a `<br>` segment has no `<`

- Idea:
  - keep the retained outer `markdownVisibleTableCells(...)` guards
  - additionally, inside the `<br>` segment loop, only call `markdownVisibleHTMLText(...)` when the
    segment itself still contains `<`
- Why:
  - after the retained outer guards, this looked like one more low-risk empty-case shortcut
- Result:
  - focused regression stayed green: `ok officeread 14.047s`
  - exact-set-only benchmarks improved
  - but the integrated path regressed again:
    - `006087.xls`: `1048374300 ns/op -> 1248900300 ns/op`
    - `008055.xls`: `1964941500 ns/op -> 2019862500 ns/op`
    - `016161.xls`: `671926200 ns/op -> 1003614300 ns/op`
- Decision:
  - Reverted.

## 2026-07-05 rejected: skip visible-line substring containment for non-exact `|` queries

- Idea:
  - in `markdownBackfillContainment.visibleLineContainsLine(...)`, keep exact hits in
    `visibleExact`, but skip the joined-string `strings.Contains(...)` fallback when the query still
    contains `|`
- Why:
  - the `006087.xls` profile still pointed straight at visible-line substring containment, and
    table-like queries with `|` already have `tableTextContainsLine(...)` as another coverage path
- Result:
  - focused regression stayed green: `ok officeread 14.278s`
  - exact-set-only benchmarks improved
  - but the integrated backfill and coverage/containment benchmarks regressed sharply:
    - whole backfill:
      - `006087.xls`: `1048374300 ns/op -> 1107237700 ns/op`
      - `008055.xls`: `1964941500 ns/op -> 2163977600 ns/op`
      - `016161.xls`: `671926200 ns/op -> 920007200 ns/op`
    - coverage+containment:
      - `006087.xls`: `483859100 ns/op -> 794799600 ns/op`
      - `008055.xls`: `1343779200 ns/op -> 1914329200 ns/op`
      - `016161.xls`: `784124000 ns/op -> 1116100000 ns/op`
- Decision:
  - Reverted.

## 2026-07-05 benchmark refinement: replay real containment queries directly

- Added two new benchmark families in [extract_bench_test.go](/D:/workprj/officeread/extract_bench_test.go):
  - `BenchmarkXLSVisibleContainmentReplayHotspots`
  - `BenchmarkXLSTableContainmentReplayHotspots`
- Purpose:
  - rebuild the actual candidate/variant query set used by `missingMarkdownTextWithOptions(...)`
  - replay those queries directly against prepared containment state
  - separate visible-line containment cost from table containment cost

Validation:
- `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract_bench_test.go`
- `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls$' -benchmem -benchtime=1x ./`
- `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSTableContainmentReplayHotspots/006087.xls$' -benchmem -benchtime=1x ./`

Observed results on `006087.xls`:
- focused regression:
  - `ok officeread 18.636s`
- visible containment replay:
  - `11825385100 ns/op`, `0 B/op`, `0 allocs/op`
- table containment replay:
  - `6055400 ns/op`, `144 B/op`, `7 allocs/op`

Takeaway:
- `006087.xls` is overwhelmingly dominated by visible-line containment replay, not table containment
- the next worthwhile `.xls` optimization should target pure substring-search behavior inside
  `visibleLineContainsLine(...)` or its prepared data shape, not allocation trimming on the table
  path

## 2026-07-05 rejected: thresholded suffix-array index for large visible-line haystacks

- Idea:
  - for large `markdownBackfillContainment.visibleJoined` strings, build a suffix-array index and
    use that index inside `visibleLineContainsLine(...)` instead of plain `strings.Contains(...)`
- Why:
  - the new replay benchmarks showed that `006087.xls` was spending nearly all of its remaining
    time in repeated visible-line substring queries against one very large joined haystack
- Result:
  - focused regression stayed green: `ok officeread 12.091s`
  - `006087.xls` replay shape improved dramatically:
    - baseline replay: `11825385100 ns/op`, `0 B/op`, `0 allocs/op`
    - suffix-index replay: `26957200 ns/op`, `293280 B/op`, `36660 allocs/op`
  - but the retained integrated hotspot mix still lost:
    - `006087.xls`: `1108063500 ns/op -> 865251400 ns/op`
    - `008055.xls`: `1738540000 ns/op -> 2506948700 ns/op`
    - `016161.xls`: `781564800 ns/op -> 949258000 ns/op`
    - `008055.xls` containment-build path regressed sharply:
      - `1343779200 ns/op -> 3665734000 ns/op`
- Interpretation:
  - this is strong evidence that the algorithmic direction is valid for the `006087` query shape
  - but the current implementation pays too much build and lookup cost on the broader retained `.xls`
    set
- Decision:
  - Reverted.

## 2026-07-05 rejected: 4-byte visible substring prefilter

- Idea:
  - build a lightweight Bloom-style bitset of 4-byte windows from large `visibleJoined` haystacks
  - before `strings.Contains(...)`, reject any query whose first 4-byte window is definitely absent
- Why:
  - after the suffix-array experiment, this looked like a much cheaper way to target the same
    `006087.xls` replay bottleneck
- Result:
  - focused regression stayed green: `ok officeread 33.154s`
  - but the replay and integrated hotspot views both regressed:
    - `006087.xls` visible replay:
      - baseline: `10330479400 ns/op`
      - prefilter experiment: `15930762700 ns/op`
    - whole backfill:
      - `006087.xls`: `1108063500 ns/op -> 2576643600 ns/op`
      - `008055.xls`: `1738540000 ns/op -> 4603367800 ns/op`
      - `016161.xls`: `781564800 ns/op -> 1917064200 ns/op`
- Interpretation:
  - the filter was too lossy to prune enough real queries, so the extra build/check work simply
    added overhead on top of the existing substring search cost
- Decision:
  - Reverted.

## 2026-07-05 rejected: targeted exact 4-byte window set for the `006087` replay shape

- Idea:
  - after collecting containment-shape stats, only build a visible-line prefilter for haystacks that
    look like `006087.xls`
  - instead of a fuzzy bitset, use an exact set of 4-byte windows from `visibleJoined`
  - reject a query only when its leading 4-byte window is definitely absent from that exact set
- Why:
  - the earlier Bloom-style filter was too lossy
  - the fresh shape stats showed that `006087` and `008055` are radically different workloads:
    - `006087.xls`: `visibleJoined=4587667`, `visibleLines=54601`, `queries=41339`, `visibleMax=271`
    - `008055.xls`: `visibleJoined=22697129`, `visibleLines=16917`, `queries=1181171`, `visibleMax=11343`
- Result:
  - focused regression stayed green: `ok officeread 11.226s`
  - the integrated hotspot view was mixed:
    - `006087.xls`: `1174221600 ns/op -> 1152815600 ns/op`
    - `008055.xls`: `2713071900 ns/op -> 1609151700 ns/op`
    - `016161.xls`: `808511400 ns/op -> 739579100 ns/op`
  - but the actual target barely moved:
    - `006087.xls` visible replay:
      - baseline: `10330479400 ns/op`
      - exact-window experiment: `10305839900 ns/op`
- Interpretation:
  - the leading 4-byte window is too common across real `006087` queries to provide meaningful
    pruning, even when tracked exactly
- Decision:
  - Reverted.

## 2026-07-05 rejected: lazy first-8-byte visible prefix buckets

- Idea:
  - for `006087`-like visible replay shapes, lazily bucket `visibleLines` by the first 8 bytes of
    the query
  - on the first query for a given prefix, build and cache a joined subset of lines containing that
    prefix
  - then run `strings.Contains(...)` against that smaller bucket instead of the whole
    `visibleJoined`
- Why:
  - the shape stats showed that `006087.xls` had only `329` unique first-8-byte prefixes across
    `41339` replay queries, which made lazy bucketing look much cheaper than a global index
- Result:
  - focused regression stayed green: `ok officeread 10.850s`
  - integrated hotspot benchmarks looked somewhat positive:
    - `006087.xls`: `1083877100 ns/op -> 985549600 ns/op`
    - `008055.xls`: `2151532100 ns/op -> 1705911700 ns/op`
  - but the real target did not improve:
    - `006087.xls` visible replay:
      - baseline: `10855155300 ns/op`
      - prefix-bucket experiment: `10938179300 ns/op`
      - `39703000 B/op`, `935 allocs/op`
- Interpretation:
  - even after prefix partitioning, the expensive `006087` queries still land in large enough
    buckets that substring scanning remains dominant
- Decision:
  - Reverted.
## 2026-07-05 rejected: narrower shape-gated suffix-array replay index

- Idea:
  - retry the suffix-array direction with much tighter enablement:
    - `visibleJoined` between `1 MiB` and `8 MiB`
    - at least `30000` visible lines
    - `visibleMaxLen <= 512`
  - this was meant to catch `006087.xls`-like replay shapes while avoiding the clearly bad
    `008055.xls`-style cases
- Why:
  - the earlier broad suffix-array experiment showed a real algorithmic win on the replay hotspot,
    but it lost badly once index cost hit the broader `.xls` mix
- Result:
  - focused sample regression stayed green: `ok officeread 0.707s`
  - the `006087.xls` replay target improved again:
    - baseline replay: `10855155300 ns/op`, `0 B/op`, `0 allocs/op`
    - narrower shape-gated suffix replay: `30303700 ns/op`, `296976 B/op`, `36721 allocs/op`
  - but the retained integrated hotspot view still was not strong enough:
    - `006087.xls`: `1083877100 ns/op -> 822505700 ns/op`
    - `008055.xls`: `2151532100 ns/op -> 2110890500 ns/op`
    - `016161.xls`: `637898400 ns/op -> 705303900 ns/op`
    - containment-build view:
      - `006087.xls`: `767560000 ns/op`
      - `008055.xls`: `1993346500 ns/op`
- Interpretation:
  - the replay-target win is real and repeatable
  - but even this narrower gating still does not turn the suffix-array approach into a clearly good
    retained optimization for the broader hotspot set
- Decision:
  - Reverted.

## 2026-07-05 baseline rerun after reverting the narrower suffix gate

- Validation:
  - hotspot benchmarks:
    - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `1110686900 ns/op`, `302044800 B/op`, `2532336 allocs/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `2338682200 ns/op`, `552087192 B/op`, `215683 allocs/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `877876400 ns/op`, `253360952 B/op`, `413942 allocs/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `679919100 ns/op`, `235380488 B/op`, `2006870 allocs/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1557958400 ns/op`, `701203200 B/op`, `917397 allocs/op`
    - `BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls`: `14059815600 ns/op`, `0 B/op`, `0 allocs/op`
    - `BenchmarkXLSTableContainmentReplayHotspots/006087.xls`: `4019000 ns/op`, `144 B/op`, `7 allocs/op`
  - repeat-aware keyset rerun written to:
    - `testdata/web-samples/reports/perf-rerun-20260705-current-keyset-rerun2.json`
    - `testdata/web-samples/reports/perf-rerun-20260705-current-keyset-rerun2.csv`
  - rerun summary:
    - `.docx`: `2/2 ok`, `1213267` text bytes, `43` images, total `6931 ms`
    - `.xls`: `7/7 ok`, `32445758` text bytes, `0` images, total `17830 ms`
    - `.xlsx`: `21/21 ok`, `279385415` text bytes, `1` image, total `23756 ms`
- Note:
  - these totals are lower than the earlier `perf-rerun-20260705-current-keyset.json` reference,
    but because the code was already back on the same retained baseline after revert, this should be
    treated as ordinary run-to-run variance rather than a new retained win

## 2026-07-05 retained: shape-gated short-digit visible substring set

- Idea:
  - stop treating the remaining `006087.xls` visible replay hotspot as a generic substring-search
    problem
  - instead, when the visible haystack looks like the measured `006087` shape
    (`1 MiB <= visibleJoined <= 8 MiB`, at least `30000` visible lines, `visibleMaxLen <= 512`),
    precompute all ASCII-digit substrings of widths `4..7` from visible numeric runs
  - then let `visibleLineContainsLine(...)` answer short all-digit queries from that exact set
    instead of running `strings.Contains(...)` over the whole joined haystack
- Why:
  - the new integrated-path instrumentation showed that the real retained `.xls` path was no longer
    dominated by tens of thousands of visible replay checks
  - on `006087.xls`, it was down to only `404` visible substring queries, all hits, and the early
    examples were short numeric fragments like `61047`, `30357`, `21595`, `123311`, and `5200`
  - those are exactly the kind of queries that can be preserved with an exact short-digit substring
    set while eliminating the full-haystack scan cost
- Result:
  - focused regression stayed green: `ok officeread 11.537s`
  - full repository regression stayed green: `ok officeread 149.727s`
  - focused hotspot benchmarks improved:
    - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `1110686900 ns/op -> 760822200 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `2338682200 ns/op -> 2104182200 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `877876400 ns/op -> 814307500 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `679919100 ns/op -> 495409100 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1557958400 ns/op -> 1489791600 ns/op`
    - `BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls`: `14059815600 ns/op -> 22748600 ns/op`
    - `BenchmarkXLSTableContainmentReplayHotspots/006087.xls`: `4019000 ns/op -> 4088400 ns/op`
  - repeat-aware broader keyset rerun kept output parity on all 30 inputs and improved the `.xls`
    subset:
    - `.xls`: `17830 ms -> 15160 ms`
    - `.docx`: `6931 ms -> 7447 ms`
    - `.xlsx`: `23756 ms -> 27947 ms`
  - the decisive repeat-aware `.xls6` hotspot rerun also kept output parity and improved strongly:
    - total: `23111 ms -> 10461 ms`
    - `006087.xls`: `5905 ms -> 1351 ms`
    - `008055.xls`: `7629 ms -> 3922 ms`
    - `016161.xls`: `2786 ms -> 1665 ms`
    - `002505.xls`: `3096 ms -> 1387 ms`
    - `019088.xls`: `1871 ms -> 1077 ms`
    - `019089.xls`: `1824 ms -> 1059 ms`
- Interpretation:
  - this succeeds where the earlier short-digit guard experiments failed because it does not change
    the coverage semantics for short numeric visible hits
  - it replaces the expensive generic substring operation with an exact precomputed answer only for
    the specific query family that the retained `.xls` hotspot still uses
- Decision:
  - Retained.

## 2026-07-06 rejected: skip exact precheck for the `006087`-style short-text `.xls` shape

- Idea:
  - keep the retained generic exact precheck, but add a very narrow `.xls`-only bypass when the
    workload looks like `006087.xls`:
    - at least `30000` normalized source lines
    - total source text at most `512 KiB`
    - the usual `.xls`-only backfill flags already active
  - in that shape, skip `markdownBackfillExactSet(...)` and
    `markdownBackfillExactLinesCoveredWithExact(...)` entirely, then go straight to the shared
    `coverage + containment` path
- Why:
  - fresh instrumentation showed that the retained heavy `.xls` samples were not all alike:
    - `008055.xls` exact precheck succeeded completely with `1181169` exact hits
    - `016161.xls` exact precheck succeeded completely with `461603` exact hits
    - `006087.xls` got `36339` exact hits, then missed only at the very end on short numeric
      fragment `61047`
  - that made `006087.xls` look like the only serious candidate for skipping the exact stage
- Result:
  - focused hotspot view was mixed:
    - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `760822200 ns/op -> 577143000 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `2104182200 ns/op -> 2109205900 ns/op`
    - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `814307500 ns/op -> 886576100 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `495409100 ns/op -> 682640500 ns/op`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1489791600 ns/op -> 1550357600 ns/op`
    - `BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls`: `22748600 ns/op -> 26933400 ns/op`
  - the decisive repeat-aware `.xls6` rerun lost to the retained baseline:
    - total: `10461 ms -> 11758 ms`
    - `006087.xls`: `1351 ms -> 1652 ms`
    - `008055.xls`: `3922 ms -> 4826 ms`
    - `016161.xls`: `1665 ms -> 1654 ms`
    - `002505.xls`: `1387 ms -> 1470 ms`
    - `019088.xls`: `1077 ms -> 1063 ms`
    - `019089.xls`: `1059 ms -> 1093 ms`
- Interpretation:
  - even when the exact stage looks redundant on one isolated sample, the retained integrated `.xls`
    path still benefits from keeping that structure intact
  - the short-digit visible substring optimization already removed the real remaining replay cost,
    so cutting the exact stage at that point mostly disrupts the broader retained balance
- Decision:
  - Reverted.

## 2026-07-06 retained: cheaper byte checks in `simpleInlineWorksheetCandidate(...)` for large inlineStr `.xlsx`

- Idea:
  - keep the retained single-pass `simpleInlineWorksheetCandidate(...)` scan, but replace its hot
    `bytes.HasPrefix(...)` checks with direct fixed-width byte comparisons
  - preserve the same rejection markers and the same `foundInlineStr` success condition
- Why:
  - a fresh `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx` profile still
    showed `simpleInlineWorksheetCandidate(...)` as the dominant frame:
    - before this change: `1.16s flat`, `2.32s cum`
    - `bytes.HasPrefix` / `bytes.Equal` / `memeqbody` together were still a large fraction of the
      total scan cost
  - the retained one-pass scan had already removed the second pass; the next obvious cost was the
    repeated generic prefix machinery on every candidate byte
- Validation:
  - focused `.xlsx` hotspot benchmarks with `-benchtime=1x`
  - repeat-aware focused rerun on the large-inline pair with
    `cmd/compatcheck -repeat 3`
  - CPU profile on
    `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx`
  - repository regression:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- Result:
  - focused baseline from the same turn before the change:
    - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `3814173000 ns/op`
    - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2513019300 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `3780165300 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1403364200 ns/op`
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1639419400 ns/op`
  - repeat-aware focused rerun after the change:
    - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`:
      `2726812800 / 3230328300 / 3328964500 ns/op`
    - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
      `2588378900 / 2622103500 / 2620755000 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`:
      `3053558500 / 3374688400 / 2709811100 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
      `1248280500 / 1232240400 / 1217979500 ns/op`
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`:
      `1252283200 / 1216756100 / 1238279800 ns/op`
  - repeat-aware focused pair rerun kept output parity and improved wall time:
    - previous retained pair report:
      `testdata\web-samples\reports\perf-exp-ai-assistant-worksheetvalue-xlsx-pair.json`
      - total `.xlsx`: `7586 ms`
      - `00012389.xlsx`: `2535 ms`, `textBytes=11148655`, `images=0`
      - `testRecordSizeExceeded.xlsx`: `5051 ms`, `textBytes=181333430`, `images=0`
    - current byte-check rerun:
      `testdata\web-samples\reports\perf-exp-ai-assistant-candidate-bytechecks-xlsx-pair.json`
      - total `.xlsx`: `5765 ms`
      - `00012389.xlsx`: `2463 ms`, `textBytes=11148655`, `images=0`
      - `testRecordSizeExceeded.xlsx`: `3302 ms`, `textBytes=181333430`, `images=0`
  - profile after the change still shows the same frame on top, but materially smaller:
    - `simpleInlineWorksheetCandidate(...)`: `0.94s flat`, `0.94s cum`
  - repository regression stayed green:
    - `ok officeread 119.878s`
- Interpretation:
  - the retained single-pass scan was still correct, but it was spending too much time in generic
    prefix helpers
  - replacing those with direct byte checks gives a clear win on the large inlineStr sheets while
    keeping the extracted text and image counts unchanged on the repeat-aware pair rerun
  - `00012389.xlsx` extraction-level timing moved slightly in the focused benchmark, but the
    repeat-aware compat rerun and worksheet-level benchmarks stayed favorable overall
- Decision:
  - Retained.

## 2026-07-06 rejected: skip the second trim in `appendWorksheetValue(...)`

- Idea:
  - narrow the worksheet-only path by assuming `cleanText(...)` already returns trimmed output
  - after `appendWorksheetValue(...)` calls `cleanText(value)`, write the result directly with
    `appendTrimmedTextBlock(...)` instead of going through `appendCleanedTextBlock(...)`
- Why:
  - the latest `.xlsx` profile after the retained candidate-scan win still showed
    `appendWorksheetValue(...)` and `cleanText(...)` as the next cluster behind
    `simpleInlineWorksheetCandidate(...)`
  - this looked like a cheap place to remove one more `strings.TrimSpace(...)` from the hottest
    simple-inline worksheet path without touching the broader text-cleaning pipeline
- Validation:
  - focused `.xlsx` hotspot benchmarks with `-benchtime=1x`
  - repeat-aware focused pair rerun with
    `cmd/compatcheck -repeat 3`
  - repository regression after revert:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- Result:
  - the focused pair rerun looked superficially promising and kept output parity:
    - previous retained pair report:
      `perf-exp-ai-assistant-candidate-bytechecks-xlsx-pair.json`
      - total `.xlsx`: `5765 ms`
      - `00012389.xlsx`: `2463 ms`, `textBytes=11148655`, `images=0`
      - `testRecordSizeExceeded.xlsx`: `3302 ms`, `textBytes=181333430`, `images=0`
    - trim-skip rerun:
      `perf-exp-ai-assistant-worksheet-trimskip-xlsx-pair.json`
      - `00012389.xlsx`: `2177 ms`, `textBytes=11148655`, `images=0`
      - `testRecordSizeExceeded.xlsx`: `3238 ms`, `textBytes=181333430`, `images=0`
      - total `.xlsx`: `5415 ms`
  - but the decisive repeat-aware focused benchmarks regressed on the real retained hotspot:
    - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`:
      `2735872500 / 2929788700 / 3210416300 ns/op`
    - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
      `2723029000 / 2643780000 / 2673465600 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`:
      `3043613900 / 3037627700 / 3182032000 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
      `1594783500 / 1658658300 / 1983508600 ns/op`
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`:
      `2024724800 / 1979255100 / 1993959500 ns/op`
  - repository regression stayed green after reverting the experiment:
    - `ok officeread 136.639s`
- Interpretation:
  - even though the repeat-aware pair rerun improved, the more targeted repeat-aware hotspot
    benchmarks showed a clear regression, especially on
    `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`
  - this is another case where a small-looking helper shortcut does not survive the real retained
    benchmark harness
  - the temporary pair improvement was too weak and too indirect to justify keeping the change
- Decision:
  - Reverted.

## 2026-07-06 rejected: first-byte gating for `maybeControlFragmentText(...)`

- Idea:
  - narrow the expensive control-fragment helpers behind a first-byte gate inside
    `maybeControlFragmentText(...)`
  - only call `looksLikeOLEIdentifierFragment(...)`, `looksLikeOLEWrapperStreamName(...)`, and
    `looksLikeOOXMLMarkupNameFragment(...)` for leading-byte families that looked plausible from the
    profile, while still keeping the existing keyword switch
- Why:
  - after the retained `.xlsx` candidate-scan win, a fresh simple-inline profile showed
    `cleanTextFastPathControlFragment(...)` and `maybeControlFragmentText(...)` were still costing
    meaningful time on the large inlineStr path
  - the helpers were being called for every 5-80 byte candidate, even though most worksheet values
    were ordinary text that would never match OLE / OOXML control fragments
- Validation:
  - focused `.xlsx` hotspot benchmarks with `-benchtime=1x`
  - repeat-aware focused pair rerun with
    `cmd/compatcheck -repeat 3`
  - CPU profile on
    `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx`
  - repository regression:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- Result:
  - the focused benchmark and pair rerun both looked promising:
    - one-shot focused view:
      - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `2532887600 ns/op`
      - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2147824100 ns/op`
      - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `2993073600 ns/op`
      - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1397696300 ns/op`
      - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1126569100 ns/op`
    - repeat-aware focused rerun:
      - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`:
        `2464512900 / 2609855300 / 2860446000 ns/op`
      - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
        `2397538500 / 2455047400 / 2429942900 ns/op`
      - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`:
        `2818320200 / 2689491300 / 2732958700 ns/op`
      - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
        `1495584300 / 1461816400 / 1525629500 ns/op`
      - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`:
        `1370466300 / 1333364900 / 1392958600 ns/op`
    - repeat-aware pair rerun kept output parity and improved total time:
      - `perf-exp-ai-assistant-controlfragment-gate-xlsx-pair.json`
      - `00012389.xlsx`: `1978 ms`, `textBytes=11148655`, `images=0`
      - `testRecordSizeExceeded.xlsx`: `3097 ms`, `textBytes=181333430`, `images=0`
      - total `.xlsx`: `5075 ms`
    - the profile also looked materially better:
      - `cleanTextFastPathControlFragment(...)`: about `0.07s cum`
      - `maybeControlFragmentText(...)`: about `0.06s cum`
  - but the experiment was semantically wrong and failed repository tests before revert:
    - `TestOLEClassFragmentsAreFilteredWithoutDroppingProse`:
      expected `CompObj` to be dropped, got `CompObj`
    - `TestOLEIdentifierFragmentsAreFilteredConservatively`:
      expected `CLSID={00020906-0000-0000-C000-000000000046}` to be dropped, but it remained
    - failing run:
      - `FAIL officeread 129.224s`
  - after reverting the experiment, repository regression returned green:
    - `ok officeread (cached)`
- Interpretation:
  - the first-byte gate improved the hot benchmark path by skipping too much real semantic work
  - `CompObj` and `CLSID=...` are exactly the kind of OLE/control fragments that this filter is
    meant to catch, so the benchmark win is not admissible
  - this is a clean case where stronger correctness evidence overrides an otherwise attractive
    performance signal
- Decision:
  - Reverted.

## 2026-07-06 rejected: `<t>` tag fast path in `simpleInlineCellText(...)`

- Idea:
  - add a narrow fast path in `simpleInlineCellText(...)` for the common non-namespaced `<t>` tag
  - avoid calling `xmlStartTagNameIs(...)` when the opening tag is already in the trivial
    `<t...>` shape, and only fall back to the generic checker for namespaced or suspicious cases
- Why:
  - recent `.xlsx` simple-inline profiles still showed `simpleInlineCellText(...)` as a visible
    cost center, with `xmlStartTagNameIs(...)` itself taking a measurable slice
  - this looked like a low-risk parser micro-optimization that should have been semantics-neutral
- Validation:
  - focused `.xlsx` hotspot benchmarks with `-benchtime=1x`
  - repeat-aware focused pair rerun with
    `cmd/compatcheck -repeat 3`
  - repository regression:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- Result:
  - focused benchmark view regressed:
    - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `2742973800 ns/op`
    - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2688273900 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `3568474200 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1763604600 ns/op`
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1618000500 ns/op`
  - repeat-aware focused pair rerun also regressed while keeping output parity:
    - `perf-exp-ai-assistant-inlinecell-tagfast-xlsx-pair.json`
    - `00012389.xlsx`: `2201 ms`, `textBytes=11148655`, `images=0`
    - `testRecordSizeExceeded.xlsx`: `3510 ms`, `textBytes=181333430`, `images=0`
    - total `.xlsx`: `5711 ms`
  - repository regression itself stayed green both before and after revert:
    - `ok officeread 126.086s`
    - post-revert: `ok officeread (cached)`
- Interpretation:
  - even though the change was tiny and semantics-neutral, it moved the hot path in the wrong
    direction on both the focused benchmark and the repeat-aware pair rerun
  - this suggests the extra branch shape and fallback logic cost more than the saved generic tag
    check in the real retained workload
- Decision:
  - Reverted.

## 2026-07-06 rejected: byte-specialized `cellRefIndexesBytes(...)` in the simple-inline `.xlsx` path

- Idea:
  - add a `[]byte`-specialized cell reference parser for the simple-inline worksheet path
  - avoid converting `xmlAttrBytes(tag, "r")` to `string` before parsing the cell reference, and
    parse the byte slice directly instead
- Why:
  - recent simple-inline profiles still showed both `cellRefIndexes(...)` and
    `runtime.slicebytetostring` on the hot path
  - this looked like a low-risk way to remove an allocation-shaped conversion and keep the
    optimization tightly scoped to the retained `.xlsx` fast path
- Validation:
  - focused `.xlsx` hotspot benchmarks with `-benchtime=1x`
  - repeat-aware focused pair rerun with
    `cmd/compatcheck -repeat 3`
  - repository regression:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- Result:
  - focused benchmark view regressed:
    - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `2768071800 ns/op`
    - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2726203100 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `3483308200 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1749171500 ns/op`
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1786622000 ns/op`
  - repeat-aware focused pair rerun also regressed while keeping output parity:
    - `perf-exp-ai-assistant-cellref-bytes-xlsx-pair.json`
    - `00012389.xlsx`: `2599 ms`, `textBytes=11148655`, `images=0`
    - `testRecordSizeExceeded.xlsx`: `3559 ms`, `textBytes=181333430`, `images=0`
    - total `.xlsx`: `6158 ms`
  - repository regression stayed green:
    - `ok officeread 131.562s`
    - after revert: `ok officeread (cached)`
- Interpretation:
  - removing the `[]byte -> string` conversion looked attractive in isolation, but the specialized
    parser was slower in the real retained workload
  - the extra custom parsing path likely cost more than the conversion it was trying to avoid
- Decision:
  - Reverted.

## 2026-07-06 current hotspot split: `00012389.xlsx` and `testRecordSizeExceeded.xlsx` are now different `.xlsx` problems

- Observation:
  - a fresh current-baseline repro shows the two retained `.xlsx` hotspot files are no longer
    dominated by the same code path
- Evidence:
  - focused current baseline measurements:
    - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2201956800 ns/op`
    - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `2908301300 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1228603400 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `2372883900 ns/op`
  - profile for `00012389.xlsx`:
    - top cumulative frame is `appendWorksheetText` at about `1.60s cum`
    - the path is dominated by `encoding/xml.Decoder` work, XML tokenization, and the general
      worksheet reader
    - `simpleInlineWorksheetCandidate(...)` does not dominate this file
  - profile for `testRecordSizeExceeded.xlsx`:
    - `appendSimpleInlineWorksheetTextPrepared(...)` is still the core path at about `1.94s cum`
    - `simpleInlineWorksheetCandidate(...)` remains the top flat frame at about `0.45s`
    - `cleanTextFastPath(...)`, `maybeControlFragmentText(...)`, and `simpleInlineCellText(...)`
      remain material within that simple-inline path
- Interpretation:
  - `testRecordSizeExceeded.xlsx` is still the retained simple-inline giant-inlineStr problem
  - `00012389.xlsx` has effectively moved back to the general XML-decoder worksheet path
  - this means single micro-optimizations aimed at the simple-inline helpers should no longer be
    expected to help both hotspot files at once
  - the next profitable `.xlsx` work likely needs to branch:
    - simple-inline specific ideas for `testRecordSizeExceeded.xlsx`
    - XML-decoder / worksheet-reader ideas for `00012389.xlsx`
- Decision:
  - Keep as guidance for subsequent optimization rounds; no code change from this finding alone.

## 2026-07-06 rejected: remove duplicate `hiddenColumnCell(...)` check from `appendWorksheetText(...)`

- Idea:
  - in the general worksheet XML path, remove `hiddenColumnCell(cellRef, hiddenCols)` from
    `skipCell` once `cellCol` has already been computed
  - keep only `rowHidden || columnHidden(cellCol, hiddenCols)` for the common cell skip decision
- Why:
  - `appendWorksheetText(...)` already parses `cellRef`, derives `cellCol`, and then immediately
    repeats another column lookup via `hiddenColumnCell(...)`
  - after confirming that `00012389.xlsx` is now a general XML-decoder problem, this looked like a
    more structural duplicate check than the recent parser micro-optimizations
- Validation:
  - focused `00012389.xlsx` benchmarks with `-count=3`
  - isolated repeat-aware rerun:
    - `cmd/compatcheck -repeat 3 ... 00012389.xlsx`
  - repository regression after revert:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- Result:
  - isolated rerun looked slightly favorable:
    - `perf-exp-ai-assistant-00012389-no-hiddenColumnCell.json`
    - `00012389.xlsx`: `2122 ms`, `min=2048`, `max=2518`, `textBytes=11148655`, `images=0`
  - but the focused benchmark signal was mixed rather than convincingly better:
    - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
      `2121202300 / 2302127400 / 2340548000 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
      `1234948500 / 1284801800 / 1317203100 ns/op`
  - repository regression stayed green after revert:
    - `ok officeread (cached)`
- Interpretation:
  - the end-to-end one-file rerun improved slightly, but the focused benchmark harness did not
    produce a strong, repeat-aware win
  - given the weak and mixed evidence, this change does not meet the current retention bar
- Decision:
  - Reverted.

## 2026-07-06 current finding: markdown preparation is the dominant extra cost on `00012389.xlsx`

- Observation:
  - a direct benchmark comparison of the same worksheet path with and without markdown preparation
    shows that the general XML-decoder cost is no longer the full story for `00012389.xlsx`
- Evidence:
  - current repeat-aware worksheet benchmark with markdown preparation enabled:
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
      `1256466300 / 1242155900 / 1320314700 ns/op`
  - temporary isolated benchmark on the same sheet path with `appendWorksheetText(..., nil)`:
    - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`:
      `800303200 / 863024900 / 820619900 ns/op`
  - allocation delta on the same benchmark pair:
    - with markdown prep: about `562-563 MB`, `~9.94M allocs/op`
    - without markdown prep: about `456-457 MB`, `~8.37M allocs/op`
- Interpretation:
  - for `00012389.xlsx`, markdown preparation inside `appendWorksheetText(...)` is a first-order
    cost, not just a side effect
  - recent failed micro-optimizations in the XML-decoder loop likely looked weak because the bigger
    opportunity is now the worksheet-markdown preparation path itself
  - any future retained optimization for this file probably needs to reduce markdown row / cell
    preparation cost, or provide a semantically acceptable way to avoid paying it when structured
    markdown is not required
- Decision:
  - Record as guidance for the next rounds; temporary benchmark removed after measurement.

## 2026-07-06 rejected: case-fold prefix check shortcut in `prepareMarkdownTableCellValue(...)`

- Idea:
  - replace the per-cell `strings.ToLower(trimmed)` allocation in the RTF-prefix check with the
    existing `hasPrefixFold(...)` helper
  - keep the same semantics for the `{\rtf` / `\rtf` detection and leave the rest of markdown cell
    preparation unchanged
- Why:
  - after confirming that markdown preparation is a first-order cost on `00012389.xlsx`, this was a
    low-risk attempt to trim obvious allocation churn inside a hot helper on that path
- Validation:
  - focused worksheet benchmark on `00012389.xlsx` with `-count=5`
  - isolated repeat-aware rerun:
    - `cmd/compatcheck -repeat 5 ... 00012389.xlsx`
  - repository regression after revert:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- Result:
  - focused worksheet benchmark showed a lower allocation profile, but mixed timing:
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
      `1222076400 / 1208177300 / 1447019500 / 1441599500 / 1468807700 ns/op`
    - alloc profile improved versus the earlier current-baseline markdown-prep benchmark:
      - about `550-551 MB`, `~9.29M allocs/op`
      - earlier baseline was about `562-563 MB`, `~9.94M allocs/op`
  - isolated rerun did not improve convincingly:
    - `perf-exp-ai-assistant-00012389-rtf-foldcheck.json`
    - `00012389.xlsx`: `2255 ms`, `min=1990`, `max=2434`, `textBytes=11148655`, `images=0`
  - repository regression stayed green after revert:
    - `ok officeread (cached)`
- Interpretation:
  - the helper change did reduce allocations, but the throughput signal was too mixed and the
    isolated rerun did not beat the earlier retained current-baseline evidence strongly enough
  - by the current acceptance bar, this is not yet a convincing performance win
- Decision:
  - Reverted.

## 2026-07-06 rejected: reusable row buffer for `appendWorksheetText(...)` markdown rows

- Idea:
  - replace the per-row `map[int]string` used by `appendWorksheetText(...)` markdown preparation
    with a reusable fixed-width row buffer
  - avoid repeated map allocation and `mapassign_fast64` churn by reusing a `[]string` scratch row
    across worksheet rows
- Why:
  - after confirming that markdown preparation is a major cost on `00012389.xlsx`, a reusable row
    buffer looked like a more structural way to attack row-level markdown preparation than helper
    micro-tuning
  - the current profile still showed map and row-compaction costs inside the markdown path
- Validation:
  - focused worksheet benchmark on `00012389.xlsx` with `-count=5`
  - isolated repeat-aware rerun:
    - `cmd/compatcheck -repeat 5 ... 00012389.xlsx`
  - repository regression after revert:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- Result:
  - focused worksheet benchmark again showed a materially better allocation profile, but mixed wall
    time:
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
      `1204844100 / 1200706100 / 1529469600 / 1532947200 / 1597046600 ns/op`
    - alloc profile improved versus the current-baseline markdown-prep benchmark:
      - about `515-516 MB`, `~9.73M allocs/op`
      - earlier baseline was about `562-563 MB`, `~9.94M allocs/op`
  - isolated rerun still did not beat the earlier current-baseline evidence convincingly:
    - `perf-exp-ai-assistant-00012389-rowbuffer.json`
    - `00012389.xlsx`: `2241 ms`, `min=2050`, `max=2592`, `textBytes=11148655`, `images=0`
  - repository regression stayed green after revert:
    - `ok officeread (cached)`
- Interpretation:
  - the row-buffer idea clearly reduced allocation pressure, which suggests it is structurally in
    the right neighborhood
  - but the wall-time result was too unstable, and the isolated rerun still failed to beat the
    earlier `2122 ms` evidence strongly enough
  - under the current standard, this is still not a retainable optimization
- Decision:
  - Reverted.

## 2026-07-06 component split: `cleanMarkdownTableCellValue(...)` dominates the `00012389.xlsx` markdown-prep chain

- Observation:
  - a temporary component benchmark split for the `00012389.xlsx` worksheet markdown-preparation
    chain shows that `cleanMarkdownTableCellValue(...)` is much larger than the downstream
    `prepareMarkdownTableCellValue(...)` and row compaction steps
- Evidence:
  - temporary focused component benchmarks with `-benchtime=3x`:
    - `BenchmarkXLSXMarkdownCellClean00012389`: `14267 ns/op`, `13850 B/op`, `8 allocs/op`
    - `BenchmarkXLSXMarkdownCellPrepare00012389`: `566.7 ns/op`, `128 B/op`, `2 allocs/op`
    - `BenchmarkXLSXMarkdownRowCompact00012389`: `1267 ns/op`, `1024 B/op`, `1 allocs/op`
- Interpretation:
  - the markdown-preparation chain is not evenly distributed
  - `prepareMarkdownTableCellValue(...)` and `compactWorksheetMarkdownRow(...)` are both real
    costs, but they are an order of magnitude smaller than `cleanMarkdownTableCellValue(...)` on
    representative `00012389.xlsx` cell data
  - this strongly suggests the next retained optimization attempt for `00012389.xlsx` should target
    the trigger surface or internal helper work of `cleanMarkdownTableCellValue(...)`, rather than
    continuing to micro-tune row compaction or final formatting
- Decision:
  - Record as guidance for the next rounds; temporary benchmark removed after measurement.

## 2026-07-06 deeper split: `00012389.xlsx` markdown cells are overwhelmingly plain single-line text

- Observation:
  - a temporary analysis pass over the representative worksheet markdown-cell corpus from
    `00012389.xlsx` shows that almost every collected value is already plain single-line text after
    `cleanText(...)` and `stripInlineHiddenOfficeReferences(...)`
- Evidence:
  - temporary analysis test over the collected worksheet markdown values:
    - `values=4102`
    - `singleLine=4071`
    - `multiLine=31`
    - `markerHits=0`
    - `discardHits=23`
    - `controlHits=0`
    - `lineDiscardHits=54`
    - `truncatedHits=0`
  - temporary stage split benchmark with `-benchtime=10000x`:
    - `cleanText`: `201.2 ns/op`
    - `stripInlineHiddenOfficeReferences`: `65.84 ns/op`
    - `maybeDiscardableHiddenOfficeText`: `137.5 ns/op`
    - `maybeControlFragmentText`: `84.87 ns/op`
- Interpretation:
  - the representative `00012389.xlsx` markdown-cell workload is not dominated by inline hidden
    reference stripping or control-fragment matches; those triggers are very rare on this corpus
  - future `00012389.xlsx` work should treat the common path as plain single-line text first, and
    only spend complexity on a retained optimization if it demonstrably helps that dominant case
- Decision:
  - Keep as profiling guidance only; temporary analysis code removed after measurement.

## 2026-07-06 rejected: prepared-row compaction fast path for `.xlsx` worksheet markdown rows

- Idea:
  - add a specialized `compactPreparedWorksheetMarkdownRow(...)` helper for worksheet markdown rows
    that skips the repeated `strings.TrimSpace(...)` check inside `compactWorksheetMarkdownRow(...)`
  - tighten markdown row population so the specialized path only sees already-prepared cell values
- Why:
  - the temporary component split now shows row compaction is nontrivial in the markdown-prep chain
  - for `.xlsx` worksheet markdown rows, the row maps are populated from
    `prepareMarkdownTableCellValue(...)`, so a trim-free compaction path looked like a conservative
    way to shave helper work without changing extraction rules
- Validation:
  - focused row-compaction microbenchmark:
    - `BenchmarkXLSXMarkdownRowCompact00012389 -count=5`
  - focused worksheet benchmark:
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx -count=5`
- Result:
  - the isolated row-compaction microbenchmark was noisy:
    - `1125 / 1283 / 1745 / 1690 / 1691 ns/op`
  - the real worksheet hotspot signal did not improve convincingly:
    - `1245719200 / 1215814200 / 1794226800 / 1747040300 / 1625623500 ns/op`
    - this is not clearly better than the earlier current-baseline signal around
      `~1.23-1.32 s/op`
- Interpretation:
  - the specialized helper may remove a little redundant work locally, but it did not translate
    into a stable end-to-end worksheet win on the retained hotspot workload
- Decision:
  - Reverted.

## 2026-07-06 rejected: `cleanTextFastPath(...)` control-fragment check reordering

- Idea:
  - keep `cleanTextFastPath(...)` semantics unchanged, but move the
    `cleanTextFastPathControlFragment(...)` call after the cheap ASCII / marker screening loop
  - the goal was to avoid paying the control-fragment check up front for strings that already fail
    fast-path eligibility for unrelated reasons
- Why:
  - after confirming that `00012389.xlsx` markdown-cell values are mostly plain single-line text,
    a semantics-preserving fast-path reorder was a low-risk candidate to test before touching any
    filtering rules
- Validation:
  - temporary equivalence check on the representative worksheet markdown values:
    - `TestCleanTextFastPathReorderedMatches00012389`
  - focused microbenchmark:
    - `BenchmarkXLSXMarkdownCleanFastPathVariants00012389 -count=5`
- Result:
  - the temporary equivalence check passed on the representative corpus
  - but the reordered variant was slower:
    - current: `117.8 / 133.6 / 152.4 / 133.8 / 169.7 ns/op`
    - reordered: `165.5 / 164.0 / 161.7 / 160.6 / 151.7 ns/op`
- Interpretation:
  - the existing check order is already favorable on the current representative markdown-cell
    workload
- Decision:
  - Not applied.

## 2026-07-06 rejected: single-line `.xlsx` markdown common-path hidden-check gating

- Idea:
  - keep `cleanMarkdownTableCellValue(...)` semantics the same for the representative
    `00012389.xlsx` corpus, but restructure the single-line path so it:
    - checks `maybeControlFragmentText(...)` first
    - only runs `maybeDiscardableHiddenOfficeText(...)` when the cleaned cell still carries
      hidden-office marker characters such as `=`, `:`, `/`, `\`, `%`, `<`, `>`, `#`, or `rid`
- Why:
  - the latest component split showed `00012389.xlsx` markdown cells are overwhelmingly plain
    single-line text, with very few hidden-reference hits
  - a temporary equivalence check on the representative corpus passed, and the isolated
    cell-clean microbenchmark improved materially
- Validation:
  - representative-corpus equivalence check before applying:
    - temporary `TestCleanMarkdownTableCellValueAltMatches00012389`
  - isolated cell-clean microbenchmark before applying:
    - current: `472.5 / 508.4 / 484.4 / 481.8 / 474.5 ns/op`
    - candidate: `385.1 / 388.8 / 408.6 / 394.9 / 377.7 ns/op`
  - applied candidate, then reran:
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx -count=5`
    - `cmd/compatcheck -repeat 5 ... 00012389.xlsx`
- Result:
  - focused worksheet benchmark remained mixed:
    - baseline before apply:
      `1243286400 / 1204917700 / 1532021200 / 1565506400 / 1538926400 ns/op`
    - candidate after apply:
      `1192275100 / 1345777500 / 1621469000 / 1593249600 / 1648980400 ns/op`
  - repeat-aware single-file rerun regressed:
    - baseline before apply:
      `ms=2150 min=2053 max=2565 runs=[2060 2150 2526 2565 2053]`
    - candidate after apply:
      `ms=2484 min=2012 max=3076 runs=[2012 2484 3076 3006 2423]`
  - output parity stayed intact on the single-file rerun:
    - `text=11148655`
    - `images=0`
- Interpretation:
  - this is another case where a clean microbenchmark win on the common path does not survive the
    integrated worksheet workload strongly enough
  - the marker-gated hidden-office check may still be locally sound, but under the current
    retention bar it is not a convincing end-to-end optimization
- Decision:
  - Reverted.

## 2026-07-06 rejected: multi-`bytes.Contains` candidate scan for `simpleInlineWorksheetCandidate(...)`

- Idea:
  - replace the current single-pass byte loop in `simpleInlineWorksheetCandidate(...)` with a
    candidate detector built from `bytes.Contains(...)` checks:
    - require one positive hit for `t="inlineStr"`
    - reject on any forbidden marker such as `<!DOCTYPE`, `<f`, `<v`, `<cols`, `<rPh`,
      `<phoneticPr`, `t="s"`, `t="b"`, `t="str"`, `hidden`, `hyperlink`, `dataValidation`,
      `Header`, `Footer`, or `ht=`
- Why:
  - the latest `testRecordSizeExceeded.xlsx` profile shows `simpleInlineWorksheetCandidate(...)`
    still dominates the simple-inline path with about `37.4%` flat CPU
  - a `bytes.Contains(...)` strategy looked attractive because the heavy positive sample is a huge
    worksheet and the runtime may search large byte slices faster than a Go-level byte-by-byte loop
- Validation:
  - temporary parity check on the two hotspot worksheets:
    - positive: `testRecordSizeExceeded.xlsx`
    - negative: `00012389.xlsx`
  - temporary focused benchmark:
    - `BenchmarkXLSXSimpleInlineWorksheetCandidateVariants -count=3`
- Result:
  - parity matched on the hotspot positive/negative worksheets
  - but the performance result was dramatically worse:
    - positive `testRecordSizeExceeded.xlsx`:
      - current: `451517533 / 456994833 / 569939900 ns/op`
      - candidate: `1861008000 / 1917337600 / 1729359400 ns/op`
    - negative `00012389.xlsx`:
      - current: `1325 / 1217 / 1216 ns/op`
      - candidate: `5762842 / 5753632 / 5738494 ns/op`
- Interpretation:
  - repeated full-slice `bytes.Contains(...)` scans are much more expensive here than the current
    single-pass loop
  - on the negative sample, the existing implementation benefits heavily from an early reject,
    which the multi-pass candidate loses completely
- Decision:
  - Not applied.

## 2026-07-06 rejected: sparse interesting-byte scan for `simpleInlineWorksheetCandidate(...)`

- Idea:
  - keep `simpleInlineWorksheetCandidate(...)` semantics the same, but replace the current
    byte-by-byte loop with a sparse scan that jumps between only the interesting lead bytes:
    `<`, `t`, `h`, `d`, `H`, `F`, and whitespace
  - use `bytes.IndexAny(...)` to skip long runs of irrelevant bytes in the huge
    `testRecordSizeExceeded.xlsx` worksheet
- Why:
  - after rejecting the multi-`bytes.Contains(...)` approach, the next plausible direction was to
    preserve the existing single-pass logic while reducing the cost of visiting every byte
- Validation:
  - temporary parity check on the hotspot positive/negative worksheets:
    - positive: `testRecordSizeExceeded.xlsx`
    - negative: `00012389.xlsx`
  - temporary focused benchmark:
    - `BenchmarkXLSXSimpleInlineWorksheetCandidateSparseVariants -count=3`
- Result:
  - parity matched on the hotspot positive/negative worksheets
  - but the sparse scan was still clearly slower:
    - positive `testRecordSizeExceeded.xlsx`:
      - current: `444585333 / 530123050 / 506335700 ns/op`
      - sparse: `1108017700 / 1101259800 / 1100491400 ns/op`
    - negative `00012389.xlsx`:
      - current: `1409 / 1210 / 1299 ns/op`
      - sparse: `2126 / 2133 / 2055 ns/op`
- Interpretation:
  - `bytes.IndexAny(...)` does not pay for itself here; the cost of repeated search restarts is
    still worse than the existing straight-line loop
  - the current implementation remains the better tradeoff on both the hotspot positive and
    early-reject negative samples
- Decision:
  - Not applied.

## 2026-07-06 rejected: `cleanTextFastPath(...)` reorder on `testRecordSizeExceeded.xlsx` simple-inline values

- Idea:
  - revisit the `cleanTextFastPath(...)` check reordering on the
    `testRecordSizeExceeded.xlsx` simple-inline workload specifically
  - keep semantics unchanged, but defer `cleanTextFastPathControlFragment(...)` until after the
    cheap ASCII / marker screening loop
- Why:
  - after amortizing the simple-inline benchmark setup noise with `-benchtime=3x`, the positive
    simple-inline text-only profile still showed `cleanTextFastPath(...)` near `10%` flat CPU
  - the representative `00012389.xlsx` result had already rejected this reorder for the general
    worksheet path, but the simple-inline value distribution is different enough to test separately
- Validation:
  - temporary representative parity check on collected simple-inline cell texts from
    `testRecordSizeExceeded.xlsx`
  - temporary focused microbenchmark:
    - `BenchmarkXLSXCleanTextFastPathVariantsTestRecord -count=5`
- Result:
  - parity matched on the collected simple-inline values
  - performance was too mixed to justify an integrated attempt:
    - current: `208.6 / 224.1 / 228.1 / 239.0 / 250.5 ns/op`
    - reordered: `219.9 / 217.5 / 229.6 / 249.6 / 263.6 ns/op`
- Interpretation:
  - unlike the earlier representative `00012389.xlsx` result, this workload does not show a clean
    regression, but it also does not show a stable enough improvement to warrant touching the real
    code path
- Decision:
  - Not applied.

## 2026-07-06 component split: `testRecordSizeExceeded.xlsx` simple-inline values are uniformly tiny fast-path strings

- Observation:
  - a temporary component split over collected simple-inline cell texts from
    `testRecordSizeExceeded.xlsx` shows an unusually uniform workload:
    - all collected cell texts hit `cleanTextFastPath(...)`
    - none of them trigger the large-value de-duplication map in `appendWorksheetValue(...)`
- Evidence:
  - temporary analysis on collected simple-inline cells:
    - `cells=3000000`
    - `large=0`
    - `fastPath=3000000`
    - `emptyAfterClean=0`
  - temporary component benchmark on the same collected cells:
    - `simpleInlineCellText`: about `72-106 ns/op`
    - `cellRefParse`: about `38-48 ns/op`
    - `appendWorksheetValue`: about `263-315 ns/op`
    - `cleanText`: about `222-342 ns/op`
    - `appendCleanedTextBlock`: about `36-52 ns/op`
- Interpretation:
  - on the current `testRecordSizeExceeded.xlsx` simple-inline workload, `appendWorksheetValue(...)`
    is dominated by `cleanText(...)`, and specifically by its fast-path common case
  - the large-value de-duplication branch is irrelevant for this workload
  - future retained work on this hotspot should target the fast-path common case itself, not the
    large-value map or cell-ref parsing
- Decision:
  - Record as guidance for the next rounds; temporary analysis code removed after measurement.

## 2026-07-06 rejected: direct fast-path append inside `appendWorksheetValue(...)`

- Idea:
  - in `appendWorksheetValue(...)`, detect `cleanTextFastPath(...)` explicitly and, on success,
    append the already-trimmed value directly with `appendTrimmedTextBlock(...)`
  - avoid routing the common fast-path case through `cleanText(...)` and then a second
    `appendCleanedTextBlock(...)` trim
- Why:
  - the temporary component split for `testRecordSizeExceeded.xlsx` showed:
    - `cleanTextFastPath(...)` hits on every collected simple-inline cell text
    - `appendWorksheetValue(...)` is much heavier than `simpleInlineCellText(...)` or
      `cellRefParse(...)`
- Validation:
  - temporary representative parity check on collected simple-inline cell texts
  - same-tree component benchmark comparing:
    - current `appendWorksheetValue(...)`
    - previous implementation helper
  - integrated validation after applying:
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx -count=5`
    - `cmd/compatcheck -repeat 5 ... testRecordSizeExceeded.xlsx`
- Result:
  - representative parity passed
  - integrated signals after applying looked plausible:
    - text-only hotspot after apply:
      `1317174900 / 1335036500 / 1275913600 / 1269446900 / 1288140700 ns/op`
    - repeat-aware single-file rerun after apply:
      `ms=2739 min=2557 max=3010`, `text=181333430`, `images=0`
  - but the decisive same-tree component comparison was not stable enough:
    - current:
      `297.0 / 271.2 / 263.0 ns/op`
    - previous:
      `260.1 / 265.3 / 289.1 ns/op`
- Interpretation:
  - this candidate is in the right neighborhood, but the local before/after signal is too mixed to
    justify retaining the code under the current standard
  - given the earlier history of micro-wins failing to survive broader validation, this one should
    not be kept yet
- Decision:
  - Reverted.

## 2026-07-06 rejected: merged single-loop `cleanTextFastPath(...)`

- Idea:
  - replace the current staged `cleanTextFastPath(...)` checks with a merged single-loop version
    that:
    - folds `strings.ContainsAny(...)` character rejection into the main byte loop
    - tracks a cheap `rid` hint during the same loop
    - defers the actual `containsRIDFold(...)` call until the hint is seen
    - keeps `cleanTextFastPathControlFragment(...)` at the end
- Why:
  - the latest simple-inline component split showed `testRecordSizeExceeded.xlsx` is a
    `100% cleanTextFastPath(...)` workload
  - a temporary representative microbenchmark looked promising enough to justify an integrated
    attempt:
    - `testRecord` representative values were mixed but not clearly worse
    - `00012389` representative values improved clearly:
      - current: `99.01 / 111.9 / 120.3 / 104.8 / 106.7 ns/op`
      - merged: `84.23 / 89.63 / 93.80 / 92.25 / 93.95 ns/op`
- Validation:
  - temporary parity check on representative values from:
    - `testRecordSizeExceeded.xlsx`
    - `00012389.xlsx`
  - integrated validation after applying:
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx -count=5`
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx -count=5`
    - `cmd/compatcheck -repeat 5 ... 00012389.xlsx`
    - `cmd/compatcheck -repeat 5 ... testRecordSizeExceeded.xlsx`
- Result:
  - representative parity passed
  - but integrated signals regressed:
    - `00012389.xlsx` worksheet hotspot:
      `1443830000 / 1555045300 / 1979124700 / 2045242800 / 2065070300 ns/op`
    - `testRecordSizeExceeded.xlsx` simple-inline text-only hotspot:
      `1317224900 / 1947406800 / 1818922200 / 1852340000 / 1782194700 ns/op`
    - `00012389.xlsx` repeat-aware rerun:
      `ms=3238 min=2461 max=3346`, `text=11148655`, `images=0`
    - `testRecordSizeExceeded.xlsx` repeat-aware rerun:
      `ms=3887 min=2984 max=3985`, `text=181333430`, `images=0`
- Interpretation:
  - this is another case where the local fast-path benchmark looked encouraging, but the real
    end-to-end extract paths got worse
  - under the current retention bar, the merged-loop idea is not keepable
- Decision:
  - Reverted.

## 2026-07-06 component split: `testRecordSizeExceeded.xlsx` simple-inline markdown cells are also uniform plain-text inputs

- Observation:
  - a temporary markdown-component split over collected simple-inline cell texts from
    `testRecordSizeExceeded.xlsx` shows the markdown branch is just as uniform as the text branch:
    - every collected value survives `cleanMarkdownTableCellValue(...)`
    - no collected value is multiline
    - no collected value looks like a control fragment after trimming
- Evidence:
  - temporary analysis on collected simple-inline markdown values:
    - `values=3000000`
    - `cleaned=3000000`
    - `multiline=0`
    - `controlLike=0`
  - temporary component benchmark:
    - `cleanMarkdownTableCellValue`: about `637-1073 ns/op`
    - `prepareMarkdownTableCellValue`: about `247-266 ns/op`
- Interpretation:
  - on the current `testRecordSizeExceeded.xlsx` simple-inline markdown workload,
    `cleanMarkdownTableCellValue(...)` is the dominant markdown-stage cost
  - because the inputs are uniformly single-line plain text with no control-like cases, any future
    retained optimization should target the plain-text common path inside
    `cleanMarkdownTableCellValue(...)`, not the multiline branch or downstream markdown formatting
- Decision:
  - Record as guidance for the next rounds; temporary benchmark removed after measurement.

## 2026-07-06 no-op probe: `testRecordSizeExceeded.xlsx` markdown plain-path has no `>80` value window to exploit

- Observation:
  - a temporary probe for the single-line plain-text branch of
    `cleanMarkdownTableCellValue(...)` on `testRecordSizeExceeded.xlsx` showed that the natural
    `maybeControlFragmentText(...)` length escape hatch (`len > 80`) has no room to help here
- Evidence:
  - temporary representative analysis:
    - `values=3000000`
    - `single=3000000`
    - `over80=0`
    - `over80NoMarkers=0`
  - temporary parity check for a candidate branch guarded by `len(value) > 80` passed, but it was
    effectively a no-op on this workload
  - temporary microbenchmark looked better numerically:
    - current: `810.0 / 827.7 / 762.5 / 762.4 / 889.2 ns/op`
    - candidate: `662.3 / 764.2 / 651.4 / 649.7 / 649.2 ns/op`
    - however, because the candidate's extra branch never actually triggered on the representative
      values, this does not establish a real causal optimization opportunity
- Interpretation:
  - this rules out the easiest length-based plain-text shortcut for the current `testRecord`
    simple-inline markdown workload
  - future markdown common-path work on this file needs a different discriminator than
    `maybeControlFragmentText(...)` length
- Decision:
  - Not applied; record as narrowing evidence only.

## 2026-07-06 rejected: no-marker single-line fast return in `cleanMarkdownTableCellValue(...)`

- Idea:
  - for the single-line branch of `cleanMarkdownTableCellValue(...)`, return early when the
    cleaned value has no marker characters from the hidden/control families:
    - no `/ \ % [ ] ( ) < > # = :`
    - no `rid` fragment
  - this skips both `maybeDiscardableHiddenOfficeText(...)` and `maybeControlFragmentText(...)`
    for plain single-line values
- Why:
  - a temporary distribution probe on `testRecordSizeExceeded.xlsx` showed the representative
    simple-inline markdown values are uniformly single-line and uniformly marker-free:
    - `values=3000000`
    - `single=3000000`
    - `noMarkers=3000000`
  - representative parity also held on the sampled `00012389.xlsx` worksheet values
  - temporary microbenchmark on the `testRecord` representative values looked strong:
    - current: `606.5 / 880.4 / 803.5 / 768.2 / 848.3 ns/op`
    - candidate: `429.0 / 482.4 / 442.7 / 463.2 / 454.4 ns/op`
- Validation:
  - representative parity on:
    - `testRecordSizeExceeded.xlsx`
    - `00012389.xlsx`
  - integrated validation after applying:
    - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx -count=5`
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx -count=5`
    - `cmd/compatcheck -repeat 5 ... testRecordSizeExceeded.xlsx`
    - `cmd/compatcheck -repeat 5 ... 00012389.xlsx`
- Result:
  - representative parity passed
  - but integrated signals regressed or stayed too weak:
    - `testRecordSizeExceeded.xlsx` simple-inline markdown hotspot:
      `2739514200 / 3057031300 / 2876826900 / 4471219700 / 2620501000 ns/op`
    - `00012389.xlsx` worksheet hotspot:
      `1525336200 / 1965797100 / 1945455300 / 1892368200 / 1888481900 ns/op`
    - `testRecordSizeExceeded.xlsx` repeat-aware rerun:
      `ms=3830 min=3462 max=5269`, `text=181333430`, `images=0`
    - `00012389.xlsx` repeat-aware rerun:
      `ms=3141 min=2426 max=3339`, `text=11148655`, `images=0`
- Interpretation:
  - the local markdown-cell win did not survive the integrated extraction workloads strongly enough
  - under the current retention bar, this candidate is not keepable
- Decision:
  - Reverted.

## 2026-07-06 current simple-inline markdown hotspot split for `testRecordSizeExceeded.xlsx`

- Observation:
  - a fresh `-benchtime=3x` profile of
    `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx` confirms that the markdown
    branch remains a combined problem, not a single helper problem
- Evidence:
  - focused benchmark:
    - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`: `2329455067 ns/op`
  - profile highlights:
    - `appendSimpleInlineWorksheetTextPrepared(...)`: `75.21%` cumulative
    - `appendWorksheetValue(...)`: `24.55%` cumulative
    - `cleanMarkdownTableCellValue(...)`: `14.96%` cumulative
    - `prepareMarkdownTableCellValue(...)`: `6.86%` cumulative
    - `simpleInlineCellText(...)`: `12.15%` cumulative
  - note:
    - `simpleInlineWorksheetCandidate(...)` still shows `15.54%` flat in this benchmark profile,
      but that frame is partly benchmark-preparation pollution because the benchmark locates and
      verifies the candidate worksheet before timed iterations begin
- Interpretation:
  - the retained next-step focus for `testRecord` markdown work should stay on the actual timed
    pipeline inside `appendSimpleInlineWorksheetTextPrepared(...)`
  - the most credible remaining structural costs are still:
    - `appendWorksheetValue(...)` text path
    - `cleanMarkdownTableCellValue(...)` single-line common path
    - `prepareMarkdownTableCellValue(...)`
- Decision:
  - Record as current hotspot guidance.

## 2026-07-06 rejected: plain-text direct return in `prepareMarkdownTableCellValue(...)`

- Idea:
  - for single-line inputs to `prepareMarkdownTableCellValue(...)`, return
    `normalizeMarkdownTextLine(trimmed)` directly when the cleaned value contains no:
    - backslash
    - pipe
    - newline
    - RTF prefix
  - skip the two `strings.ReplaceAll(...)` calls entirely for that plain-text case
- Why:
  - a temporary representative shape check on `testRecordSizeExceeded.xlsx` showed the markdown
    prepare-stage inputs are perfectly uniform for this discriminator:
    - `inputs=3000000`
    - `noSpecial=3000000`
    - `hasSlash=0`
    - `hasPipe=0`
    - `hasNewline=0`
    - `hasRTF=0`
  - representative parity also held on sampled `00012389.xlsx` prepare-stage values
  - representative microbenchmark on `testRecord` looked modestly favorable:
    - current: `324.0 / 272.2 / 253.9 / 249.4 / 261.0 ns/op`
    - candidate: `244.0 / 248.7 / 297.5 / 242.8 / 240.4 ns/op`
- Validation:
  - representative parity on:
    - `testRecordSizeExceeded.xlsx`
    - `00012389.xlsx`
  - integrated validation after applying:
    - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx -count=5`
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx -count=5`
    - `cmd/compatcheck -repeat 5 ... testRecordSizeExceeded.xlsx`
    - `cmd/compatcheck -repeat 5 ... 00012389.xlsx`
- Result:
  - representative parity passed
  - but integrated signals regressed:
    - `testRecordSizeExceeded.xlsx` simple-inline markdown hotspot:
      `4822266600 / 3355737000 / 3236953000 / 2938747300 / 2391371100 ns/op`
    - `00012389.xlsx` worksheet hotspot:
      `2011261400 / 3265431300 / 1996105300 / 2180943500 / 2158418600 ns/op`
    - `testRecordSizeExceeded.xlsx` repeat-aware rerun:
      `ms=4246 min=3563 max=5623`, `text=181333430`, `images=0`
    - `00012389.xlsx` repeat-aware rerun:
      `ms=3412 min=2758 max=5034`, `text=11148655`, `images=0`
- Interpretation:
  - despite the perfectly uniform representative input shape, this prepare-stage shortcut still did
    not survive integrated extraction validation
  - under the current retention bar, it is not keepable
- Decision:
  - Reverted.

## 2026-07-06 rejected: row-slice markdown row builder for `appendSimpleInlineWorksheetTextPrepared(...)`

- Idea:
  - replace the simple-inline markdown row accumulator in
    `appendSimpleInlineWorksheetTextPrepared(...)` from `map[int]string` to `[]string`
  - append cells by `cellCol-1` directly and trim only trailing empty cells at row flush
- Why:
  - a temporary shape probe on `testRecordSizeExceeded.xlsx` showed an unusually regular worksheet:
    - `rows=200000`
    - `cells=3000000`
    - `maxCol=15`
    - `maxRowWidth=15`
    - `parsedRefs=3000000`
    - `rowStartsAt1=200000`
    - `sequentialWithinRow=2800000`
  - this made the current `map[int]string` row container look like pure overhead on the dominant
    simple-inline markdown workload
- Validation:
  - temporary parity check against the current implementation on the same worksheet
  - same-tree benchmark:
    - `BenchmarkXLSXSimpleInlineRowSliceVariants -count=3`
  - integrated validation after applying:
    - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx -count=5`
    - `cmd/compatcheck -repeat 5 ... testRecordSizeExceeded.xlsx`
    - `go test ./...`
- Result:
  - parity passed
  - same-tree benchmark looked promising:
    - current:
      `2381263500 / 3238469000 / 2584495000 ns/op`
      `1542 MB`, `7.900M allocs/op`
    - row-slice:
      `2272514900 / 2587540500 / 2547998800 ns/op`
      `1472 MB`, `7.600M allocs/op`
  - integrated benchmark after applying still looked mixed:
    - `2401889200 / 3072494800 / 3012996800 / 4022670400 / 2890919300 ns/op`
    - `1472 MB`, `7.600M allocs/op`
  - repeat-aware single-file rerun did not beat the earlier retained current baseline strongly
    enough:
    - `ms=3957 min=2969 max=4342`, `text=181333430`, `images=0`
  - repository regression passed after revert:
    - `ok officeread 182.484s`
- Interpretation:
  - this is the closest recent structural candidate to a keepable optimization: the allocation
    reduction is real and the same-tree benchmark is directionally good
  - but the integrated `compatcheck` evidence is still too weak against the existing retained
    baseline to keep it under the current standard
- Decision:
  - Reverted.

## 2026-07-06 retained: simple-inline switches to text-only after markdown row cap

- Idea:
  - in `appendSimpleInlineWorksheetTextPrepared(...)`, stop doing per-cell coordinate parsing once
    worksheet markdown has already reached `maxMarkdownTableRows`
  - keep the remaining worksheet scan on the existing text extraction path only
- Why:
  - `testRecordSizeExceeded.xlsx` has `200000` rows, but the structured markdown layer retains only
    the first `50000`
  - after that point, `r=` lookup, `cellRefIndexes(...)`, and markdown cell preparation are all
    wasted work for the discarded tail
- Validation:
  - targeted tests:
    - `go test -run 'TestExtractXLSX|TestExtractWorksheet|TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue' ./`
  - focused hotspot benchmarks:
    - `go test -run '^$' -bench 'BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx$|BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx$|BenchmarkExtractXLSXHotspots/...testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
  - repeat-aware pair rerun:
    - `cmd/compatcheck -repeat 3 ... testRecordSizeExceeded.xlsx 00012389.xlsx`
  - decisive `.xlsx` keyset rerun:
    - `cmd/compatcheck -repeat 3 ... <xlsx-keyset-21>`
  - parity / regression:
    - `cmd/compatcheck ... <xlsx-keyset-21>`
    - `go test ./...`
- Result:
  - focused hotspot evidence was mixed but favorable on the integrated end-to-end path:
    - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`:
      `3130030700 ns/op -> 3079023600 ns/op`
    - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`:
      `1931084000 ns/op -> 2295981300 ns/op`
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`:
      `1423902100 ns/op -> 1229985100 ns/op`
    - allocs on integrated extract hotspot:
      `2249931048 B/op`, `12052922 allocs/op`
      `-> 1974691216 B/op`, `7602921 allocs/op`
  - repeat-aware pair rerun improved materially:
    - `00012389.xlsx`: `2056 ms -> 1758 ms`
    - `testRecordSizeExceeded.xlsx`: `6162 ms -> 2816 ms`
  - decisive 21-file `.xlsx` keyset rerun improved clearly with parity preserved:
    - retained previous total: `23930 ms`
    - current total: `20827 ms`
    - parity: `21/21 ok`, `textBytes=279385415`, `images=1`
  - repository regression passed:
    - `ok officeread 140.403s`
- Interpretation:
  - this keeps the retained shared-scan path intact and removes a large block of useless work from
    the simple-inline markdown tail on very large worksheets
  - despite one local markdown-heavy benchmark regressing, the repeat-aware workload evidence is
    decisive enough to retain the change under the current bar
- Decision:
  - Retained.
