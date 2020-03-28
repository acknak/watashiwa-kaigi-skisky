const ParentMenuConfig = {
	cName: "Watashiwa Kaigi SkiSky",
	cParent: "Edit"
}

const RenumPagesMenuConfig = {
	cName: "Renumber Pages",
	cParent: "Watashiwa Kaigi SkiSky",
	cExec: "RenumPages()"
}

const SetConfMenuConfig = {
	cName: "Set Confidentiality",
	cParent: "Watashiwa Kaigi SkiSky",
	cExec: "SetConfMenu()"
}

const SetDocNumMenuConfig = {
	cName: "Set Document Number",
	cParent: "Watashiwa Kaigi SkiSky",
	cExec: "SetDocNum()"
}

const SetDocSecMenuConfig = {
	cName: "Set Document Security",
	cParent: "Watashiwa Kaigi SkiSky",
	cExec: "SetDocSec()"
}

app.addSubMenu(ParentMenuConfig);
app.addMenuItem(RenumPagesMenuConfig);
app.addMenuItem(SetConfMenuConfig);
app.addMenuItem(SetDocNumMenuConfig);
app.addMenuItem(SetDocSecMenuConfig);

const pdfHorizontalMargin = 25.4;
const pdfVerticalMargin = 12.7;
const wordHorizontalMargin = 30.0;
const wordVerticalMargin = 35.0;

function RenumPages() {
	renumPagesDialog.doc = this;
	app.execDialog(renumPagesDialog);
}

var renumPagesDialog = {
	doc: null,
	initialize: dialog => {
		dialog.load({
			"nmbr": "2",
			"disp": "1",
		})
	},
	commit: function (dialog) {
		const doc = this.doc;
		this.deleteAllPage();
		const pageRng = new Rng(Math.max(0, dialog.store()["nmbr"] - 1), doc.numPages);
		pageRng
			.arr()
			.forEach((p, i) => {
				const boxWidth = 100;
				const rect = doc.getPageBox("Bleed", p);
				const pageWidth = rect[2] - rect[0];
				var f = doc.addField(setFieldName("shiryoPage")(p), "text", p,
					[
						pageWidth - boxWidth - mm2unitsize(pdfHorizontalMargin, doc, p),
						mm2unitsize(pdfVerticalMargin, doc, p),
						pageWidth - mm2unitsize(pdfHorizontalMargin, doc, p),
						mm2unitsize(pdfVerticalMargin, doc, p) - 15
					]
				);
				f.value = String((i + parseInt(dialog.store()["disp"]))
					+ " / "
					+ String(pageRng.len() + parseInt(dialog.store()["disp"]) - 1));
				f.fillColor = color.transparent;
				f.textSize = 7.1;
				f.alignment = "right";
				f.textFont = "HeiseiKakuGo-W5-UniJIS-UCS2-H";
				f.readonly = true;
			});
		app.alert("ページ番号を振りなおしました", 3, 0, "Watashiwa Kaigi Skisky: Renumber Pages");
	},
	other: function (dialog) {
		this.deleteAllPage();
		app.alert("ページ番号を全て削除しました", 3, 0, "Watashiwa Kaigi Skisky: Renumber Pages");
		dialog.end();
	},
	deleteAllPage: function () {
		const doc = this.doc;
		const pageFieldName = setFieldName("shiryoPage");
		new Rng(0, doc.numPages)
			.arr()
			.map(pageFieldName)
			.forEach(pfn => doc.removeField(pfn));
	},
	description: {
		name: "Watashiwa Kaigi Skisky: Renumber Pages",
		align_children: "align_left",
		width: 360,
		elements: [
			{
				type: "view",
				align_children: "align_left",
				elements: [
					{
						type: "view",
						align_children: "align_row",
						elements: [
							{
								type: "static_text",
								name: "付番開始ページ："
							}, {
								type: "edit_text",
								item_id: "nmbr",
								width: 32,
								SpinEdit: true
							}
						]
					}, {
						type: "view",
						align_children: "align_row",
						elements: [
							{
								type: "static_text",
								name: "　　　　開始番号："
							}, {
								type: "edit_text",
								item_id: "disp",
								width: 32,
								SpinEdit: true
							}
						]
					}, {
						type: "ok_cancel_other",
						ok_name: "OK",
						other_name: "全削除",
						cancel_name: "キャンセル"
					}
				]
			}
		]
	}
};

function SetConfMenu() {
	setConfDialog.doc = this;
	app.execDialog(setConfDialog);
}

var setConfDialog = {
	doc: null,
	initialize: dialog => {
		dialog.load({
			strp: "1",
			endp: "999",
			cnfp: {
				"機密性１情報": -1,
				"機密性２情報（社内限り）": 2,
				"機密性２情報（メンバー限り）": -3,
				"機密性２情報（メンバー・オブザーバー限り）": -4,
				"機密性２情報（メンバー・オブザーバー・部長・課長限り）": -5,
				"機密性２情報（メンバー・オブザーバー・本部事業組織・拠点幹部限り）": -6,
				"機密性３情報（メンバー限り）": -7,
			},
			cnft: "機密性２情報\r共有範囲：社内限り",
		})
	},
	commit: function (dialog) {
		const doc = this.doc;
		const rng = this.pageRng(dialog, this.doc);
		this.deleteConf(rng);
		rng
			.arr()
			.forEach(p => {
				const rect = doc.getPageBox("Bleed", p);
				const pageHeight = rect[1] - rect[3];
				var f = doc.addField(setFieldName("shiryoConf")(p), "text", p,
					[
						mm2unitsize(pdfHorizontalMargin, doc, p),
						pageHeight - mm2unitsize(pdfVerticalMargin, doc, p) + 19,
						mm2unitsize(pdfHorizontalMargin, doc, p) + 500,
						pageHeight - mm2unitsize(pdfVerticalMargin, doc, p),
					]
				);
				f.multiline = true;
				f.value = dialog.store()["cnft"];
				f.fillColor = color.transparent;
				f.textSize = 7.1;
				f.alignment = "left";
				f.textFont = "HeiseiKakuGo-W5-UniJIS-UCS2-H";
				f.readonly = true;
			});
		app.alert("機密性を設定しました", 3, 0, "Watashiwa Kaigi Skisky: Set Confidentiality");
	},
	cnfp: function (dialog) {
		const elements = dialog.store()["cnfp"];
		for (var e in elements) {
			if (elements[e] > 0) {
				dialog.load({
					cnft: e.replace("（", "\n共有範囲：").replace("）", "")
				})
				break;
			}
		}
	},
	other: function (dialog) {
		this.deleteConf(this.pageRng(dialog, this.doc));
		app.alert("指定範囲の機密性を解除しました", 3, 0, "Watashiwa Kaigi Skisky: Set Confidentiality");
		dialog.end();
	},
	pageRng: (dialog, doc) => new Rng(
		Math.max(0, parseInt(dialog.store()["strp"]) - 1),
		Math.min(parseInt(dialog.store()["endp"]), doc.numPages)
	),
	deleteConf: function (rng) {
		const doc = this.doc;
		const pageFieldName = setFieldName("shiryoConf");
		rng.arr()
			.map(pageFieldName)
			.forEach(pfn => doc.removeField(pfn));
	},
	description: {
		name: "Watashiwa Kaigi Skisky: Set Confidentiality",
		align_children: "align_left",
		width: 360,
		elements: [
			{
				type: "view",
				align_children: "align_left",
				elements: [
					{
						type: "view",
						align_children: "align_row",
						elements: [
							{
								type: "static_text",
								name: "開始ページ："
							}, {
								type: "edit_text",
								item_id: "strp",
								width: 32,
								SpinEdit: true
							}
						]
					}, {
						type: "view",
						align_children: "align_row",
						elements: [
							{
								type: "static_text",
								name: "終了ページ："
							}, {
								type: "edit_text",
								item_id: "endp",
								width: 32,
								SpinEdit: true
							}
						]
					}, {
						type: "view",
						align_children: "align_row",
						elements: [
							{
								type: "static_text",
								name: "表示",
								char_height: 2
							}, {
								type: "popup",
								item_id: "cnfp",
								width: 96,
							}
						]
					}, {
						type: "edit_text",
						item_id: "cnft",
						multiline: true,
						width: 320,
						height: 96,
					}, {
						type: "ok_cancel_other",
						ok_name: "OK",
						other_name: "解除",
						cancel_name: "キャンセル",
					}
				]
			}
		]
	}
};

function SetDocNum() {
	setDocNumDialog.doc = this;
	app.execDialog(setDocNumDialog);
}

var setDocNumDialog = {
	doc: null,
	initialize: function (dialog) {
		dialog.load({
			nump: String(this.doc.pageNum + 1),
			strw: "72",
			dnmp: {
				"資料１": 1,
				"審議１": -2,
				"報告１": -3,
				"参考１": -4,
				"限定１": -5,
				"要承１": -6,
				"その他": -7,
			},
			dnmt: "資料１"
		})
	},
	commit: function (dialog) {
		const doc = this.doc;
		const pageNum = parseInt(dialog.store()["nump"]) - 1;
		this.deleteDocNum(pageNum);
		const boxWidth = dialog.store()["strw"];
		const rect = doc.getPageBox("Bleed", pageNum);
		const pageHeight = rect[1] - rect[3];
		const pageWidth = rect[2] - rect[0];
		var f = doc.addField(setFieldName("shiryoDocNum")(pageNum), "text", pageNum,
			[
				pageWidth - boxWidth - mm2unitsize(wordHorizontalMargin, doc, pageNum),
				pageHeight - mm2unitsize(pdfVerticalMargin, doc, pageNum),
				pageWidth - mm2unitsize(wordHorizontalMargin, doc, pageNum),
				pageHeight - mm2unitsize(pdfVerticalMargin, doc, pageNum) - 24
			]
		);
		f.value = dialog.store()["dnmt"];
		f.fillColor = color.transparent;
		f.textSize = 20;
		f.borderColor = color.black;
		f.borderWidth = 1;
		f.alignment = "center";
		f.textFont = "HeiseiKakuGo-W5-UniJIS-UCS2-H";
		f.readonly = true;
		app.alert("資料番号を設定しました", 3, 0, "Watashiwa Kaigi Skisky: Set Document Number")
	},
	dnmp: function (dialog) {
		const elements = dialog.store()["dnmp"];
		for (var e in elements) {
			if (elements[e] > 0) {
				dialog.load({
					dnmt: e
				})
				break;
			}
		}
	},
	other: function (dialog) {
		this.deleteDocNum(this.pageRng(dialog, this.doc));
		app.alert("指定ページの資料番号を削除しました", 3, 0, "Watashiwa Kaigi Skisky: Set Document Number");
		dialog.end();
	},
	pageRng: (dialog, doc) => new Rng(
		Math.max(0, parseInt(dialog.store()["strp"]) - 1),
		Math.min(parseInt(dialog.store()["endp"]), doc.numPages)
	),
	deleteDocNum: function (pfn) {
		const doc = this.doc;
		doc.removeField(setFieldName("shiryoDocNum")(pfn));
	},
	description: {
		name: "Watashiwa Kaigi Skisky: Set Document Number",
		align_children: "align_left",
		width: 360,
		elements: [
			{
				type: "view",
				align_children: "align_left",
				elements: [
					{
						type: "view",
						align_children: "align_row",
						elements: [
							{
								type: "static_text",
								name: "対象ページ："
							}, {
								type: "edit_text",
								item_id: "nump",
								width: 32,
								SpinEdit: true
							}, {
								type: "static_text",
								name: "文字幅："
							}, {
								type: "edit_text",
								item_id: "strw",
								width: 32,
								SpinEdit: true
							}
						]
					},
					{
						type: "cluster",
						name: "表記",
						align_children: "align_left",
						elements: [
							{
								type: "view",
								align_children: "align_row",
								elements: [
									{
										type: "static_text",
										name: "例文",
										char_height: 2
									}, {
										type: "popup",
										item_id: "dnmp",
										width: 96,
									}
								]
							}, {
								type: "edit_text",
								item_id: "dnmt",
								width: 320,
							}
						]
					}, {
						type: "ok_cancel_other",
						ok_name: "OK",
						other_name: "表記削除",
						cancel_name: "キャンセル",
					}
				]
			}
		]
	}
};

function SetDocSec() {
	setDocSecDialog.doc = this;
	app.execDialog(setDocSecDialog);
}

var setDocSecDialog = {
	doc: null,
	initialize: dialog => getSecPoli(dialog),
	commit: function (dialog) {
		setSecPoli(dialog.store(), this.doc);
	},
	other: dialog => {
		editSecPoli();
		getSecPoli(dialog);
	},
	description: {
		name: "Watashiwa Kaigi Skisky: Set Document Security",
		align_children: "align_left",
		width: 350,
		elements: [
			{
				type: "view",
				alignment: "align_fill",
				elements: [
					{
						type: "static_text",
						name: "セキュリティ設定を選択："
					}, {
						type: "list_box",
						item_id: "poli",
						width: 300,
						height: 80
					}, {
						type: "static_text",
						name: "Adobe Acrobat Pro DCでのセキュリティ設定追加方法：\n"
							+ "1.「セキュリティ設定を追加」→「新規」を選択\n"
							+ "2.「パスワードを使用」を選択して「次へ」を選択\n"
							+ "3.「ポリシー名」に会議名を入力（e.g.○○連絡会議）して「次へ」を選択\n"
							+ "4.当該会議資料にかけるべきパスワード等を設定\n"
							+ "5.「完了」を選択",
						height: 100,
					}, {
						type: "static_text",
						name: "セキュリティ設定の編集・削除方法：\n"
							+ "1.「acrobat」「セキュリティーポリシーの管理」でググる\n"
							+ "2. He that is without sin among you, let him first cast a stone at author.",
						height: 60,
					}, {
						alignment: "align_right",
						type: "ok_cancel_other",
						ok_name: "OK",
						cancel_name: "キャンセル",
						other_name: "セキュリティ設定を追加"
					}
				]
			},
		]
	}
};

const getSecPoli = app.trustedFunction(function (dialog) {
	app.beginPriv();
	dialog.load({
		"poli": security.getSecurityPolicies()
			.filter(sp => sp.policyId !== "SP_PKI_default" && sp.policyId !== "SP_STD_default")
			.reduce((map, sp) => {
				map[sp.name] = sp.policyId;
				return map;
			}, {})
	})
	app.endPriv();
});

const setSecPoli = app.trustedFunction(function (results, doc) {
	app.beginPriv();
	const aPols = security.getSecurityPolicies();
	const oMyPolicy = aPols
		.filter((aPol, i) => results["poli"][aPol.name] > 0)
		.pop();
	if (oMyPolicy) {
		const rtn = doc.encryptUsingPolicy({ oPolicy: oMyPolicy });
		if (rtn.errorCode != 0) {
			app.alert("セキュリティポリシー「" + oMyPolicy.name + "」が適用できませんでした。\nエラー: " + rtn.errorText, 0, 0, "Watashiwa Kaigi Skisky: Set Document Security");
		} else {
			app.alert("セキュリティポリシー「" + oMyPolicy.name + "」を適用しました", 3, 0, "Watashiwa Kaigi Skisky: Set Document Security");
		}
	} else {
		app.alert("いずれかのセキュリティポリシーを選択してください", 1, 0, "Watashiwa Kaigi Skisky: Set Document Security");
	}
	app.endPriv();
});

const editSecPoli = app.trustedFunction(function () {
	app.beginPriv();
	security.chooseSecurityPolicy();
	app.endPriv();
});

function Rng(start, end) {
	this.start = start;
	this.end = end;
};
Rng.prototype.len = function () {
	return Math.max(0, this.end - this.start);
};
Rng.prototype.arr = function () {
	var tmp_arr = [];
	for (var i = this.start; i < this.end; i++) tmp_arr.push(i);
	return tmp_arr;
};

const unitsize2mm = (us, doc, page) => us * doc.getUserUnitSize(page) / 72 * 25.4;
const mm2unitsize = (mm, doc, page) => mm / 25.4 * 72 / doc.getUserUnitSize(page);
const setFieldName = fieldId => pageNum => String(fieldId + pageNum);
