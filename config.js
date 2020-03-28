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
		app.alert("�y�[�W�ԍ���U��Ȃ����܂���", 3, 0, "Watashiwa Kaigi Skisky: Renumber Pages");
	},
	other: function (dialog) {
		this.deleteAllPage();
		app.alert("�y�[�W�ԍ���S�č폜���܂���", 3, 0, "Watashiwa Kaigi Skisky: Renumber Pages");
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
								name: "�t�ԊJ�n�y�[�W�F"
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
								name: "�@�@�@�@�J�n�ԍ��F"
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
						other_name: "�S�폜",
						cancel_name: "�L�����Z��"
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
				"�@�����P���": -1,
				"�@�����Q���i�Г�����j": 2,
				"�@�����Q���i�����o�[����j": -3,
				"�@�����Q���i�����o�[�E�I�u�U�[�o�[����j": -4,
				"�@�����Q���i�����o�[�E�I�u�U�[�o�[�E�����E�ے�����j": -5,
				"�@�����Q���i�����o�[�E�I�u�U�[�o�[�E�{�����Ƒg�D�E���_��������j": -6,
				"�@�����R���i�����o�[����j": -7,
			},
			cnft: "�@�����Q���\r���L�͈́F�Г�����",
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
		app.alert("�@������ݒ肵�܂���", 3, 0, "Watashiwa Kaigi Skisky: Set Confidentiality");
	},
	cnfp: function (dialog) {
		const elements = dialog.store()["cnfp"];
		for (var e in elements) {
			if (elements[e] > 0) {
				dialog.load({
					cnft: e.replace("�i", "\n���L�͈́F").replace("�j", "")
				})
				break;
			}
		}
	},
	other: function (dialog) {
		this.deleteConf(this.pageRng(dialog, this.doc));
		app.alert("�w��͈͂̋@�������������܂���", 3, 0, "Watashiwa Kaigi Skisky: Set Confidentiality");
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
								name: "�J�n�y�[�W�F"
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
								name: "�I���y�[�W�F"
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
								name: "�\��",
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
						other_name: "����",
						cancel_name: "�L�����Z��",
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
				"�����P": 1,
				"�R�c�P": -2,
				"�񍐂P": -3,
				"�Q�l�P": -4,
				"����P": -5,
				"�v���P": -6,
				"���̑�": -7,
			},
			dnmt: "�����P"
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
		app.alert("�����ԍ���ݒ肵�܂���", 3, 0, "Watashiwa Kaigi Skisky: Set Document Number")
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
		app.alert("�w��y�[�W�̎����ԍ����폜���܂���", 3, 0, "Watashiwa Kaigi Skisky: Set Document Number");
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
								name: "�Ώۃy�[�W�F"
							}, {
								type: "edit_text",
								item_id: "nump",
								width: 32,
								SpinEdit: true
							}, {
								type: "static_text",
								name: "�������F"
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
						name: "�\�L",
						align_children: "align_left",
						elements: [
							{
								type: "view",
								align_children: "align_row",
								elements: [
									{
										type: "static_text",
										name: "�ᕶ",
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
						other_name: "�\�L�폜",
						cancel_name: "�L�����Z��",
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
						name: "�Z�L�����e�B�ݒ��I���F"
					}, {
						type: "list_box",
						item_id: "poli",
						width: 300,
						height: 80
					}, {
						type: "static_text",
						name: "Adobe Acrobat Pro DC�ł̃Z�L�����e�B�ݒ�ǉ����@�F\n"
							+ "1.�u�Z�L�����e�B�ݒ��ǉ��v���u�V�K�v��I��\n"
							+ "2.�u�p�X���[�h���g�p�v��I�����āu���ցv��I��\n"
							+ "3.�u�|���V�[���v�ɉ�c������́ie.g.�����A����c�j���āu���ցv��I��\n"
							+ "4.���Y��c�����ɂ�����ׂ��p�X���[�h����ݒ�\n"
							+ "5.�u�����v��I��",
						height: 100,
					}, {
						type: "static_text",
						name: "�Z�L�����e�B�ݒ�̕ҏW�E�폜���@�F\n"
							+ "1.�uacrobat�v�u�Z�L�����e�B�[�|���V�[�̊Ǘ��v�ŃO�O��\n"
							+ "2. He that is without sin among you, let him first cast a stone at author.",
						height: 60,
					}, {
						alignment: "align_right",
						type: "ok_cancel_other",
						ok_name: "OK",
						cancel_name: "�L�����Z��",
						other_name: "�Z�L�����e�B�ݒ��ǉ�"
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
			app.alert("�Z�L�����e�B�|���V�[�u" + oMyPolicy.name + "�v���K�p�ł��܂���ł����B\n�G���[: " + rtn.errorText, 0, 0, "Watashiwa Kaigi Skisky: Set Document Security");
		} else {
			app.alert("�Z�L�����e�B�|���V�[�u" + oMyPolicy.name + "�v��K�p���܂���", 3, 0, "Watashiwa Kaigi Skisky: Set Document Security");
		}
	} else {
		app.alert("�����ꂩ�̃Z�L�����e�B�|���V�[��I�����Ă�������", 1, 0, "Watashiwa Kaigi Skisky: Set Document Security");
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
