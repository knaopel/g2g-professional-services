G2G.Utilities.Workbook = function () {
	if (!this instanceof G2G.Utilities.Workbook) return new G2G.Utilities.Workbook();
	this.SheetNames = [];
	this.Sheets = {};
};

G2G.Utilities.S2Abs = function (s) {
	var buff = new ArrayBuffer(s.length);
	var view = new Uint8Array(buff);
	for (var i = 0; i != s.length; ++i) {
		view[i] = s.charCodeAt(i) & 0xFF;
	}
	return buff;
};

G2G.Apps.ContentSection.prototype.LazyGetTemplate = function (name) {
	var deferred = $1_10_2.Deferred();
	if ($1_10_2.templates[name]) {
		deferred.resolve();
	} else {
		G2G.Utilities.GetContent(this.templateUrl + name + '.tmpl.html', function (tmplMarkup) {
			$1_10_2.templates(name, tmplMarkup);
			deferred.resolve();
		});
	}
	return deferred.promise();
};

G2G.Apps.ContentSection.prototype.GetSheetTitle = function (roleTitle, format) {
	var title = 'Pricing for ' + roleTitle;
	if (format) {
		switch (format) {
			case 'fileName':
				title = title.replace(/[ &/:]/g, '-');
				break;
			case 'sheetName':
				title = title.replace(/[/\\*\[\]\?:]/g, '');
				break;
			default:
				// unrecognized do not change
				break;
		}
	}
	return title;
};

G2G.Apps.ContentSection.prototype.WorksheetFromData = function (data) {
	function datenum(v, date1904) {
		if (date1904) v + 1462;
		var epoch = Date.parse(v);
		return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
	}

	function sheet_from_array_of_arrays(data, opts) {
		var ws = { '!merges': [] };
		var range = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };

		// insert row 0
		var cell = { v: 'Pricing for ROle Titkle', t: 's' };
		var cellRef = XLSX.utils.encode_cell({ c: 0, r: 0 });
		ws[cellRef] = cell;
		ws['!merges'].push({ e: { c: 3, r: 0 }, s: { c: 0, r: 0 } });

		// Heading row
		range.e.r++;
		cell.v = 'Vendor Name';
		cellRef = XLSX.utils.encode_cell({ c: range.e.c, r: range.e.r });
		ws[cellRef] = cell;

		// data.forEach(function (row) {
		// 	range.s.r++;
		// 	cell.v = row.Vendor.Title;
		// 	cellRef = XLSX.utils.encode_cell({ c: range.s.c, r: range.s.r });
		// 	ws[cellRef] = cell;
		// });

		ws['!ref'] = XLSX.utils.encode_range(range);

		// data = [["Pricing for Role Title"]];

		// for (var R = 0; R != data.length; ++R) {
		// 	for (var C = 0; C != data[R].length; ++C) {
		// 		if (range.s.r > R) range.s.r = R;
		// 		if (range.s.c > C) range.s.c = C;
		// 		if (range.e.r < R) range.e.r = R;
		// 		if (range.e.c < C) range.e.c = C;
		// 		var cell = { v: data[R][C] };
		// 		if (cell.v == null) continue;
		// 		var cell_ref = XLSX.utils.encode_cell({ c: C, r: R });

		// 		if (typeof cell.v === 'number') cell.t = 'n';
		// 		else if (typeof cell.v === 'boolean') cell.t = 'b';
		// 		else if (cell.v instanceof Date) {
		// 			cell.t = 'n';
		// 			cell.z = XLSX.SSF._table[14];
		// 			cell.v = datenum(cell.v);
		// 		}
		// 		else cell.t = 's';

		// 		ws[cell_ref] = cell;
		// 	}
		// }

		// if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
		// ws['!merges'].push({ e: { c: 3, r: 1 }, s: { c: 0, r: 1 } });
		// ws['!cols'] = []
		return ws;
	}

	// var _data = [[this.GetSheetTitle(data.results[0].Role.Title)], ["Vendor Name", "Year 1", "Year 2", "Year 3", "Vendor Contact", "Vendor Email", "Primary Phone"], [true, false, null, "sheetjs"], ["foo", "bar", new Date("2014-02-19T14:30Z"), "0.3"], ["baz", null, "qux"]];
	var ws_name = this.GetSheetTitle(data.results[0].Role.Title, 'sheetName');

	// function Workbook() {
	// 	if (!(this instanceof Workbook)) return new Workbook();
	// 	this.SheetNames = [];
	// 	this.Sheets = {};
	// }

	var wb = new G2G.Utilities.Workbook(), ws = sheet_from_array_of_arrays(data.results);

	wb.SheetNames.push(ws_name);
	wb.Sheets[ws_name] = ws;

	var wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });

	// function s2abs(s) {
	// 	var buf = new ArrayBuffer(s.length);
	// 	var view = new Uint8Array(buf);
	// 	for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
	// 	return buf;
	// }

	saveAs(new Blob([G2G.Utilities.S2Abs(wbout)], { type: "application/octet-stream" }), this.GetSheetTitle(data.results[0].Role.Title, 'fileName') + ".xlsx");

	// var title = this.GetSheetTitle(data.results[0].Role.Title);
	// // create object for worksheet
	// var ws = {};
	// var range = { s: { c: 1, r: 1 }, e: { c: 0, r: 0 } };
	// var cellRef = XLSX.utils.encode_cell({ c: 0, r: 0 });
	// var cell = { v: title, t: 's' };
	// ws[cellRef] = cell;
	// ws['!ref'] = XLSX.utils.encode_range(range);
	// return ws;
};

G2G.Apps.ContentSection.prototype.ExportWorksheetToExcel = function (ws, sheetTitle) {
	var wb = new G2G.Utilities.Workbook();
	wb.SheetNames.push(sheetTitle);
	wb.Sheets[sheetTitle] = ws;
	var wbOut = XLSX.write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });
	saveAs(new Blob([G2G.Utilities.S2Abs(wbOut)], { type: 'application/octet-stream' }), sheetTitle.replace(/[ &/:]/g, '-') + '.xlsx');
};

G2G.Apps.ContentSection.prototype.DownloadXLSX = function () {
	debugger;
	this.WorksheetFromData(this.RoleData);

	// var ws = this.WorksheetFromData(this.RoleData);
	var title = this.GetSheetTitle(this.RoleData.results[0].Role.Title, 'sheetName');
	// this.ExportWorksheetToExcel(ws, title);
};

G2G.Apps.ContentSection.prototype.PopulatePricing = function (roleId) {
	var app = this;
	var deferred = $1_10_2.Deferred();
	var apiUrl;
	if (roleId) {
		apiUrl = String.format("/professional-services/_api/Web/Lists/GetByTitle('Vendor Rates')/Items/?$Select=Role/Id,Role/Title,Vendor/Id,Vendor/Title,Vendor/FirstName,Vendor/LastName,Vendor/email,Vendor/Phone,Vendor/SecondaryPhone,Year1,Year2,Year3&$Expand=Role,Vendor&$Filter=Role/Id eq {0}", roleId);
	} else {
		apiUrl = "/professional-services/_api/Web/Lists/GetByTitle('Vendor Rates')/Items/?$Select=Role/Id,Role/Title,Vendor/Id,Vendor/Title,Vendor/FirstName,Vendor/LastName,Vendor/email,Vendor/Phone,Vendor/SecondaryPhone,Year1,Year2,Year3&$Expand=Role,Vendor&$Top=1"
	}

	G2G.Utilities.spApiJson(apiUrl, Function.createDelegate(this, function (data) {
		this.RoleData = data.d;
		this.LazyGetTemplate('staff-role-pricing').done(Function.createDelegate(this, function () {
			var markup = $1_10_2.render['staff-role-pricing'](data.d);

			$1_10_2(this.popId).html(markup).enhanceWithin();

			var $downloadLink = $1_10_2(this.popId).find('a.download-data');

			$downloadLink.click(Function.createDelegate(this, this.DownloadXLSX));

			deferred.resolve();
		}));
	}));
	return deferred;
};