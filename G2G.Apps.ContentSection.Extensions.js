G2G.Utilities.Workbook = function() {
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

G2G.Apps.ContentSection.prototype.GetSheetTitle = function(roleTitle, fileNameSafe) {
	var title = 'Pricing for ' + roleTitle;
	if (fileNameSafe) {
    	return title.replace(/[ &/:]/g, '-');
	} else {
		return title;
	}
};

G2G.Apps.ContentSection.prototype.WorksheetFromData = function(data){
	var title = this.GetSheetTitle(data);
   	// create object for worksheet
 	var ws = {};
 	var cellRef = XLSX.utils.encode_cell({ c: 0, r: 0 });
 	var cell = { v: title, t: 's' };
	ws[cellRef] = cell;
 	return ws;
};

G2G.Apps.ContentSection.prototype.ExportWorksheetToExcel = function(ws, sheetTitle) {
  	var wb = new G2G.Utilities.Workbook();
  	wb.SheetNames.push(sheetTitle);
    wb.Sheets[sheetTitle] = ws;
  	var wbOut = XLSX.write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });
  	saveAs(new Blob([G2G.Utilities.S2Abs(wbOut)], { type: 'application/octet-stream' }), sheetTitle.replace(/[ &/:]/g,'-') + '.xslx');
};
		
G2G.Apps.ContentSection.prototype.DownloadXLSX = function () {
	debugger;
	var ws = this.WorksheetFromData(this.RoleData);
    var title = this.GetSheetTitle(this.RoleData.results[0].Role.Title);
    this.ExportWorksheetToExcel(ws, title);
};

G2G.Apps.ContentSection.prototype.PopulatePricing = function (roleId) {
    var app = this;
    // debugger;
    var deferred = $1_10_2.Deferred();
    var apiUrl;
    if (roleId) {
      apiUrl = String.format("/professional-services/_api/Web/Lists/GetByTitle('Vendor Rates')/Items/?$Select=Role/Id,Role/Title,Vendor/Id,Vendor/Title,Vendor/FirstName,Vendor/LastName,Vendor/email,Vendor/Phone,Vendor/SecondaryPhone,Year1,Year2,Year3&$Expand=Role,Vendor&$Filter=Role/Id eq {0}", roleId);
    } else {
      apiUrl = "/professional-services/_api/Web/Lists/GetByTitle('Vendor Rates')/Items/?$Select=Role/Id,Role/Title,Vendor/Id,Vendor/Title,Vendor/FirstName,Vendor/LastName,Vendor/email,Vendor/Phone,Vendor/SecondaryPhone,Year1,Year2,Year3&$Expand=Role,Vendor&$Top=1"
	}
	
	G2G.Utilities.spApiJson(apiUrl, Function.createDelegate(this, function(data) {
		this.RoleData = data.d;
 	   	//$1_10_2.when(this.LazyGetTemplate('staff-role-pricing'), this.LazyGetTemplate('staff-role-pricing-csv')).done(function () {
 	   	this.LazyGetTemplate('staff-role-pricing').done(Function.createDelegate(this, function () {
 	   		debugger;
 	   		var markup = $1_10_2.render['staff-role-pricing'](data.d);
        	// var csv = $1_10_2.render['staff-role-pricing-csv'](data.d);
        	// var filename = 'Pricing-for-' + data.d.results[0].Role.Title.replace(/[ &/:]/g, '-') + '.csv';
        	// var csvData = 'data:text/csv;charset-utf-8,' + encodeURI(csv);

        	$1_10_2(this.popId).html(markup).enhanceWithin();

        	var $downloadLink = $1_10_2(this.popId).find('a.download-data');
                        
        	$downloadLink.click(Function.createDelegate(this, this.DownloadXLSX));
        	        
	        deferred.resolve();
	    }));
	}));
	return deferred;
};