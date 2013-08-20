$(document).ready(function () {
    $('#grid').jqGrid({
        url: 'callGrid.asp',
        datatype: 'xml',
        mtype: 'GET',
        colNames: ['Id', 'Name', 'Address', 'City', 'State', 'Zip', 'Phone'],
        colModel: [
			{ name: 'Id', index: 'Id', width: 225 },
			{ name: 'Name', index: 'Name', width: 300 },
			{ name: 'AddressLine1', index: 'AddressLine1', width: 300 },
			{ name: 'AddressCity', index: 'AddressCity', width: 240 },
			{ name: 'AddressState', index: 'AddressState', width: 60 },
			{ name: 'AddressZip', index: 'AddressZip', width: 120, sortable: false },
			{ name: 'Phone', index: 'Phone', sortable: false }
		],
        pager: '#pager',
        sortname: 'Name',
        rowNum: 100,
        rowList: [100, 200],
        sortorder: "asc",
        width: 700,
        height: 250,
        caption: 'Results',
        viewrecords: true,
		ondblClickRow: function (id) {
            OnSelect();
        },
        gridComplete: function () {
            var recs = parseInt($("#grid").getGridParam("records"), 10);
            if ((recs == 0) && ($('#current').hasClass('active'))) {
                $("#current").attr('disabled', true);
                setTimeout(function () { $('#query').click(); }, 1000);
            }
        }
    });

});