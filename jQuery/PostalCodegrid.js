$(document).ready(function () {

    $('#grid').jqGrid({
        url: 'emptygrid.asp',
        datatype: 'json',
        mtype: 'GET',
        jsonReader: { root: 'Rows',
            page: 'Page',
            total: 'Total',
            records: 'Records',
            repeatitems: false,
            id: 'Id'
        },
        colNames: ['Id', 'City', 'State', 'Zip', 'County'],
        colModel: [
			{ name: 'Id', index: 'Id', hidden: true },
			{ name: 'City', index: 'City', width: 40, sortable: false },
			{ name: 'State', index: 'State', width: 12, sortable: false },
			{ name: 'Zip', index: 'Zip', width: 20, sortable: false },
			{ name: 'County', index: 'County', width: 20, sortable: false }
		],
        pager: '#pager',
        sortname: 'Zip',
        rowNum: 10,
        rowList: [10, 20, 30],
        sortorder: "asc",
        width: 400,
        height: 120,
        caption: 'Postal Codes',
        viewrecords: true,
        ondblClickRow: function (id) {
            OnSelect();
        },
        gridComplete: function () {
            var id = $('#grid').getDataIDs();
            if(id && id.length > 0)
                $('#grid').setSelection(id[0]);
        }
    });
});