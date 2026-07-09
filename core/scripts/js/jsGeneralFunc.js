//testalert();
function testalert() {
    alert("GG");
}

//$(function () {
//    //$('#tblResult thead th').click(function () {
//    //    alert("GG");
//    //});

//    $('#tblResult thead th').click(function () {
//        var table = $(this).parents('table').eq(0)
//        var rows = table.find('tr:gt(0)').toArray().sort(comparer($(this).index()))
//        this.asc = !this.asc
//        if (!this.asc) { rows = rows.reverse() }
//        for (var i = 0; i < rows.length; i++) { table.append(rows[i]) }
//    });

//    //$('th').click(function () {
//    //    //var table = $(this).parents('table').eq(0)
//    //    //var rows = table.find('tr:gt(0)').toArray().sort(comparer($(this).index()))
//    //    //this.asc = !this.asc
//    //    //if (!this.asc) { rows = rows.reverse() }
//    //    //for (var i = 0; i < rows.length; i++) { table.append(rows[i]) }
//    //})

//});
$(document).ready(function () {
    $("#myTable").tablesorter();




});


