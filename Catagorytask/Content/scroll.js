$(document).ready(function () {
    var table = $('#AllCatagory_Table').DataTable({
        scrollY: "300px",
        scrollX: true,
        scrollCollapse: true,
        paging: false,
        fixedColumns: {
            left: 1,
            right: 1
        }
    });
});