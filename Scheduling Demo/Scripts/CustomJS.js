$(document).ready(function () {
    $(".checkbox1").change(function () {
        if ($(this).is(":checked")) {
            $(this).closest('tr').css('background-color', '#dfffcc');
            $(this).attr("checked", returnVal);
        }
        else {
            $(this).closest('tr').css('background-color', 'white');
            $(this).attr("checked", returnVal);
        }


    });
});

$(document).ready(function () {
    $('#form').submit(function () {
        var date = $('#GroupDate').val();
        var regex = new RegExp('[0-9]{2}/[0-9]{2}/[0-9]{4}$');
        if (regex.test(date)) {
            $('.Processing').css('display', 'block');
        }
        else {
            return false;
        }
    });
});

$(document).ready(function () {
    $('#formCheckIn').submit(function () {
        $('.Processing').css('display', 'block');
    });
})
