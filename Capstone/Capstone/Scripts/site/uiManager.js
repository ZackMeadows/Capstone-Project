// ---------------------------------------
// ABZ Capstone
// ---------------------------------------

$(document).ready(function(){
    $("input:file").on("change", function () {
        var input = $(this), label = input.val().replace(/\\/g, '/').replace(/.*\//, '');
        if (input.val()) {
            $(this).parent().find('.btn-label').text(label);
            $(this).parent().parent().find('#schedule-generate').css("display", "inline-block");
        }
    })
})
