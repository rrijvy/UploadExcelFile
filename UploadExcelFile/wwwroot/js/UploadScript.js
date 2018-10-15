function fileUpload(id) {
    var inputFile = document.getElementById(id);
    var files = inputFile.files;
    var fdata = new FormData();
    for (var i = 0; i <= files.length; i++) {
        fdata.append("files", files[i]);
    }
    $.ajax({
        url: '/Upload',
        data: fdata,
        processData: false,
        contentType: false,
        type: 'POST',
        success: function (data) {
            alert("File upload.");
        }
    });



    
        
    
}