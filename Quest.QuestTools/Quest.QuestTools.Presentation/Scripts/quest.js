/*
Performs an AJAX call

url: Server url and webmethod, action or API

params: data that will be sent through this call

done: function which will be executed in case of a success

fail: function which will be executed in case of failure

always: function which will be executed after performing this call regardless of whether it was successful or not.
*/
function performAjaxCallback(url, params, done, fail, always)
{
    $.ajax({
        type: "POST", //HTTP method
        url: url,
        data: JSON.stringify({data: JSON.stringify(params)}), 
        contentType: "application/json; charset=utf-8",
        dataType: "json"
     })
     .done(done)
     .fail(fail)
     .always(always);     
}
function performAjaxGet(url, done, fail, always) {
    $.ajax({
        type: "GET", //HTTP method
        url: url,
        contentType: "application/json; charset=utf-8",
        dataType: "json"
    })
        .done(done)
        .fail(fail)
        .always(always);  
}
function performAjaxRequest(url, data, done, fail, always) {
    $.ajax({
        type: "POST", //HTTP method
        url: url,
        data: JSON.stringify(data),
        contentType: "application/json; charset=utf-8",
        dataType: "json",
    })
        .done(done)
        .fail(fail)
        .always(always);
}
function getLoadingImgHtml() {
    return "<div><img src=\"/images/loading.gif\" class=\"loadingImage\"></div>";
}
function getDataThroughAjaxRequest(url, done, fail, always) {
    $.ajax({
        type: "GET", //HTTP method
        url: url,
        contentType: "application/json; charset=utf-8",
        dataType: "json",
    })
        .done(done)
        .fail(fail)
        .always(always);
}
function showDangerAlert(msg) {
    $.notify({
        title: '<center><strong>Important</strong></center>',
        message: '<center>' + msg + '<center>'
    	}, {
            type: 'danger',
            z_index: 5000,
            placement: {
                align: "center"
            },
            offset: {
                y: 200,
                x: 200
            },
            animate: {
				enter: 'animated fadeInRight',
				exit: 'animated fadeOutRight'
			},
            delay: 1000
			
    });
}
function showWarningAlert(msg) {
    $.notify({
        title: '<center><strong>Important</strong></center>',
        message: '<center>' + msg + '<center>'
    	}, {
            type: 'warning',
            z_index: 5000,
            placement: {
                align: "center"
            },
            offset: {
                y: 200,
                x: 200
            },
            animate: {
				enter: 'animated fadeInRight',
				exit: 'animated fadeOutRight'
			},
            delay: 1000
			
        });
}
function showSuccessAlert(msg) {
   	$.notify({
        	title: '<center><strong>Success!</strong></center>',
        	message: '<center>' + msg + '<center>'
    	}, {
            type: 'success',
            z_index: 5000,
            placement: {
                align: "center"
            },
            offset: {
                y: 200,
                x: 200
            },
            animate: {
                enter: 'animated fadeInUp',
                exit: 'animated fadeOutDown'
            },
			delay: 5000
        });
}