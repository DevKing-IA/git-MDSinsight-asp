
//Ping server every 60 seconds
//var sessServerAliveTime = 10 * 60 * 60;

//Ping server every 3 minutes
var sessServerAliveTime = 10 * 60 * 180;

//Set jQuery session timeout to 60 seconds
//var sessionTimeout = 1 * 60000;

//Set jQuery "manual" session timeout to 6 hours, or 360 minutes
var sessionTimeout = 360 * 60000;

var sessLastActivity;
var idleTimer, remainingTimer;
var isTimout = false;

var sess_intervalID, idleIntervalID;
var sess_lastActivity;
var timer;
var isIdleTimerOn = false;

localStorage.setItem('sessionSlide', 'isStarted');

function sessPingServer() {
    if (!isTimout) {
        //$.ajax({
        //    url: 'sessionPingServer.asp',
        //    dataType: "json",
        //    async: false,
        //    type: "POST"
        //});
        
        //return true;
        
        //console.log("sessPingServer");
        
    	$.ajax({
			type:"POST",
			url: "/inc/sessionPingServer.asp",
			cache: false,
			data: "action=CheckIfServerSessionAlive",
			success: function(response)
			 {
				console.log("response: " + response);
				
				if (response == "ALIVE")
				{					
				    return true;									
				}
				else {

					if (localStorage.getItem('sessionSlide') !== "loggedOut") {
						$("#session-expire-warning-modal").modal('hide');
				        $("#session-expired-modal").modal('show');
				        clearInterval(sess_intervalID);
				        clearInterval(idleIntervalID);
				        localStorage.setItem('sessionSlide', 'loggedOut');			
				        isTimout = true;
				    }


				}
               	 
             }
		});
  
    }
}

function sessServerAlive() {
    sess_intervalID = setInterval('sessPingServer()', sessServerAliveTime);
}

function initSessionMonitor() {
	
    $(document).bind('keypress.session', function (ed, e) {
        sessKeyPressed(ed, e);
    });
    $(document).bind('mousedown keydown', function (ed, e) {

        sessKeyPressed(ed, e);
    });
    sessServerAlive();
    startIdleTime();

}



$(window).scroll(function (e) {

    localStorage.setItem('sessionSlide', 'isStarted');
    startIdleTime();

});

function sessKeyPressed(ed, e) {
    var target = ed ? ed.target : window.event.srcElement;
    var sessionTarget = $(target).parents("#session-expire-warning-modal").length;

    if (sessionTarget != null && sessionTarget != undefined) {
        if (ed.target.id != "btnSessionExpiredCancelled" && ed.target.id != "btnSessionModal" && ed.currentTarget.activeElement.id != "session-expire-warning-modal" && ed.target.id != "btnExpiredOk"
             && ed.currentTarget.activeElement.className != "modal fade modal-overflow in" && ed.currentTarget.activeElement.className != 'modal-header'
    && sessionTarget != 1 && ed.target.id != "session-expire-warning-modal") {
            localStorage.setItem('sessionSlide', 'isStarted');
            startIdleTime();
        }
    }
}

function startIdleTime() {
    stopIdleTime();
    localStorage.setItem("sessIdleTimeCounter", $.now());
    idleIntervalID = setInterval('checkIdleTimeout()', 1000);
    isIdleTimerOn = false;
}

var sessionExpired = document.getElementById("session-expired-modal");
function sessionExpiredClicked(evt) {
	localStorage.setItem('sessionSlide', "loggedOut");
    window.location = "/logout.asp";
}





sessionExpired.addEventListener("click", sessionExpiredClicked, false);
function stopIdleTime() {
    clearInterval(idleIntervalID);
    clearInterval(remainingTimer);
}

function checkIdleTimeout() {
     // $('#sessionValue').val() * 60000;
    var idleTime = (parseInt(localStorage.getItem('sessIdleTimeCounter')) + (sessionTimeout)); 
    if ($.now() > idleTime + 60000) {
        $("#session-expire-warning-modal").modal('hide');
        $("#session-expired-modal").modal('show');
        clearInterval(sess_intervalID);
        clearInterval(idleIntervalID);

        $('.modal-backdrop').css("z-index", parseInt($('.modal-backdrop').css('z-index')) + 100);
        $("#session-expired-modal").css('z-index', 2000);
        $('#btnExpiredOk').css('background-color', '#428bca');
        $('#btnExpiredOk').css('color', '#fff');

        isTimout = true;

        sessLogOut();

    }
    else if ($.now() > idleTime) {
        ////var isDialogOpen = $("#session-expire-warning-modal").is(":visible");
        if (!isIdleTimerOn) {
            ////alert('Reached idle');
            localStorage.setItem('sessionSlide', false);
            countdownDisplay();

            $('.modal-backdrop').css("z-index", parseInt($('.modal-backdrop').css('z-index')) + 500);
            $('#session-expire-warning-modal').css('z-index', 1500);
            $('#btnOk').css('background-color', '#5cb85c');
            $('#btnOk').css('color', '#fff');
            $('#btnSessionExpiredCancelled').css('background-color', '#428bca');
            $('#btnSessionExpiredCancelled').css('color', '#fff');
            $('#btnLogoutNow').css('background-color', '#428bca');
            $('#btnLogoutNow').css('color', '#fff');

            $("#seconds-timer").empty();
            $("#session-expire-warning-modal").modal('show');


            isIdleTimerOn = true;
        }
    }
}

$("#btnSessionExpiredCancelled").click(function () {
    $('.modal-backdrop').css("z-index", parseInt($('.modal-backdrop').css('z-index')) - 500);
});

$("#btnOk").click(function () {
    $("#session-expire-warning-modal").modal('hide');
    $('.modal-backdrop').css("z-index", parseInt($('.modal-backdrop').css('z-index')) - 500);
    startIdleTime();
    clearInterval(remainingTimer);
    localStorage.setItem('sessionSlide', 'isStarted');
});

$("#btnLogoutNow").click(function () {
    localStorage.setItem('sessionSlide', 'loggedOut');
    window.location = "/logout.asp";
    sessLogOut();
    $("#session-expired-modal").modal('hide');

});
$('#session-expired-modal').on('shown.bs.modal', function () {
    $("#session-expire-warning-modal").modal('hide');
    $(this).before($('.modal-backdrop'));
    $(this).css("z-index", parseInt($('.modal-backdrop').css('z-index')) + 1);
});

$("#session-expired-modal").on("hidden.bs.modal", function () {
    window.location = "/logout.asp";
});
$('#session-expire-warning-modal').on('shown.bs.modal', function () {
    $("#session-expire-warning-modal").modal('show');
    $(this).before($('.modal-backdrop'));
    $(this).css("z-index", parseInt($('.modal-backdrop').css('z-index')) + 1);
});


function countdownDisplay() {

    var dialogDisplaySeconds = 60;

    remainingTimer = setInterval(function () {
        if (localStorage.getItem('sessionSlide') == "isStarted") {
            $("#session-expire-warning-modal").modal('hide');
            startIdleTime();
            clearInterval(remainingTimer);
        }
        else if (localStorage.getItem('sessionSlide') == "loggedOut") {         
            $("#session-expire-warning-modal").modal('hide');
            $("#session-expired-modal").modal('show');
        }
        else {

            $('#seconds-timer').html(dialogDisplaySeconds);
            dialogDisplaySeconds -= 1;
        }
    }
    , 1000);
};

function sessLogOut() {
   // $.ajax({
   //     url: 'Logout.html',
   //     dataType: "json",
  //      async: false,
  //      type: "POST"
 //   });
	
	window.location = "/logout.asp";

}
