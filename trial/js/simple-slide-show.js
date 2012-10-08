var delay = 1000;
var start_frame = 0;
var end_frame;
var cur_frame = 0;
var timeout;
var lis;

function init() {
	lis = $('slide-images').getElementsByTagName('li');
	
	for( i=0; i < lis.length; i++){
		if(i!=0){
			lis[i].style.display = 'none';
		}
	}
	end_frame = lis.length -1;
	document.getElementById('slideImg-text').innerText = " 1 of " + lis.length;
	document.getElementById('slideImg-text').textContent = " 1 of " + lis.length;
	start_slideshow(start_frame, end_frame, delay, lis);
	
}

function nextSlide() {
    lis = $('slide-images').getElementsByTagName('li');
    Effect.Fade(lis[frame]);
    cur_frame = cur_frame + 1;
    clearTimeout(timeout);
    timeout = setTimeout(fadeInOut(cur_frame, start_frame, end_frame, delay, lis));    
}

function a() {
    alert("s");
}

function previousSlide() {
    lis = $('slide-images').getElementsByTagName('li');    
    Effect.Fade(lis[frame]);
    cur_frame = cur_frame - 1;
    clearTimeout(timeout);
    timeout = setTimeout(fadeInOut(cur_frame, start_frame, end_frame, delay, lis));    
}

function first() {
    setTimeout(fadeInOut(start_frame, start_frame, end_frame, delay, lis));    
}

function last() {
    setTimeout(fadeInOut(end_frame, start_frame, end_frame, delay, lis));
}

function start_slideshow(start_frame, end_frame, delay, lis) {
	timeout = setTimeout(fadeInOut(start_frame,start_frame,end_frame, delay, lis), delay);
}

function stop_slideshow(start_frame, end_frame, delay, lis) {
    clearTimeout(timeout);
}

function fadeInOut(frame, start_frame, end_frame, delay, lis) {
    return (function() {
        lis = $('slide-images').getElementsByTagName('li');       
        Effect.Fade(lis[frame]);
        cur_frame = frame;        
        if (frame == end_frame) { frame = start_frame; } else { frame++; }
        lisAppear = lis[frame];
        setTimeout("Effect.Appear(lisAppear);", 0);
        document.getElementById('slideImg-text').innerText = (frame + 1) + " of " + lis.length;
        document.getElementById('slideImg-text').textContent = (frame + 1) + " of " + lis.length;
        timeout = setTimeout(fadeInOut(frame, start_frame, end_frame, delay), delay + 1850);

        
    })
	
}


Event.observe(window, 'load', init, false);