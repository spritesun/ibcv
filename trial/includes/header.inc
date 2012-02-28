<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>IBCV</title>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<BASE href="http://www.ibcv.org.au/trial/">		
	<link rel="stylesheet" href="css/styles.css" type="text/css" media="screen" />
	<link rel="stylesheet" href="css/slimbox2.css" type="text/css" media="screen" />
	<script type="text/javascript" src="js/jquery-1.3.2.min.js"></script>
	<script type="text/javascript" src="js/jquery.cycle.all.min.js"></script>
	<script text="text/javascript" src="js/nantien.js"></script>
	<script type="text/javascript" src="js/slimbox2.js"></script>
	<script src="js/prototype.js" type="text/javascript"></script>
	<script src="js/scriptaculous.js" type="text/javascript"></script>
	<script text="text/javascript" src="js/simple-slide-show.js"></script>
	
	
	<%
		function GetCnDay(i)
			select case i
				case vbMonday
					GetCnDay = "一"
				case vbTuesday
					GetCnDay = "二"
				case vbWednesday 
					GetCnDay = "三"
				case vbThursday 
					GetCnDay = "四"
				case vbFriday 
					GetCnDay = "五"
				case vbSaturday
					GetCnDay = "六"
				case vbSunday 
					GetCnDay = "日"
			end select
		end function	
	%> 
</head>

<body>
	<div class="container_16 page front show_grid clearfix">
		<div class="top clearfix">
			<div class="logo grid_4">
				<a href="./"><img src="img/logo.png" border="0 "></a>
			</div>
			<div class="topNav grid_12">
				<div class="left">
					<%response.write(Year(Date)& "年" & Month(Date) & "月" & Day(Date) & "日" )%>
					 <%response.write("星期" & GetCnDay(WeekDay(Date))) %>
					 <%response.write(FormatDateTime(now, 3)) %>
				</div>
				<div class="right">  
					開放時間: 每周二至周日及公眾假期 9:00AM-5:00PM
					<a href="../FGS_Mel/FGS_intro.htm" id="langFont">ENGLISH</a>
				</div>
			</div>
			<div class="nav grid_12">
				<div class="left"><img src="img/nav-left.png"></div>
				<div class="right"><img src="img/nav-right.png"></div>
				<ul>
					<li><a href="./">首頁</a></li>
					<li><a>活動預告</a>
					  <ul>
					    <li><a href="./ceremonies.asp">法會活動</a></li>
					    <li><a href="./courses.asp">社教課程</a></li>
					  </ul>
					</li>
					<li><a href="./news_list.asp">活動新聞</a></li>
					<li><a>墨爾本道場</a>
					  <ul>
					    <li><a href="./melbourne.asp">佛光緣美術館</a></li>
					    <li><a href="./boxhill.asp">博士山佛光緣</a></li>
					  </ul>
					</li>
					<li><a href="./chinese-school.asp">中華學校</a>
					</li>
					<li><a>佛光人</a>
						<ul>
							<li><a href="./bio.asp">星雲大師</a></li>
							<li><a href="./international.asp">國際佛光會墨爾本協會</a></li>
							<li><a href="./volunteer.asp">義工</a></li>
							<li><a href="./youth.asp">青年團</a></li>
						</ul>
					</li>					
					<li><a>功德園地</a>
						<ul>
							<li><a href="./worship.asp">佛像供養</a></li>
							<li><a href="./light.asp">光明燈</a></li>
						</ul>
					</li>
					<li><a href="./contact.asp">聯絡我們</a></li>
				</ul>
			</div>
			<div class="banner grid_16">
				<img src="img/banner-img.jpg">
			</div>
		</div>
	
		<div class="middle clearfix">