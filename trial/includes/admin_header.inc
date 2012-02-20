<!--#include file="../includes/function_custom.inc"-->	

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
          (僅内部使用)
        </div>
        <div class="right">  
        </div>
      </div>
      <div class="nav grid_12">
        <div class="left"><img src="img/nav-left.png"></div>
        <div class="right"><img src="img/nav-right.png"></div>
        <ul>
          <li>
            <a href="./admin/activities.asp">通告管理</a>
            <ul>
              <li><a href="./admin/activities.asp">通告列表</a></li>
              <li><a href="./admin/edit_activity.asp">創建通告</a></li>
            </ul>
          </li>
          <li>
            <a href="./admin/news_list.asp">新聞管理</a>
            <ul>
              <li><a href="./admin/news_list.asp">新聞列表</a></li>
              <li><a href="./admin/edit_news.asp">創建新聞</a></li>
            </ul>
          </li>
          <li><a href="./admin/logout.asp">登出</a></li>
        </ul>
      </div>
    </div>
    <div class="middle clearfix"></div>
