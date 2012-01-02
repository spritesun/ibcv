<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/admin_header.inc"-->	

<div class="grid_14 prefix_1 suffix_1">

  <form action="./admin/create_activity.asp" method="post">
    <h2>通告名稱</h2>
    <input type="text" name="name">
    <h2>日期</h2>             
    <input type="text" name="date">
    <h2>時間</h2>             
    <input type="text" name="time">
    <h2>地點</h2>             
    <input type="text" name="venue">
    <h2>内容介紹</h2>
    <textarea rows="20" cols="60" name="content"></textarea>
    <input type="submit" value="提交通告"></input>
  </form>

</div>

<!--#include file="../includes/footer.inc"-->