<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>test demo</title>

<%
	pageContext.setAttribute("APP_PATH", request.getContextPath());
%>
	<script type="text/javascript" src="${APP_PATH}/static/jquery.js"></script>

<body>

		<h1><a href="${APP_PATH}/LAP_Exports">下载</a></h1>
		<h1><a href="${APP_PATH}/LAP_Export">下载</a></h1>

		

		
</body>
</html>