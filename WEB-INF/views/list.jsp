﻿<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<%@taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>员工列表</title>
<%
	pageContext.setAttribute("APP_PATH",request.getContextPath());
%>

<!-- 
	以/开始的路径，是以服务器开始找资源的
 -->
<script type="text/javascript" src="${APP_PATH}/static/jquery.js"></script>
<link rel="stylesheet"
	href="${APP_PATH}/static/bootstrap-3.3.7-dist/css/bootstrap.min.css">
<script
	src="${APP_PATH}/static/bootstrap-3.3.7-dist/js/bootstrap.min.js"></script>

</head>
<body>

	<!-- 搭建显示页面 -->
	<div class="container">
	
		<!-- 标题 -->
		<div class="row">
			<div class="col-md-12">
				<h1>SSM-CRUD</h1>
			</div>
		</div>
		
		<!-- 按钮 -->
		<div class="row">
			<div class="col-md-4 col-md-offset-12">
				<button class="btn btn-primary">新增</button>
				<button class="btn btn-danger">删除</button>
			</div>
		</div>
		
		<!-- 显示表格数据 -->
		<div class="row"></div>
			<div class="col-md-12">
				<table class="table table-striped table-hover">
					<tr>
						<th>员工ID</th>
						<th>员工姓名</th>
						<th>性别</th>
						<th>邮箱</th>
						<th>部门名</th>
						<th>操作</th>
					</tr>
					<c:forEach items="${pageInfo.list }" var="emp">
					<tr>
						<th>${emp.empId} </th>
						<th>${emp.empName }</th>
						<th>${emp.gender=="M"?"男":"女"} </th>
						<th>${emp.email} </th>
						<th>${emp.department.deptName}</th>
						<th>
							<button class="btn btn-primary btn-xs">
							  <span class="glyphicon glyphicon-pencil"></span>
							     编辑
							</button>
							<button class="btn btn-danger btn-xs">
							<span class="glyphicon glyphicon-remove"></span>
							删除
							</button>
						</th>
					</tr>
					</c:forEach>
				</table>
			</div>
		<!-- 显示分页信息 -->
		<div class="row col-md-12">
		<!-- 分页信息 -->
		<div class="col-md-6">
				当前${pageInfo.pageNum }页，总共${pageInfo.pages }页，总共${pageInfo.total } 记录数 
		</div>
		<!-- 分页条 -->
		<div class="col-md-6">
				<nav aria-label="Page navigation">
						<ul class="pagination">
						<li><a href="${APP_PATH }/emps?pn=1">首页</a>
						<c:if test="${pageInfo.hasPreviousPage }">
				 		<li>
				  			 <a href="${APP_PATH }/emps?pn=${pageInfo.pageNum-1}" aria-label="Previous">
				     			<span aria-hidden="true">&laquo;</span>
				   			</a>
				 			</li>
						</c:if>
				 			<!-- 连续要显示的页码 -->
				 			<c:forEach items="${pageInfo.navigatepageNums }" var="page_Num">
				 				<c:if test="${page_Num==pageInfo.pageNum }">
				 					<li class="active"><a href="#">${page_Num } </a></li>
				 				</c:if>
				 				<c:if test="${page_Num!=pageInfo.pageNum }">
				 					<li><a href="${APP_PATH }/emps?pn=${page_Num }">${page_Num } </a></li>
				 				</c:if>
				 			</c:forEach>
				 			<c:if test="${pageInfo.hasNextPage }">
							 <li>
				   			<a href="${APP_PATH }/emps?pn=${pageInfo.pageNum-1}" aria-label="Next">
				     		<span aria-hidden="true">&raquo;</span>
				  			</a>
				 			</li>
				 			</c:if>
				 			<li><a href="${APP_PATH }/emps?pn=${pageInfo.pages}">尾页</a>
							</ul>
				 </nav>
		
		</div>
		</div>
		
    </div>

</body>
</html>
	
