<%
'''  数据库操作
class Cls_dbbase
	private rs
	public conn
	
	
	''打开数据库
	public sub dbopen()
		on error resume next
		if isobject(conn) then exit sub
		dim connstr
		if webroot="" then '未定义根目录
			Cls.echo "友情提示:请先执行安装程序"
			Cls.die
		end if
		if datatype then  '判断是否access
			connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath(webroot&datapath&"/"&datadbname)
		else
			connstr="Provider=Sqloledb;User ID="&datauser&";Password="&datapass&";Initial CataLog="&datasqlname&";Data Source="&datahost&";"
		end if
		set conn=server.createobject("adodb.connection")
		conn.open connstr
		if err then
			Cls.echo "数据库连接失败,请检查Config.asp内的虚拟目录参数webroot"
			err.clear
			Cls.die
		end if
	end sub
	
	''关闭数据库
	public sub dbclose()
		on error resume next
		if isobject(conn) or conn.state=1 then
			conn.close
			set conn=nothing
		end if
		err.clear
	end sub
	
	'单次执行查询
	public function exedb(byval t0)
		on error resume next
		set exedb=conn.execute(t0)
		if err then
			'err.clear
			'Cls.die
		end if
	end function
	
	'插入数据,t0为表名,t1为数据组
	public sub insert(byval t0,byval t1)
		dim dbfield,dbvalue,dbarray,i,temparr,fieldlen,filedtype,sqlvalue
		dbfield=""
		dbvalue=""
		dbarray=t1
		filedtype=""
		for i=0 to ubound(dbarray)
			temparr=dbarray(i)
			if len(dbfield)>0 then
				dbfield=dbfield&","
				dbvalue=dbvalue&","
			end if
			if instr(temparr(0),"(")>0 or instr(temparr(0),")")>0 or instr(temparr(0)," ")>0 or instr(temparr(0),".")>0 then
				dbfield=dbfield&temparr(0)
			else
				dbfield=dbfield&"["&temparr(0)&"]"
			end if
			fieldlen=temparr(2)
			filedtype=temparr(3)
			sqlvalue=replace(temparr(1),"'","''")
			if fieldlen=0 then
				if filedtype=1 then
					dbvalue=dbvalue&"'"&sqlvalue&"'"
				else
					dbvalue=dbvalue&sqlvalue
				end if
			else
				if filedtype=1 then
					dbvalue=dbvalue&"'"&left(sqlvalue,fieldlen)&"'"
				else
					dbvalue=dbvalue&left(sqlvalue,fieldlen)
				end if
			end if
		next
		exedb("insert into "&t0&" ("&dbfield&") values ("&dbvalue&")")
	end sub
	
	''带条件添加新数据,t0为表名,t1为数据组，t2为where判断
	public function dbnew(byval t0,byval t1,byval t2)
		dim dbfield,dbvalue,dbnum,dbarray,i,temparr,fieldlen,sqlvalue
		dbfield=""
		dbvalue=""
		dbarray=t1
		dbnew=0
		dbnum=ubound(dbarray)
		if len(t2)>0 then t2="where "&t2
		dim fieldarr(999)
		for i=0 to dbnum
			temparr=dbarray(i)
			if len(dbfield)>0 then
				dbfield=dbfield&","
			end if
			fieldlen=temparr(2)
			sqlvalue=temparr(1)
			if fieldlen=0 then
				fieldarr(i)=sqlvalue
			else
				fieldarr(i)=left(sqlvalue,fieldlen)
			end if
			if instr(temparr(0),"(")>0 or instr(temparr(0),")")>0 or instr(temparr(0)," ")>0 or instr(temparr(0),".")>0 then
				dbfield=dbfield&temparr(0)
			else
				dbfield=dbfield&"["&temparr(0)&"]"
			end if
		next

		set rs=server.createobject("adodb.recordset")
		rs.open "select "&dbfield&" from "&t0&" "&t2&"",conn,1,3
		if len(t2)>0 then
			if not rs.eof then
				rs.close
				set rs=nothing
				exit function
			end if
		end if
		rs.addnew
		for i=0 to dbnum
			rs(i)=fieldarr(i)
		next
		rs.update
		rs.close
		set rs=nothing
		dbnew=1
	end function
	
	''更新数据,t0为表名,t1为数据组，t2为where判断
	public function dbupdate(byval t0,byval t1,byval t2)
		dbupdate=0
		dim i,dbarray,sqlvalue
		dim n1,n2,n3,n4
		dbarray=t2
		sqlvalue=""
		for i=0 to ubound(dbarray)
			if len(sqlvalue)>0 then
				sqlvalue=sqlvalue&","
			end if
			 n1=dbarray(i)(0)
			 n2=dbarray(i)(1)
			 n3=dbarray(i)(2)
			 n4=dbarray(i)(3)
			 if n3<>0 then n2=left(n2,n3)
			 if n4=1 then n2=Cls.sqlstr(n2)
			 if instr(n1,"(")>0 or instr(n1,")")>0 or instr(n1," ")>0 or instr(n1,".")>0 then
				  sqlvalue=sqlvalue&n1&"="&n2&""
			 else
				sqlvalue=sqlvalue&"["&n1&"]="&n2&""
			 end if
		next
		dim sql
		sql="update "&t0&" set "&sqlvalue
		if len(t1)>0 then sql=sql&" where "&t1
		exedb(sql)
		dbupdate=1
	end function

	''查询数据库，t0为top,t1为字段,t2为表名,t3为where条件，t4为排序
	public function dbload(byval t0,byval t1,byval t2,byval t3,byval t4)
		if len(t0)>0 then t0=" top "&t0
		if len(t1)=0 then t1=" * "
		if len(t3)>0 then t3=" where "&t3
		if left(lcase(t4),3)="rnd" then
			if datatype then
				randomize
				dim orderid
				if len(t4)>4 then orderid=right(t4,len(t4)-4)
				if len(orderid)=0 then orderid="id"
				t4="rnd(-("&orderid&" +"&rnd()&"))"
			else
				t4="newid()"
			end if
		end if
		if len(t4)>0 then t4=" order by "&t4
		if t1<>" * " then
			dim str:str=""
			dim arr:arr=split(t1,",")
			dim i
			for i=0 to ubound(arr)
				if i>0 then str=str&","
				if instr(arr(i),"(")>0 or instr(arr(i),")")>0 or instr(arr(i)," ")>0 or instr(arr(i),".")>0 then
					str=str&arr(i)
				else
					str=str&"["&arr(i)&"]"
				end if
			next
			t1=str
		end if
		set rs=exedb("select "&t0&" "&t1&" from "&t2&" "&t3&" "&t4)
		if rs.eof then
			dbload=array()
		else
			dbload=rs.getrows()
		end if
		rs.close
		set rs=nothing
	end function
	
	''单次查询数据库，t0为字段，t1为表名,t2为where条件
	public function dbloadone(byval t0,byval t1,byval t2)
		if len(t2)>0 then t2=" where "&t2
		dbloadone=exedb("select "&t0&" from "&t1&" "&t2&"")(0)
	end function
	
	''删除数据,t0为表名,t1为where条件
	public sub dbdel(byval t0,byval t1)
		dim sql
		sql="delete from "&t0
		if len(t1)>0 then sql=sql&" where "&t1
		exedb(sql)
	end sub
	
	''计算数据数量,t0为表名,t1为where条件
	public function dbcount(byval t0,byval t1)
		dim sql
		sql="select count(1) from "&t0
		if len(t1)>0 then sql=sql&" where "&t1
		dbcount=exedb(sql)(0)
	end function
	
	''查询最新插入的ID，t0为字段，t1为表名
	public function insertid(byval t0,byval t1)
		'if datatype then
			insertid=exedb("select max("&t0&") from "&t1&"")(0)
		'else
			'insertid=exedb("select @@identity")
		'end if
	end function
	
	''查询最新插入的ID，t0为字段，t1为表名，t2为where条件
	public function insertid_new(byval t0,byval t1,byval t2)
		if len(t2)>0 then t2=" where "&t2
		insertid_new=exedb("select max("&t0&") from "&t1&" "&t2&"")(0)
	end function
end class
%>