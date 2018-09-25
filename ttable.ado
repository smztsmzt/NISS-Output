capture program drop ttable
program define ttable,rclass // 
/*
 Stata two-way pre-formatted table
 
 By: Pengpeng Ye
 Creation of Date: 2017-02-03
 Last revision of Date：2017-11-12
 
 Version 2.0 Stata14.1
 
	.适合格式化输出两个变量交叉表
	.可以生成列合计或行合计
	.可以自定义excel表名
	.可以自定义sheet名称
	.可以自定义在表中输出的起始行数、列数
	.可以自定义常用三线表三条线的粗细格式
	.可以自定义输出字体的类型、大小及颜色
	.可以自定义输出列名称
	.可以自定义合计名称
	.可以自定义是否间隔一列输出
	.可以自定义频数和构成输出的格式
	.可以自定义增加表格最后一行备注
	.可以自定义增加表格第一行表名
	
 Next plan:
 
	.增加无行列合计的输出
	.增加无频数，只有行合计或列合计的输出
	.可调整列合计输出的位置

	.增加是否显示缺失值选项
	.增加支持if选项
	.增加支持in选项
	.增加支持by选项

	
	.改写为program
	.支持多表格在同sheet中输出 :对于空表 只需每次将表格下移几行即可  对于非空表 
	 可以用import excel "",describle 返回的r(range_1)得到的最后数字加上一个数字从第a列开始输出
	
 References: 
  .tabdisp
  .tab3way
  .tw3xls
  .tab2xl
  .tab2xl2
  .tabout
  .logout
  .baselinetable
  .dyndoc
  .http://blog.stata.com/2017/01/10/creating-excel-tables-with-putexcel-part-1-introduction-and-formatting/
  .http://blog.stata.com/2017/01/24/creating-excel-tables-with-putexcel-part-2-macro-picture-matrix-and-formula-expressions/
*/

*! ttable v1.0 YPP 12DEC2017
*! ttable v1.1 YPP 17SEP2018

	version 14.1
	
	syntax varlist(max=2) [if] [in] [, missing]
	
	//tempvar touse
	//mark `touse' `if' `in'
	//global mi = cond("`missing'"~="","`missing'","")
    //tabulate `rowvar' `colvar' if `touse', $mi matrow(row) matcol(col) matcell(cell)  //cell为频数矩阵 上述可以替换
	
	count `if ' `in'
	if `r(N)' ==0 {
		error 2000
	}
	
	/*参数设置*/
	mata:mata clear
	matrix drop _all
	
	tokenize `varlist'
	local rowvar `1' //行变量，请自行设置
	macro shift
	local colvar `*' //列变量，请自行设置
	
	display "varlist contains |`varlist'|"
	display " if contains |`if'|"
	display " in contains |`in'|"
	display " missing contains |`missing'|"
	
	local tablename "table.xlsx" //Excel表名称，默认为table.xlsx，请自行设置
	local sheetname "Sheet1" //表单名称，默认为Sheet1，请自行设置

local stcol_freq=1 //输出的起始列数，默认为第1列，请自行设置
local strow_freq=1 //输出的起始行数，默认为第1行，请自行设置
local startcol=`stcol_freq'
local startcol_letter="A" //起始列数对应的字母，默认为A
mata: st_local("startcol_letter", numtobase26(`startcol'))
local startcol_letter2="A" //用于后期相同的列名称合并，默认为A

local firstline="medium" //三线表第一条横线，默认为medium，请自行设置，thin\medium\thick
local secondline="thin" //三线表第二条横线，默认thin，请自行设置，thin\medium\thick
local thirdline="medium" //三线表第二条横线，默认为medium，请自行设置，thin\medium\thick

local fontmat="Times New Roman" //字体格式，默认为Times New Roman，请自行设置，请按照excel已有的字体
local fontsize=11 //字体大小，默认为11号大小，请自行设置，请按照excel已有的大小
local fontcolor="black" //字体颜色，默认为black，请自行设置，请按照excel已有的颜色

local colname1="例次" //列名称1，默认为例次，请自行设置
local colname2="构成比(%)" //列名称2，默认为构成比(%)，请自行设置
local totalname1="合计" // 行合计名称，默认为合计，请自行设置
local totalname2="合计" //列合计名称，默认为合计，请自行设置

local percent="rowpercent" //构成比，默认为列构成，还可设置为行构成：rowpercent，colpercent_only，rowpercent_only，freq_only，

local interval=1 //开启间隔空列输出，关闭请设置为0，开启请设置为1

local ff1="#,###0"
local ff2="### ### ###"
local ff3="0"
local freqformat="`ff1'" //输出频数的格式，可自行增加设置

local pf1="0.00"
local pf2="#.00"
local percentformat="`pf1'" //输出构成的格式，可自行增加设置

local note_yn=1 //开启备注行，在表格最后一行
local note_content="注：" //可自行增加文字，关闭请设置为0，开启请设置为1

local title_yn=1 //开启表名行，在表格最前一行，合并居中
local title_content="表 " //可自行增加文字，关闭请设置为0，开启请设置为1
if `title_yn'==1 {
	local strow_freq=`strow_freq'+1 //输出的起始行数，开启表名以后，空一行留给表名位置
}

	
/*tabulate内置命令 获取相关矩阵*/
tabulate `varlist' `if' `in', matrow(row) matcol(col) matcell(cell) `missing' 
matrix list cell

/*设置写入的excel表名和sheet名*/
putexcel set `tablename',sheet(`sheetname') modify // replace

scalar firstrow_newtable=0
mata: mata clear
mata: st_numscalar("firstrow_newtable",emptyexcelyn("`tablename'","`sheetname'"))
local strow_freq=firstrow_newtable
di "`strow_freq'"


/*生成矩阵
cell 为主频数矩阵 mainfreq
rowi 为单位矩阵 
coli 为单位矩阵

allfreq 为全频数矩阵 含边际合计 
freq_colpercent 为列合计矩阵
freq_rowpercent矩阵 为行合计矩阵

allfreqcp 为全频数与列合计矩阵
allfreqrp 为全频数与行合计矩阵

null 为空矩阵
allfreqcp_interval 为全频数与列合计矩阵 含空列间隔
allfreqrp_interval 为全频数与行合计矩阵 含空列间隔

colname_firstline 为列变量名矩阵
*/

//allfreq矩阵
local row_mainfreq=rowsof(cell) //频数矩阵 行数
local col_mainfreq=colsof(cell) //频数矩阵 列数

matrix rowi_mainfreq=J(`row_mainfreq',1,1) //生成rowi单位矩阵，coltotal为频数矩阵的列合计矩阵
matrix coltotal_mainfreq=(cell'*rowi_mainfreq)' 

matrix coli_mainfreq=J(`col_mainfreq',1,1) //生成coli单位矩阵，rowtotal为频数矩阵的行合计矩阵
matrix rowtotal_mainfreq=cell*coli_mainfreq 

matrix sumtotal_mainfreq=coltotal_mainfreq*coli_mainfreq //sumtotal为总和矩阵，只有一个元素
matrix allfreq=(cell\coltotal_mainfreq),(rowtotal_mainfreq\sumtotal_mainfreq) //将频数矩阵、列合计矩阵、行合计矩阵，总和矩阵合并，allcell为频数全矩阵

matrix colnames allfreq="`colname1'" //为allfreq矩阵命令列名称

local row_allfreq=rowsof(allfreq)
local col_allfreq=colsof(allfreq)

//freq_colpercent矩阵
matrix cell_col=cell,rowtotal_mainfreq  //以下3步矩阵运算为构造列合计矩阵 
matrix rowi_freqcp=J(`row_mainfreq',1,1)
matrix freq_colpercent=allfreq*(syminv(diag((cell_col'*rowi_freqcp)'/100)))
matrix colnames freq_colpercent="`colname2'" //为freq_colpercent矩阵命令列名称


//freq_rowpercent矩阵
matrix cell_row=cell',coltotal_mainfreq' //以下3步矩阵运算为构造行合计矩阵 
matrix coli_freqrp=J(`col_mainfreq',1,1)
matrix freq_rowpercent=(allfreq'*(syminv(diag((cell_row'*coli_freqrp)/100))))'
matrix colnames freq_rowpercent="`colname2'" //为freq_rowpercent矩阵命令列名称


//合并allfreq与freq_colpercent，freq_rowpercent矩阵
matrix allfreqcp=J(`row_allfreq',1,.) //allfreq和col_percent合并，无间隔输出
matrix allfreqrp=J(`row_allfreq',1,.) //allfreq和row_percent合并，无间隔输出
matrix allfreqcp_interval=J(`row_allfreq',1,.) //allfreq和col_percent合并，有间隔输出
matrix allfreqrp_interval=J(`row_allfreq',1,.) //allfreq和row_percent合并，有间隔输出
matrix null=J(`row_allfreq',1,.) //空矩阵

forvalues i=1/`col_allfreq' {
	matrix allfreqcp=allfreqcp,allfreq[1...,`i'..`i'],freq_colpercent[1...,`i'..`i']
	matrix allfreqrp=allfreqrp,allfreq[1...,`i'..`i'],freq_rowpercent[1...,`i'..`i']
	matrix allfreqcp_interval=allfreqcp_interval,allfreq[1...,`i'..`i'],freq_colpercent[1...,`i'..`i'],null
	matrix allfreqrp_interval=allfreqrp_interval,allfreq[1...,`i'..`i'],freq_rowpercent[1...,`i'..`i'],null
}
matrix allfreqcp=allfreqcp[1...,2...] //去除第一列空值
matrix allfreqrp=allfreqrp[1...,2...] //去除第一列空值
matrix allfreqcp_interval=allfreqcp_interval[1...,2..`=`col_allfreq'*3'] //去除第一列和最后一列空值
matrix allfreqrp_interval=allfreqrp_interval[1...,2..`=`col_allfreq'*3'] //去除第一列和最后一列空值

forvalues i=1/`row_mainfreq' { //为所有矩阵命名行名称，暂不采用levelsof命令，因为变量值标签存在顺序问题
	local val = row[`i',1]
	local val_lab : label (`rowvar') `val' //宏取变量值标签，还可以利用decode 和 levelsof获取名称
	local rownames `rownames' `val_lab'
}
local rownames `rownames' "`totalname1'" //追加合计名称
matrix rownames freq_colpercent=`rownames'
matrix rownames freq_rowpercent=`rownames'
matrix rownames allfreq=`rownames'
matrix rownames allfreqcp=`rownames'
matrix rownames allfreqrp=`rownames'
matrix rownames allfreqcp_interval=`rownames'
matrix rownames allfreqrp_interval=`rownames'


//列变量名称矩阵

if `interval'==1 { //开启间隔输出模式
	matrix colname_firstline=J(1,`=`col_allfreq'*3-1',.) //开启间隔输出以后，删除合计后面多余一列
}
else {
 	matrix colname_firstline=J(1,`=`col_allfreq'*2',.) 
}

forvalues i=1/`col_mainfreq' { //为所有矩阵命名行名称，暂不采用levelsof命令，因为变量值标签存在顺序问题
	local val = col[1,`i']
	local colval_lab : label (`colvar') `val' //宏取变量值标签，还可以利用decode 和 levelsof获取名称
	
	if `interval'==1 { //开启间隔输出模式
		local colnames `colnames' `colval_lab' `colval_lab' `"` '"' //开启间隔输出以后，增加间隔列
	}
	else {
		local colnames `colnames' `colval_lab' `colval_lab' 
	}
}
 
local colnames `colnames' "`totalname2'" "`totalname2'" //追加合计名称	
matrix colnames colname_firstline=`colnames'


/*输出excel*/

//输出第一行第一列 行变量名称
local startrow=`strow_freq' //起始行数
local rowvar_lab : variable label `rowvar' //宏获取变量标签
mata: st_local("startcol_letter", numtobase26(`startcol')) 
putexcel `startcol_letter'`startrow':`startcol_letter'`=`startrow'+1'="`rowvar_lab'", ///
          merge vcenter hcenter font("`fontmat'",`fontsize',`fontcolor') 

		  
//输出第一行 各列变量名称

local startcol=`startcol'+1 //列位置加1
mata: st_local("startcol_letter", numtobase26(`startcol')) 
putexcel `startcol_letter'`startrow'=matrix(colname_firstline),colnames vcenter hcenter font("`fontmat'",`fontsize',`fontcolor') 


forvalues i=1/`col_allfreq' { //合并相同的列名称
	
	mata: st_local("startcol_letter2", numtobase26(`=`startcol'+1')) //此处加1是列名称横向合并的单元格
	quietly putexcel `startcol_letter'`startrow':`startcol_letter2'`startrow',merge
	
	if `interval'==1 { //开启间隔输出模式
		mata: st_local("startcol_letter2", numtobase26(`=`startcol'+2')) //此处加2是消除空列的列名称
		quietly putexcel `startcol_letter2'`startrow':`startcol_letter2'`=`startrow'+1'="",merge
		mata: st_local("startcol_letter2", numtobase26(`=`startcol'+1')) //此处加1是恢复位置
		local  startcol=`startcol'+3 //此处加3是转至下一个列名称所在的单元格 
	}
	else {
		local  startcol=`startcol'+2 //此处加2是转至下一个列名称所在的单元格 
	}
	mata: st_local("startcol_letter", numtobase26(`startcol'))		
}

//输出矩阵
local startcol=`stcol_freq'
mata: st_local("startcol_letter", numtobase26(`startcol')) 
if (`interval'!=1)&("`percent'"=="colpercent") {
	quietly putexcel `startcol_letter'`=`startrow'+1'=matrix(allfreqcp),names vcenter hcenter font("`fontmat'",`fontsize',`fontcolor') nformat(`percentformat')
}
else if (`interval'!=1)&("`percent'"=="rowpercent") {
	quietly putexcel `startcol_letter'`=`startrow'+1'=matrix(allfreqrp),names vcenter hcenter font("`fontmat'",`fontsize',`fontcolor') nformat(`percentformat')
}
else if (`interval'==1)&("`percent'"=="colpercent") { //开启间隔输出模式
	quietly putexcel `startcol_letter'`=`startrow'+1'=matrix(allfreqcp_interval),names vcenter hcenter font("`fontmat'",`fontsize',`fontcolor') nformat(`percentformat')
}
else if (`interval'==1)&("`percent'"=="rowpercent") { //开启间隔输出模式
	quietly putexcel `startcol_letter'`=`startrow'+1'=matrix(allfreqrp_interval),names vcenter hcenter font("`fontmat'",`fontsize',`fontcolor') nformat(`percentformat')
}


//输出表格框线
quietly putexcel `startcol_letter'`startrow':`startcol_letter2'`startrow',border(top,`firstline',black) 
quietly putexcel `startcol_letter'`startrow':`startcol_letter2'`startrow',border(bottom,`secondline',black) 
quietly putexcel `startcol_letter'`=`startrow'+1':`startcol_letter2'`=`startrow'+1',border(bottom,`secondline',black) 
quietly putexcel `startcol_letter'`=(`startrow'+`row_allfreq'+1)':`startcol_letter2'`=(`startrow'+`row_allfreq'+1)',border(bottom,`thirdline',black) 

di "ypp"
/*
调整表格数值格式 因为矩阵是一次性输出，所以无法同时满足频数（整数）与构成（小数）的显示格式，
这里再次调整整数显示格式，会增加程序操作excel的运行时间，但减少后期手工调整格式负担
*/
local startcol=`startcol'+1
mata: st_local("startcol_letter", numtobase26(`startcol')) 
forvalues i=1/`col_allfreq' {
	quietly putexcel `startcol_letter'`=`startrow'+2':`startcol_letter'`=(`startrow'+1+`row_allfreq')',nformat(`freqformat')
	
	if `interval'==1 { //开启间隔输出模式
		local startcol=`startcol'+3
		mata: st_local("startcol_letter", numtobase26(`startcol')) 
	}
	else {
		local startcol=`startcol'+2
		mata: st_local("startcol_letter", numtobase26(`startcol')) 
	}
}
local startcol=`stcol_freq' //回复到起始列数
mata: st_local("startcol_letter", numtobase26(`startcol')) 

//开启备注行
if `note_yn'==1 {
	di "开启注释行"
	quietly putexcel `startcol_letter'`=(`startrow'+2+`row_allfreq')'="`note_content'",vcenter hcenter font("`fontmat'",`fontsize',`fontcolor')
}

//开启表名行
if `title_yn'==1 {
	di "开启表名行"
	quietly putexcel `startcol_letter'`=`startrow'-1':`startcol_letter2'`=`startrow'-1'="`title_content'",merge vcenter hcenter font("`fontmat'",`fontsize',`fontcolor')
}
di "最终行数位置：" `startrow' 
di "最终列数位置：" `startcol'
di "最终单元格位置：" "`startcol_letter'"`startrow'
di "最终单元格位置：" "`startcol_letter2'"`startrow'
di "row_mainfreq：" "`row_mainfreq'"
di "col_mainfreq：" "`col_mainfreq'"
di "row_allfreq：" "`row_allfreq'"
di "col_allfreq：" "`col_allfreq'"
di "stcol_freq：" "`stcol_freq'"
di "strow_freq：" "`strow_freq'"

end

/* 可以通过矩阵运算合并，但无法保留列名，暂不使用
matrix cx_allfreq=J(`col_allfreq',`=`col_allfreq'*2',0)
forvalues i=1/`col_allfreq' {
	matrix cx_allfreq[`i',`=2*`i'-1']=1
}
matrix cx_allfreq=allfreq*cx_allfreq

matrix cx_freqcp=J(`col_allfreq',`=`col_allfreq'*2',0)
forvalues i=1/`col_allfreq' {
	matrix cx_freqcp[`i',`=2*`i'']=1
}
matrix cx_freqcp=freq_colpercent*cx_freqcp

matrix allfreqcp=cx_allfreq+cx_freqcp
matrix list allfreqcp
*/

mata:
real emptyexcelyn(string scalar tablename,string scalar sheetname)
{

	class xl scalar   b 		// 初始化xl()类
	string vector     cells     // 声明字符串向量
	real scalar       i			// 判断指定区域的长度
	real scalar 	  r1		// 指定区域行起始位置
	real scalar       r2		// 指定区域行终末位置
	real vector       rowstart  // 声明实数向量 行区域
	real vector       colstart  // 声明实数向量 列区域
	
	r1=2						// 行起始位置为1
	r2=4						// 行终末位置为2
	rowstart=(r1\r2)			// 行区域初始化为第1和2行
	colstart=1					// 列位置初始化为第1列
	
	b.load_book(tablename)		// 加载表名
	b.set_sheet(sheetname)      // 加载sheet名
	cells=b.get_string(rowstart,colstart) // 获取指定位置的内容
	i=colsum(ustrlen(cells))    // 判断指定区域的长度
	
	while (i) {					// 如果长度不为0
		r1=r1+2					// 则行区域递增2
		r2=r2+2					// 同上
		rowstart=(r1\r2)	    // 更新指定的区域	
		cells=b.get_string(rowstart,colstart) // 列区域不变
		i=colsum(ustrlen(cells)) // 判断新的指定区域的长度
	}

	return(r2)					 // 返回输出新表的起始行数 可以修改为r2
}
end

