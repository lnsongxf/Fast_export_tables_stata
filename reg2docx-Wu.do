*--------------------------------------------------------
* 
*       君生我未生！Stata - 论文四表一键出
*
*-简书原文链接：http://www.jianshu.com/p/97c4f291ee1e
*
* 一步到位，两步见表！只要两步，输出实证论文四张基本表格
*
*--------------------------------------------------------
*
*-Author: 吴水亭(广西财经学院)
*-Date: 2017/11/6 10:45


*=============
*-1 写在前面
*=============

* 描述性统计、相关系数矩阵、组间均值差异检验和回归结果四张常用表格如何快速输出?
* 平时我们常用“logout”, "tabout", "esttab"等命令，但各有一些不便捷的问题。

help logout    //输出统计表格或列式矩阵
help tabout    //输出统计表格或列式矩阵
help esttab    //输出估计结果

* STATA15 四个热门外部命令：“sum2docx","corr2docx","t2docx"和”reg2docx"帮你解决问题
* 其中：

help sum2docx  //将描述性统计量表直接输出到一个 docx 文件中；
help corr2docx //将相关系数直接输入到一个 docx 文件中；
help t2docx    //将分组均值t检验的结果导出到一个 docx 文件中；
help reg2docx  //可以将回归结果保存到 docx 文件中，用法类似于 esttab 。


*===============
*-2 下载及安装
*===============

* 在 Stata 命令窗口中输入 ssc hot 命令，会显示出最热门的 10 个外部命令。
ssc hot 
* 下载只需在 Stata 命令窗口执行 ssc install **2docx, replace 即可。
* 以下载 sum2docx 命令为例： 
ssc install sum2docx, replace //附加 replace 选项, 下载最新版


*=====================================
*-3 输出四张常用基本表格至 Word 文档
*=====================================

*---------------------------------
*-3.1 输出基本统计量：sum2docx 命令

*-3.1.1 语法结构

* --------------------------------------------------------
*   sum2docx varlist [if] [in] using filename , [options]  
* --------------------------------------------------------

* 其中，varlist指数值型变量列表，filename指的是输出的文件名

*-3.1.2 范例

sysuse "auto.dta", clear
sum2docx price-foreign using "Table01.docx", replace  ///
         obs mean(%9.2f) sd min(%9.0g) median(%9.0g) max(%9.0g) /// 
         title("表 1: 描述性统计")
shellout "Table01.docx" // 从 Stata 中打开 Word 文档

*-------------------------------------
*-3.2 输出相关系数矩阵：corr2docx 命令

*-3.2.1语法结构

* --------------------------------------------------------
*   corr2docx varlist[if] [in] using filename, [options]  
* --------------------------------------------------------

* 其中，varlist 指数值型变量列表，filename 指的是输出的文件名

*-3.2.2 范例

sysuse "auto.dta", clear
corr2docx price-foreign using "Table02.docx", star(* 0.05) fmt(%4.2f)  ///
          title("表 2：相关系数矩阵") 
shellout "Table02.docx"
*-Notes: 
*-(1) star(* 0.05) 表示仅在 5% 以上显著的系数旁打一个星号;
*-(2) fmt(%4.2f) 表示小数点后保留两位有效数字

*-------------------------------------
*-3.3 输出组间均值差异检验：t2docx 命令

*-3.3.1 语法结构

* --------------------------------------------------------
*   t2docx varlist[if] [in] using filename, [options]
* --------------------------------------------------------

*-3.3.2 范例

sysuse "auto.dta", clear
t2docx price weight length mpg using "Table03.docx", replace by(foreign) ///
       title("表 3：t检验")
shellout "Table03.docx"  //当然也可以改变小数点位数，加星星啥的，和上面一样一样的。


*--------------------------------
*-3.4 输出回归结果：reg2docx 命令

* 内啥，我们想把几个回归结果何在一张表里，能整不？

sysuse "auto.dta", clear

* 接下来做两个回归，并将结果放在一个表中，输出至 4.docx
* 比如先做一个线性回归

reg price mpg weight length
est store m1

reg price mpg weight length foreign
est store m2

* 然后再做一个Probit回归

probit foreign price weight length
est store m3

reg2docx m1 m2 m3 using "Table04.docx", replace ///
         r2(%9.3f) ar2(%9.2f) b(%9.3f) t(%7.2f) ///
         title("表4: 回归结果")
shellout "Table04.docx"


*=======================================
*-4 将上述四张表输出至一个 Word 文档中
*=======================================

*-思路：
*-(1) 用 putdocx 命令生成一个空白 Word 文档 - [My_Table.docx]，
*     进而使用 putdocx text 等命令设定文档属性;
*-(2) 用 sum2docx 生成 [表1], 并使用 sum2docx 命令的 append 选项将这张表
*     追加到 [My_Table.docx] 文档尾部;
*-(3) 按第二步的方法, 
*     依次使用 corr2docx, t2docx, reg2docx 命令添加后续表格

*-写作建议:
*
* 写论文正文时，新建一个 Word 文档, 命名为:     [My_Paper.docx]
*   在需要插入表格的地方写上
*
*     [---------Table # Here--------]
* 
* Stata 自动生成的表格则存放于另一个 Word 文档：[My_Table.docx] 
*    里面存放 [Table 1], [Table 2], .....



*  #  | 文档类型   |   文档名称     | 文档内容
*-----|------------|-----------------|------------------------
*  1  | 论文主文档 |  My_Paper.docx  |  引言,文献综述,理论分析等
*  2  | 表格文档   |  My_Table.docx  |  统计和回归表格

*------------------------------------------------------
*-------------------My_Paper.docx---------------beign--
*
*  一、引言
*  二、文献综述
*  三、研究假设
*  四、研究设计
*  五、结果及分析
*
*     x x x xx x xxx xx 
*
*        [----Table 1 here----]
*
*     xxxxxxxxxxxxxxxxxxxxxx
*
*        [----Table 2 here----]
*
*  六、结论
*  参考文献

*         -------------new page--------
*
*  附：文中表格
*      [------贴入 My_Table.docx 文档中的表格------]
*
*
*-------------------My_Paper.docx---------------over--
*------------------------------------------------------

clear all
set more off
putdocx begin                     //新建 Word 文档
putdocx paragraph, halign(center) //段落居中
*-定义字体、大小等基本设置
putdocx text ("附：文中待插入表格"), ///
        font("华为楷体",16,black) bold linebreak  
*-保存名为 My_Table.docx 的 Word 文档		
putdocx save "My_Table.docx", replace 

sysuse "auto.dta", clear 

*-Table 1
sum2docx price-length using "My_Table.docx", append ///
         obs mean(%9.2f) sd min(%9.0g) median(%9.0g) max(%9.0g) ///
         title("表 1: 描述性统计")
*-Note: 选项 append 的作用是将这张新表追加到 "My_Table.docx" 尾部, 下同.

*-Table 2
putdocx begin
putdocx pagebreak
putdocx save "My_Table.docx", append 
corr2docx price-length using "My_Table.docx", append ///
          star(* 0.05) fmt(%4.2f) ///
          title("表 2：相关系数矩阵") 

*-Table 3
putdocx begin
putdocx pagebreak
putdocx save "My_Table.docx", append
t2docx price-length using "My_Table.docx", append ///
       by(foreign) title("表 3：组间均值差异 t 检验")

*-Table 4
putdocx begin
putdocx pagebreak
putdocx save "My_Table.docx", append

reg price mpg weight length
est store m1
reg price mpg weight length foreign
est store m2
probit foreign price weight length
est store m3
reg2docx m1 m2 m3 using "My_Table.docx", append  ///
         r2(%9.3f) ar2(%9.2f) b(%9.3f) t(%7.2f) ///
		 title("表4: 回归结果")

shellout "My_Table.docx"  //大功告成！打开生成的 Word 文档
