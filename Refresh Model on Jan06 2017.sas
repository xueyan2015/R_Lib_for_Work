
%let path = %str(D:\Project_Files2\CN PMR\百多邦\Jan01_2017);

libname datadir2 "&path";

/*proc import out=datadir2.basic_info*/
/*datafile = "&path\百多邦项目140家门店销售数据_2016年4月-11月_Read_IN.xlsx"*/
/*dbms=excel replace;*/
/*range= "基本情况$A3:M143";*/
/*getnames = yes;*/
/*run;*/
/*proc import out=datadir2.bai_duo_bang*/
/*datafile = "&path\百多邦项目140家门店销售数据_2016年4月-11月_Read_IN.xlsx"*/
/*dbms=excel replace;*/
/*range= "百多邦销售数据$A3:AM144";*/
/*getnames = yes;*/
/*run;*/
/*proc import out=datadir2.pi_de_bang*/
/*datafile = "&path\百多邦项目140家门店销售数据_2016年4月-11月_Read_IN.xlsx"*/
/*dbms=excel replace;*/
/*range= "匹得邦销售数据$A3:AM144";*/
/*getnames = yes;*/
/*run;*/
/*proc import out=datadir2.jing_wan_hong*/
/*datafile = "&path\百多邦项目140家门店销售数据_2016年4月-11月_Read_IN.xlsx"*/
/*dbms=excel replace;*/
/*range= "京万红销售数据$A3:AM144";*/
/*getnames = yes;*/
/*run;*/
/**/
/*data datadir2.bai_duo_bang;*/
/*set datadir2.bai_duo_bang;*/
/*if City NE "占位符";*/
/*run;*/
/*data datadir2.pi_de_bang;*/
/*set datadir2.pi_de_bang;*/
/*if City NE "占位符";*/
/*run;*/
/*data datadir2.jing_wan_hong;*/
/*set datadir2.jing_wan_hong;*/
/*if City NE "占位符";*/
/*run;*/


/*----start processing raw data-----*/
%include "D:\Project_Files2\China\code_lib\sas_macro\ForEach.sas";
%macro process_raw_data(raw_data, product_name, start_mth, end_mth,  smaller_dose, larger_dose, pre_period_mth_count=6);

data temp;
do i = &start_mth to &end_mth;
	mth_id = put(i, z2.) ;
	output;
end;
run;

proc sql;
select mth_id into: mth_id_list separated by " " from temp;quit;

/*with no prefix 0*/

%let pre_period_end = %eval( &start_mth + &pre_period_mth_count - 1);
%let post_period_start = %eval(&start_mth + &pre_period_mth_count);

/*with prefix 0*/
%let pre_period_start = %sysfunc(putn(&start_mth, z2.));
%let pre_period_end = %sysfunc(putn(&pre_period_end, z2.));
%let post_period_start = %sysfunc(putn(&post_period_start, z2.));
%let post_period_end = %sysfunc(putn(&end_mth, z2.));

%let prefix_01A =  Sales_volume_&smaller_dose.  ;
%let prefix_01B=    Sales_volume_&larger_dose.  ;

%let prefix_02A=  Sales_value_&smaller_dose.;
%let prefix_02B=  Sales_value_&larger_dose.;

%let volume_smaller_dose = %foreach(v,&mth_id_list, %nrstr(&prefix_01A._&v));
%let volume_larger_dose = %foreach(v,&mth_id_list, %nrstr(&prefix_01B._&v));
%let value_smaller_dose = %foreach(v,&mth_id_list, %nrstr(&prefix_02A._&v));
%let value_larger_dose = %foreach(v,&mth_id_list, %nrstr(&prefix_02B._&v));

%put &volume_smaller_dose;
%put &volume_larger_dose;
%put &value_smaller_dose;
%put &value_larger_dose;

data &product_name._sum ;

retain Pharma_Store_ID
City
Chain_or_Single
&volume_smaller_dose
&volume_larger_dose
&value_smaller_dose
&value_larger_dose

;
set &raw_data;

array s1 &prefix_01A._&pre_period_start - &prefix_01A._&pre_period_end;
array s2 &prefix_01B._&pre_period_start - &prefix_01B._&pre_period_end;
array s3 &prefix_02A._&pre_period_start - &prefix_02A._&pre_period_end;
array s4 &prefix_02B._&pre_period_start - &prefix_02B._&pre_period_end;

array s5 &prefix_01A._&post_period_start - &prefix_01A._&post_period_end;
array s6 &prefix_01B._&post_period_start - &prefix_01B._&post_period_end;
array s7 &prefix_02A._&post_period_start - &prefix_02A._&post_period_end;
array s8 &prefix_02B._&post_period_start - &prefix_02B._&post_period_end;

&prefix_01A._pre_test = sum(of &prefix_01A._&pre_period_start - &prefix_01A._&pre_period_end);
&prefix_01B._pre_test = sum(of &prefix_01B._&pre_period_start - &prefix_01B._&pre_period_end);
&prefix_02A._pre_test = sum(of &prefix_02A._&pre_period_start - &prefix_02A._&pre_period_end);
&prefix_02B._pre_test = sum(of &prefix_02B._&pre_period_start - &prefix_02B._&pre_period_end);

&prefix_01A._post_test = sum(of &prefix_01A._&post_period_start - &prefix_01A._&post_period_end);
&prefix_01B._post_test = sum(of &prefix_01B._&post_period_start - &prefix_01B._&post_period_end);
&prefix_02A._post_test = sum(of &prefix_02A._&post_period_start - &prefix_02A._&post_period_end);
&prefix_02B._post_test = sum(of &prefix_02B._&post_period_start - &prefix_02B._&post_period_end);

&prefix_01A._pre_missing = 0;
&prefix_01B._pre_missing = 0;
&prefix_02A._pre_missing  = 0;
&prefix_02B._pre_missing  = 0;

&prefix_01A._post_missing  = 0;
&prefix_01B._post_missing  = 0;
&prefix_02A._post_missing  = 0;
&prefix_02B._post_missing  = 0;


do over s1;
	
		if s1 =. then  &prefix_01A._pre_missing  = &prefix_01A._pre_missing  + 1;
		if s2 =.  then &prefix_01B._pre_missing  = &prefix_01B._pre_missing  + 1;
		if s3 =.  then &prefix_02A._pre_missing  = &prefix_02A._pre_missing  + 1;
		if s4 =.  then &prefix_02B._pre_missing  = &prefix_02B._pre_missing  + 1;

end;

do over s5;

		if s5 =. then &prefix_01A._post_missing  = &prefix_01A._post_missing   +1 ;
		if s6=.  then &prefix_01B._post_missing  = &prefix_01B._post_missing   + 1 ;
		if s7=.  then &prefix_02A._post_missing  = &prefix_02A._post_missing   + 1 ;
		if s8=.  then &prefix_02B._post_missing  = &prefix_02B._post_missing   + 1 ;

end;

run;

%mend process_raw_data;
options mprint;
%process_raw_data(datadir2.Bai_duo_bang, Bai_duo_bang, 4, 12, 5g, 10g);
%process_raw_data(datadir2.Pi_de_bang, Pi_de_bang, 4, 12, 5g, 10g);
%process_raw_data(datadir2.Jing_wan_hong, Jing_wan_hong, 4, 12, 20g, 50g);


%macro merge_data(product_name);

proc sort data = datadir2.Basic_info; by Pharma_Store_ID; run;
proc sort data=&product_name._sum; by Pharma_Store_ID; run;

data &product_name._merged;
merge datadir2.Basic_info(in = a)   &product_name._sum(in = b  drop = City Chain_or_Single );
if a and b;
run;

%mend merge_data;

%merge_data(Bai_duo_bang);

%let post_test_nmth_available = 2 ;

/*test ANCOVA*/
data Bai_duo_bang_model;
set Bai_duo_bang_merged;

/*Option_Position_Dummy = (Option_Position="A类位置");*/
/*Option_Product_Range_Dummy =  (Option_Product_Range = "5g+10g") ;*/
/*Option_Shining_Dummy =  ( Option_Shining = "2 个陈列面");*/
/*Option_POSM_Dummy =  (Option_POSM = "有");*/

/*Chain_or_Single_Dummy = (Chain_or_Single = "连锁");*/
Chai_or_Single_Combined = Chain_or_Single;
if Chai_or_Single_Combined NE "连锁" then Chai_or_Single_Combined = "非连锁";

/*note: 10g is converted to 5g equivalent units*/
sales_volume_pre_test = sum(Sales_volume_5g_pre_test, Sales_volume_10g_pre_test * 2)/6;
sales_volume_post_test = sum(Sales_volume_5g_post_test, Sales_volume_10g_post_test * 2)/&post_test_nmth_available.;  /*note: will need to be dvived by number of months when data are complete for post-period*/

sales_value_pre_test = sum(Sales_value_5g_pre_test, Sales_value_10g_pre_test)/6;
sales_value_post_test = sum(Sales_value_5g_post_test, Sales_value_10g_post_test)/&post_test_nmth_available.;  /*note: will need to be dvived by number of months when data are complete for post-period*/


avg_sales_volume_5g_pre = Sales_volume_5g_pre_test/6;
avg_sales_volume_5g_post = Sales_volume_5g_post_test/&post_test_nmth_available.;
avg_sales_value_5g_pre = Sales_value_5g_pre_test/6;
avg_sales_value_5g_post = Sales_value_5g_post_test/&post_test_nmth_available.;

avg_sales_volume_10g_pre = Sales_volume_10g_pre_test/6;
avg_sales_volume_10g_post = Sales_volume_10g_post_test/&post_test_nmth_available.;
avg_sales_value_10g_pre = Sales_value_10g_pre_test/6;
avg_sales_value_10g_post = Sales_value_10g_post_test/&post_test_nmth_available.;



run;

data Bai_duo_bang_model_reset;
set Bai_duo_bang_model;

array pre_sales   sales_volume_pre_test   sales_value_pre_test  avg_sales_volume_5g_pre  avg_sales_value_5g_pre  avg_sales_volume_10g_pre  avg_sales_value_10g_pre;
array post_sales  sales_volume_post_test   sales_value_post_test  avg_sales_volume_5g_post  avg_sales_value_5g_post  avg_sales_volume_10g_post  avg_sales_value_10g_post; 

do over pre_sales;

if pre_sales = . then pre_sales=0;

if post_sales=. then post_sales=0;

end;

run;


/*check 5g sales and 10g sales missing*/
proc sql;
select nmiss(avg_sales_volume_5g_pre)
, nmiss(avg_sales_volume_5g_post)
, nmiss(avg_sales_value_5g_pre)
, nmiss(avg_sales_value_5g_post)
, nmiss(avg_sales_volume_10g_pre)
, nmiss(avg_sales_volume_10g_post)
, nmiss(avg_sales_value_10g_pre)
, nmiss(avg_sales_value_10g_post)

, nmiss(sales_volume_pre_test)
, nmiss(sales_volume_post_test)

, nmiss(sales_value_pre_test)
, nmiss(sales_value_post_test)
from Bai_duo_bang_model;quit;



proc sort data=Bai_duo_bang_model_reset; by Pharma_Store_ID; run;
proc sort data=Pi_de_bang_sum; by Pharma_Store_ID; run;
proc sort data=Jing_wan_hong_sum; by Pharma_Store_ID; run; 

data Bai_duo_bang_model_2;

merge 

Bai_duo_bang_model_reset(in = a)  

Pi_de_bang_sum(in =b  
keep = Pharma_Store_ID   Sales_volume_5g_pre_test  Sales_volume_10g_pre_test  Sales_value_5g_pre_test  Sales_value_10g_pre_test  Sales_volume_5g_post_test  Sales_volume_10g_post_test  Sales_value_5g_post_test  Sales_value_10g_post_test
rename=(Sales_volume_5g_pre_test= pi_Sales_volume_5g_pre_test  
Sales_volume_10g_pre_test = pi_Sales_volume_10g_pre_test
Sales_value_5g_pre_test = pi_Sales_value_5g_pre_test
Sales_value_10g_pre_test = pi_Sales_value_10g_pre_test
Sales_volume_5g_post_test = pi_Sales_volume_5g_post_test
Sales_volume_10g_post_test = pi_Sales_volume_10g_post_test
Sales_value_5g_post_test = pi_Sales_value_5g_post_test
Sales_value_10g_post_test = pi_Sales_value_10g_post_test))  

Jing_wan_hong_sum(in=c  
keep= Pharma_Store_ID  Sales_volume_20g_pre_test  Sales_volume_50g_pre_test  Sales_value_20g_pre_test  Sales_value_50g_pre_test  Sales_volume_20g_post_test  Sales_volume_50g_post_test  Sales_value_20g_post_test  Sales_value_50g_post_test  
rename = (Sales_volume_20g_pre_test  = ji_Sales_volume_20g_pre_test
Sales_volume_50g_pre_test  = ji_Sales_volume_50g_pre_test
Sales_value_20g_pre_test  = ji_Sales_value_20g_pre_test
Sales_value_50g_pre_test  = ji_Sales_value_50g_pre_test
Sales_volume_20g_post_test  = ji_Sales_volume_20g_post_test
Sales_volume_50g_post_test  = ji_Sales_volume_50g_post_test
Sales_value_20g_post_test  = ji_Sales_value_20g_post_test
Sales_value_50g_post_test  = ji_Sales_value_50g_post_test));
by  Pharma_Store_ID;
if a and b and c;

pi_sales_volume_pre_test = sum(pi_Sales_volume_5g_pre_test ,  2 * pi_Sales_volume_10g_pre_test)/6;
pi_sales_volume_post_test = sum(pi_Sales_volume_5g_post_test, 2 * pi_Sales_volume_10g_post_test)/&post_test_nmth_available.;  /*will need to be divided by month when data is complete*/
pi_sales_value_pre_test = sum(pi_Sales_value_5g_pre_test, pi_Sales_value_10g_pre_test)/6;
pi_sales_value_post_test = sum(pi_Sales_value_5g_post_test, pi_Sales_value_10g_post_test)/&post_test_nmth_available.;  /*will need to be divided by month when data is complete*/

ji_sales_volume_pre_test = sum(ji_Sales_volume_20g_pre_test ,  2.5 * ji_Sales_volume_50g_pre_test)/6;
ji_sales_volume_post_test = sum(ji_Sales_volume_20g_post_test, 2.5 * ji_Sales_volume_50g_post_test)/&post_test_nmth_available.;  /*will need to be divided by month when data is complete*/
ji_sales_value_pre_test = sum(ji_Sales_value_20g_pre_test, ji_Sales_value_50g_pre_test)/6;
ji_sales_value_post_test = sum(ji_Sales_value_20g_post_test, ji_Sales_value_50g_post_test)/&post_test_nmth_available.;  /*will need to be divided by month when data is complete*/

array pi_sales  pi_sales_volume_pre_test  pi_sales_volume_post_test  pi_sales_value_pre_test  pi_sales_value_post_test;
array ji_sales   ji_sales_volume_pre_test  ji_sales_volume_post_test    ji_sales_value_pre_test  ji_sales_value_post_test;

do over pi_sales;
		if pi_sales = . then pi_sales=0;
		if ji_sales = . then ji_sales = 0;
end;

run;
/*for QC*/
/*proc export data= Bai_duo_bang_model_2*/
/*outfile = "D:\Project_Files2\CN PMR\百多邦\Jan01_2017\Bai_duo_bang_model_2.xlsx"*/
/*dbms=excel replace;*/
/*run;*/


/*check distribution of pre-test and post-test sales*/
/*proc univariate data= Bai_duo_bang_model;*/
/*histogram sales_volume_pre_test    sales_volume_post_test  sales_value_pre_test  sales_value_post_test;*/
/*INSET N = 'Number of Pharmacy Stores' MEDIAN (8.2) MEAN (8.2) STD='Standard Deviation' (8.3)/ POSITION = ne;*/
/*run;*/



/*compare sales volume -- 5g+10g*/
proc glm data= Bai_duo_bang_model_reset;

/*with only design variables included*/
/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM ;*/
/*model sales_volume_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM   sales_volume_pre_test/ss3 solution;*/

/*with design variables +  Chain_or_Single + Store_Area*/
/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;*/
/*model sales_volume_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined   sales_volume_pre_test  Store_Area/ss3 solution;*/
/*note: store_Area is not significant*/

/*with design variables +  Chain_or_Single*/
class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;
model sales_volume_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined   sales_volume_pre_test /ss3 solution;
/*note: effect of Option_POSM becomes very insignificant after Chai_or_Single_Combined brought in. the direction of estimated coefficient is opposite to the very first option.*/
lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM   Chai_or_Single_Combined/stderr pdiff;

run;quit;


/*compare sales value  -- 5g+10g*/
proc glm data= Bai_duo_bang_model_reset;

/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM ;*/
/*model sales_value_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  sales_value_pre_test/ss3 solution;*/
/*lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM /stderr pdiff;*/

/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;*/
/*model sales_value_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined  sales_value_pre_test  Store_Area/ss3 solution;*/
/*lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM   Chai_or_Single_Combined/stderr pdiff;*/
/*note: store_Area is not significant*/

class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;
model sales_value_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined  sales_value_pre_test/ss3 solution;
lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM   Chai_or_Single_Combined/stderr pdiff;


run;quit;

/*-- another test --*/

/*compare sales volume/value -- 5g only*/
proc glm data= Bai_duo_bang_model_reset;
/*proc glm data= Bai_duo_bang_model;*/
/*-- sales volume--*/
/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM ;*/
/*model avg_sales_volume_5g_post = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  avg_sales_volume_5g_pre /ss3 solution;*/

class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;
model avg_sales_volume_5g_post = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined  avg_sales_volume_5g_pre /ss3 solution;

/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;*/
/*model avg_sales_volume_5g_post = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined  avg_sales_volume_5g_pre Store_Area/ss3 solution;*/

/*-- sales value --*/
/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM ;*/
/*model avg_sales_value_5g_post = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  avg_sales_value_5g_pre /ss3 solution;*/

/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;*/
/*model avg_sales_value_5g_post = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined  avg_sales_value_5g_pre /ss3 solution;*/

/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;*/
/*model avg_sales_value_5g_post = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined  avg_sales_value_5g_pre  Store_Area/ss3 solution;*/

run;quit;

/*compare sales volume/value -- 10g only*/
proc glm data= Bai_duo_bang_model_reset;
/*proc glm data= Bai_duo_bang_model;*/
/*-- sales volume--*/
/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM ;*/
/*model avg_sales_volume_10g_post = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  avg_sales_volume_10g_pre /ss3 solution;*/

class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;
model avg_sales_volume_10g_post = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined  avg_sales_volume_10g_pre /ss3 solution;

/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;*/
/*model avg_sales_volume_10g_post = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined  avg_sales_volume_10g_pre Store_Area/ss3 solution;*/

/*-- sales value --*/
/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM ;*/
/*model avg_sales_value_10g_post = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  avg_sales_value_10g_pre /ss3 solution;*/

/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;*/
/*model avg_sales_value_10g_post = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined  avg_sales_value_10g_pre /ss3 solution;*/

/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;*/
/*model avg_sales_value_10g_post = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined  avg_sales_value_10g_pre Store_Area/ss3 solution;*/


run;quit;

/*---------------------------------------------------------------------------*/
/*---------------------------------------------------------------------------*/

/*some additional analysis.*/
proc glm data = Bai_duo_bang_model_2;
class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;
model sales_volume_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined   
sales_volume_pre_test   pi_sales_volume_post_test  ji_sales_volume_post_test/ss3 solution;
/*lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM   Chai_or_Single_Combined/stderr pdiff;*/
run;

proc glm data = Bai_duo_bang_model_2;
class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined;
model sales_value_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined   
sales_value_pre_test   pi_sales_value_post_test  ji_sales_value_post_test/ss3 solution;
/*lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM   Chai_or_Single_Combined/stderr pdiff;*/
run;

proc corr data= Bai_duo_bang_model_2;
var sales_volume_post_test;
with sales_volume_pre_test   pi_sales_volume_post_test  ji_sales_volume_post_test;
run;
proc corr data= Bai_duo_bang_model_2;
var sales_value_post_test;
with sales_value_pre_test   pi_sales_value_post_test  ji_sales_value_post_test;
run;


proc import out = promo_activity
datafile = "D:\Project_Files2\CN PMR\百多邦\Jan01_2017\百多邦项目140家门店销售数据_2016年4月-11月_Read_IN.xlsx"
dbms=excel replace;
range= "促销数据$Y4:AG144";
getnames=yes;
run;
data promo_activity(rename=(New_Pharma_Store_ID=Pharma_Store_ID ));
set promo_activity;
New_Pharma_Store_ID = strip(put(Pharma_Store_ID, 3.));
drop Pharma_Store_ID;
run;
proc sort data = promo_activity; by Pharma_Store_ID;run;
data Bai_duo_bang_model_3;
merge Bai_duo_bang_model_2(in=a) promo_activity(in=b);
by Pharma_Store_ID;
if a and b;

if Baiduobang_Oct_Promo ="有促销" or Baiduobang_Nov_Promo = "有促销" then Baiduobang_Promo = "有促销";
else Baiduobang_Promo= "无促销";

if Competitor_Oct_Promo = "有促销"  or Competitor_Oct_Promo = "有促销" then Competitor_Promo = "有促销";
else Competitor_Promo = "无促销";

if Store_Oct_Promo = "有促销"  or  Store_Nov_Promo = "有促销" then Store_Promo = "有促销";
else Store_Promo = "无促销";

run;

proc freq data= Bai_duo_bang_model_3;
tables Baiduobang_Oct_Promo  	Competitor_Oct_Promo	
Baiduobang_Nov_Promo	 Competitor_Nov_Promo	
Store_Oct_Promo	 Store_Nov_Promo	
Overall_Oct_Promo	Overall_Nov_Promo  Store_Oct_Promo*Overall_Oct_Promo  Store_Nov_Promo*Overall_Nov_Promo /missing;run;
proc freq data= Bai_duo_bang_model_3; tables Baiduobang_Promo  Competitor_Promo  Baiduobang_Promo*Competitor_Promo  Store_Oct_Promo/missing;run;

/*baseline model*/
proc glm data =  Bai_duo_bang_model_3;
class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined  ;
model sales_volume_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined   sales_volume_pre_test/ss3 solution;
run;quit;

/*test some more models --*/
proc glm data =  Bai_duo_bang_model_3;
class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined  
/*Baiduobang_Promo*/
Store_Promo
/*Competitor_Promo*/

;
model sales_volume_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM   
Chai_or_Single_Combined   
/*Baiduobang_Promo*/
Store_Promo
/*Competitor_Promo*/
/*Chai_or_Single_Combined(Baiduobang_Promo)*/
/*Chai_or_Single_Combined * Baiduobang_Promo */
/*Chai_or_Single_Combined|Baiduobang_Promo*/
/*Option_Position(Baiduobang_Promo)*/  
/*Option_Product_Range(Baiduobang_Promo)*/
/*Option_Shining(Baiduobang_Promo)*/
/*Option_POSM(Baiduobang_Promo)*/
/*Option_POSM * Baiduobang_Promo*/

/*Baiduobang_Promo * Chai_or_Single_Combined*/
Store_Promo * Chai_or_Single_Combined
/*Competitor_Promo * Chai_or_Single_Combined*/

/*Option_Product_Range * Option_Shining*/
/*Option_Product_Range * Option_POSM*/
sales_volume_pre_test    
sales_volume_pre_test * Chai_or_Single_Combined
/*sales_volume_pre_test * Option_Product_Range*/
/*sales_volume_pre_test * Option_Position*/
/*sales_volume_pre_test *  Option_Shining*/
/*sales_volume_pre_test *  Option_POSM*/
/*sales_volume_pre_test * Chai_or_Single_Combined*/
/*sales_volume_pre_test * Baiduobang_Promo*/
/*sales_volume_pre_test * Store_Promo*/
/*sales_volume_pre_test * Competitor_Promo*/
/*pi_sales_volume_post_test   */
/*ji_sales_volume_post_test*/

/*ji_sales_volume_post_test*Option_Position*/
/*ji_sales_volume_post_test*Option_Product_Range*/
/*ji_sales_volume_post_test*Option_Shining*/
/*ji_sales_volume_post_test*Option_POSM*/
/*ji_sales_volume_post_test * sales_volume_pre_test*/

/*Chai_or_Single_Combined * Option_Position  */
/*Chai_or_Single_Combined * Option_Product_Range   */
/*Chai_or_Single_Combined * Option_Shining   */
Chai_or_Single_Combined * Option_POSM   

/ss3 solution;
lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined
/*Baiduobang_Promo*/
Store_Promo

/*Competitor_Promo*/
;
run;quit;

proc glm data =  Bai_duo_bang_model_3;
class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined  
/*Baiduobang_Promo*/
Store_Promo
/*Competitor_Promo*/

;
model sales_value_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM   
Chai_or_Single_Combined   
/*Baiduobang_Promo*/
Store_Promo
/*Competitor_Promo*/

/*Baiduobang_Promo * Chai_or_Single_Combined*/
Store_Promo * Chai_or_Single_Combined
/*Competitor_Promo * Chai_or_Single_Combined*/

/*Option_Product_Range * Option_Shining*/
/*Option_Product_Range * Option_POSM*/
sales_value_pre_test    
sales_value_pre_test * Chai_or_Single_Combined

Chai_or_Single_Combined * Option_POSM   
/*Chai_or_Single_Combined * Option_Product_Range*/

/ss3 solution;
lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined
/*Baiduobang_Promo*/
Store_Promo/stderr pdiff;

/*Competitor_Promo*/
;
run;quit;

/*-------------------------*/
/*-------------------------*/


/*one good model for memo*/
proc glm data =  Bai_duo_bang_model_3;
class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined  ;
model sales_volume_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined
sales_volume_pre_test    
sales_volume_pre_test * Option_Position
sales_volume_pre_test *  Option_POSM/ss3 solution;
lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined
/*sales_volume_pre_test*/
/*sales_volume_pre_test * Option_Position*/
;
run;quit;


proc glm data =  Bai_duo_bang_model_3;
class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined  ;
model sales_value_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined
sales_value_pre_test    
sales_value_pre_test * Option_Position
sales_value_pre_test *  Option_POSM/ss3 solution;
lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined

;
run;quit;

/*re-run the first model without any interaction term*/
proc glm data =  Bai_duo_bang_model_3;
class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined  ;
model sales_volume_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined
sales_volume_pre_test /ss3 solution;
lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined
;
run;quit;

proc glm data =  Bai_duo_bang_model_3;
class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined  ;
model sales_value_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined   
/*Chai_or_Single_Combined*Option_Position*/
sales_value_pre_test /ss3 solution;
lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined
/* Chai_or_Single_Combined*Option_Position*/
;
run;quit;


/*proc glm data =  Bai_duo_bang_model_3;*/
/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  */
/*Chai_or_Single_Combined  ;*/
/*model sales_volume_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined*/
/*sales_volume_pre_test */
/*Chai_or_Single_Combined * Option_Position*/
/*Chai_or_Single_Combined * Option_Product_Range*/
/*Chai_or_Single_Combined * Option_Shining*/
/*Chai_or_Single_Combined * Option_POSM*/
/*/ss3 solution;*/
/*lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM  */
/*Chai_or_Single_Combined*/
/**/
/*;*/
/*run;quit;*/

/*--- descriptive analysis  ---*/

/*proc summary data = Bai_duo_bang_model_3 nway;*/
/*by Option_Position  Option_Product_Range   Option_Shining   Option_POSM  */
/*Chai_or_Single_Combined*/
/*by Option_Position;*/

%macro summary_by_var(input_data, by_var);

proc summary data = &input_data  nway;
class &by_var;
var   sales_value_pre_test    sales_value_post_test ;
output out = mean_by_&by_var(drop = _:) mean = /autoname;
run;

proc export data = mean_by_&by_var
outfile = "D:\Project_Files2\CN PMR\百多邦\Jan01_2017\pre_model_summary.xlsx"
dbms = excel replace;
sheet= "&by_var";
run;

%mend summary_by_var;

%let by_var_list =  Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined  ;

/* %foreach(v,&mth_id_list, %nrstr(&prefix_01A._&v))*/
%foreach(v, &by_var_list, %nrstr(%summary_by_var(Bai_duo_bang_model_3,  &v)));

 %summary_by_var(Bai_duo_bang_model_3, Store_Promo );
