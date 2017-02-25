
/*

This version is based on "Refresh Model on Feb06 2017 - 副本 v2.sas".

The main update is to calculate score of each combination of factors;
also do ancova analysis for the product category.
-- Feb 24, 2017

*/


%let path = %str(D:\Project_Files2\CN PMR\百多邦\Feb06_2017);

libname datadir2 "D:\Project_Files2\CN PMR\百多邦\Feb06_2017\update";

/*proc import out=datadir2.basic_info*/
/*datafile = "&path\副本百多邦项目140家门店基本情况_170207.xlsx"*/
/*dbms=excel replace;*/
/*range= "基本情况$A3:J143";*/
/*getnames = yes;*/
/*run;*/
/*proc import out=datadir2.bai_duo_bang*/
/*datafile = "&path\百多邦项目140家门店销售数据_2016年4月-12月_Read_IN.xlsx"*/
/*dbms=excel replace;*/
/*range= "百多邦销售数据$A3:AM144";*/
/*getnames = yes;*/
/*run;*/
/*proc import out=datadir2.pi_de_bang*/
/*datafile = "&path\百多邦项目140家门店销售数据_2016年4月-12月_Read_IN.xlsx"*/
/*dbms=excel replace;*/
/*range= "匹得邦销售数据$A3:AM144";*/
/*getnames = yes;*/
/*run;*/
/*proc import out=datadir2.jing_wan_hong*/
/*datafile = "&path\百多邦项目140家门店销售数据_2016年4月-12月_Read_IN.xlsx"*/
/*dbms=excel replace;*/
/*range= "京万红销售数据$A3:AM144";*/
/*getnames = yes;*/
/*run;*/
/**/
/*data datadir2.bai_duo_bang;*/
/*set datadir2.bai_duo_bang;*/
/*if (City NE "占位符") ;*/
/**/
/*run;*/
/*data datadir2.pi_de_bang;*/
/*set datadir2.pi_de_bang;*/
/*if (City NE "占位符") ;*/
/**/
/*run;*/
/*data datadir2.jing_wan_hong;*/
/*set datadir2.jing_wan_hong;*/
/*if (City NE "占位符") ;*/
/**/
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

%let post_test_nmth_available = 3 ;  /*the final version*/

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


proc import out = promo_activity
datafile = "&path.\百多邦项目140家门店销售数据_2016年4月-12月_Read_IN.xlsx"
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

/*do some regression and detect some outstanding store with weired sales value in Dec*/
data Bai_duo_bang_model_3reg;
set Bai_duo_bang_model_3;
if Pharma_Store_ID not in ("132","11","87","88");

sales_value_mth_10_11 = sum(Sales_value_5g_10, Sales_value_5g_11, Sales_value_10g_10, Sales_value_10g_11);
sales_value_mth_12 = sum(Sales_value_5g_12, Sales_value_10g_12);

if sales_value_mth_10_11 = .  then sales_value_mth_10_11=0;
if sales_value_mth_12=. then sales_value_mth_12=0;

run;

proc reg data=Bai_duo_bang_model_3reg;
model sales_value_mth_12 = sales_value_mth_10_11;
output out=reg_out RSTUDENT=residual_student;
run;quit;

%let ALPHA = 1.96;

data pharm_to_delete;
set reg_out(keep = Pharma_Store_ID residual_student);
if abs(residual_student) > &ALPHA.;
run;
/*--end--*/
proc print data= pharm_to_delete;
var Pharma_Store_ID;
run;

/*correction for store ID = 23*/
data Bai_duo_bang_model_3mod;
set Bai_duo_bang_model_3;

if Pharma_Store_ID not in ("132","11","87","88");
/*if Pharma_Store_ID not in ("57","58","59","60","61","62","63","64","46","49","51","52","53","54","55");*/
if Pharma_Store_ID ="23" then do;
sales_value_post_test = sales_value_post_test *3/2;
end;

/*add on Feb 08*/
if pi_sales_value_post_test =. then pi_sales_value_post_test=0;
if ji_sales_value_post_test=. then ji_sales_value_post_test=0;
if pi_sales_value_pre_test =. then pi_sales_value_pre_test=0;
if ji_sales_value_pre_test=. then ji_sales_value_pre_test=0; 

piji_sales_val_post_test  = sum(pi_sales_value_post_test, ji_sales_value_post_test) ;
piji_sales_val_all = sum(6*pi_sales_value_pre_test, 6*ji_sales_value_pre_test, &post_test_nmth_available.*pi_sales_value_post_test, &post_test_nmth_available.*ji_sales_value_post_test)/9;


run;

/*to roll back to first 2 months pos-test*/
data Bai_duo_bang_model_3QC;
set Bai_duo_bang_model_3;
if Pharma_Store_ID not in ("132","11","87","88");
/*if Pharma_Store_ID not in ("57","58","59","60","61","62","63","64","46","49","51","52","53","54","55");*/

if Sales_value_5g_12 = . then Sales_value_5g_12= 0;
if Sales_value_10g_12 =. then Sales_value_10g_12 = 0;

/*roll back to first 2 months*/
sales_value_post_test_2mth = (sales_value_post_test * 3 -  sum(Sales_value_5g_12 , Sales_value_10g_12) )/2;

sales_value_pre_test_3mth = sum(Sales_value_5g_07, Sales_value_5g_08, Sales_value_5g_09, Sales_value_10g_07, Sales_value_10g_08, Sales_value_10g_09)/3;

/*add on Feb 08*/
if pi_sales_value_post_test =. then pi_sales_value_post_test=0;
if ji_sales_value_post_test=. then ji_sales_value_post_test=0;
if pi_sales_value_pre_test =. then pi_sales_value_pre_test=0;
if ji_sales_value_pre_test=. then ji_sales_value_pre_test=0; 

piji_sales_val_post_test  = sum(pi_sales_value_post_test, ji_sales_value_post_test) ;
piji_sales_val_all = sum(6*pi_sales_value_pre_test, 6*ji_sales_value_pre_test, &post_test_nmth_available.*pi_sales_value_post_test, &post_test_nmth_available.*ji_sales_value_post_test)/9;

run;
proc sql;
select nmiss(sales_value_pre_test_3mth) from Bai_duo_bang_model_3QC;quit;

%macro delete_outlier(input_data);

proc sql;
create table &input_data._del as
select * from &input_data
where  Pharma_Store_ID not in (select Pharma_Store_ID from pharm_to_delete);
quit;

%mend;

%delete_outlier(Bai_duo_bang_model_3mod);
%delete_outlier(Bai_duo_bang_model_3QC);

%let INPUT_DATA_1 = Bai_duo_bang_model_3mod_del;
%let INPUT_DATA_2 = Bai_duo_bang_model_3QC_del;

/*base model without interaction term*/
/*proc glm data =  &INPUT_DATA_1;*/
/*class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  */
/*Chai_or_Single_Combined  ;*/
/*model sales_value_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM  Chai_or_Single_Combined   */
/**/
/*sales_value_pre_test /ss3 solution;*/
/*lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM  */
/*Chai_or_Single_Combined*/
/*;*/
/**/
/*run;quit;*/
/*proc freq data=&INPUT_DATA_1;tables Competitor_Promo  Store_Promo  Baiduobang_Promo/missing;run;*/


proc glm data =  &INPUT_DATA_1;
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

sales_value_pre_test    
sales_value_pre_test * Chai_or_Single_Combined

Chai_or_Single_Combined * Option_POSM   
/*Chai_or_Single_Combined * Option_Product_Range*/
/*Chai_or_Single_Combined * Option_Shining*/
Chai_or_Single_Combined * Option_Position

/*Store_Promo * Option_Product_Range*/
/*Store_Promo * Option_POSM   */
/*Store_Promo * Option_Position   */
/*Store_Promo * Option_Shining   */

/*Competitor_Promo * Option_Product_Range*/
/*Competitor_Promo * Option_POSM */
/*Competitor_Promo * Option_Position   */
/*Competitor_Promo * Option_Shining   */

/*Baiduobang_Promo * Option_Product_Range*/
/*Baiduobang_Promo * Option_POSM */
/*Baiduobang_Promo * Option_Position   */
/*Baiduobang_Promo * Option_Shining   */

/ss3 solution;
lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined
/*Baiduobang_Promo*/
Store_Promo
/*Competitor_Promo*/

/stderr pdiff /*adjust = NELSON*/ ;

run;quit;

proc freq data = &INPUT_DATA_1;
tables 
Option_Position  
Option_Product_Range   
Option_Shining   
Option_POSM  
Chai_or_Single_Combined  
Store_Promo/missing;
run;

proc means data = &INPUT_DATA_1 mean;
var sales_value_pre_test;
run;

/*
average of sales value_pre_test : 
sales_value_pre_test  = 349.9155906

will use 350 as the average
*/

/*added on Feb 24, 2017*/
data full_design;

Intercept = 1;

do Chai_or_Single_Combined =0 to 1;
		do  Store_Promo = 0 to 1;
			do Option_Position =0 to 1;
					do Option_Product_Range = 0 to 1;
							do Option_Shining = 0 to 1;
									do Option_POSM = 0 to 1;
											/*	interaction term	*/

											Chain_Sing_INT_Store_Promo =  Chai_or_Single_Combined * Store_Promo;
											sales_value_pre = 350;
											sales_value_INT_Chain_Sing = 350 *  Chai_or_Single_Combined;
											Option_POSM_INT_Chai_Sing  =  Option_POSM  *  Chai_or_Single_Combined;
											Option_Position_INT_Chai_Sing = Option_Position * Chai_or_Single_Combined;

											output;

									end;
							end;
					end;
				end;
			end;
	end;
		

run;

proc export data= full_design
outfile = "D:\Project_Files2\CN PMR\百多邦\Feb06_2017\百多邦初步分析(基于10-12月数据) 更新.xlsx"
dbms=excel replace;
sheet= "all_combos";
run;

/*data for consulting team*/
data temp_out_for_CS;

retain Pharma_Store_ID 

sales_value_post_test 

sales_value_pre_test

Option_Position  Option_Product_Range   Option_Shining   Option_POSM  

Chai_or_Single_Combined  
Store_Promo
;

set &INPUT_DATA_1.;

keep Pharma_Store_ID 

sales_value_post_test 

sales_value_pre_test

Option_Position  Option_Product_Range   Option_Shining   Option_POSM  

Chai_or_Single_Combined  
Store_Promo;

run;

proc export data = temp_out_for_CS
outfile = "D:\Project_Files2\CN PMR\百多邦\Feb06_2017\Model_Data_For_CS_Team_Feb24.xlsx"
dbms=excel replace;
sheet= "Model_Data";
run;

/*model done in Jan*/
proc glm data =  &INPUT_DATA_2;
class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined  
Store_Promo
;
model sales_value_post_test_2mth = Option_Position  Option_Product_Range   Option_Shining   Option_POSM   
Chai_or_Single_Combined   
Store_Promo
Store_Promo * Chai_or_Single_Combined
sales_value_pre_test
sales_value_pre_test * Chai_or_Single_Combined
Chai_or_Single_Combined * Option_POSM   
/ss3 solution;
lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined
Store_Promo/stderr pdiff;

;
run;quit;


proc sql;
select nmiss(pi_sales_value_post_test), nmiss(ji_sales_value_post_test), nmiss(piji_sales_val_post_test), nmiss(piji_sales_val_all) from &INPUT_DATA_1;quit;

/*include competitor sales*/
proc glm data =  &INPUT_DATA_1;
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

sales_value_pre_test    
sales_value_pre_test * Chai_or_Single_Combined

Chai_or_Single_Combined * Option_POSM   
/*Chai_or_Single_Combined * Option_Product_Range*/
/*Chai_or_Single_Combined * Option_Shining*/
Chai_or_Single_Combined * Option_Position

/*Store_Promo * Option_Product_Range*/
/*Store_Promo * Option_POSM   */
/*Store_Promo * Option_Position   */
/*Store_Promo * Option_Shining   */

/*Competitor_Promo * Option_Product_Range*/
/*Competitor_Promo * Option_POSM */
/*Competitor_Promo * Option_Position   */
/*Competitor_Promo * Option_Shining   */

/*Baiduobang_Promo * Option_Product_Range*/
/*Baiduobang_Promo * Option_POSM */
/*Baiduobang_Promo * Option_Position   */
/*Baiduobang_Promo * Option_Shining   */

piji_sales_val_post_test
/*piji_sales_val_post_test * Chai_or_Single_Combined*/
/*pi_sales_value_post_test*/
/*ji_sales_value_post_test*/
/*piji_sales_val_all*/
/*piji_sales_val_all* Chai_or_Single_Combined*/

/ss3 solution;
lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined
/*Baiduobang_Promo*/
Store_Promo
/*Competitor_Promo*/

/stderr pdiff /*adjust = NELSON*/ ;

run;quit;


/*-------------------------*/
/*try new model with 3 months pre-test period*/
/*-------------------------*/
proc glm data =  &INPUT_DATA_2;
class Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined  
Store_Promo
;
model sales_value_post_test = Option_Position  Option_Product_Range   Option_Shining   Option_POSM   
Chai_or_Single_Combined   
Store_Promo
Store_Promo * Chai_or_Single_Combined
sales_value_pre_test_3mth
sales_value_pre_test_3mth * Chai_or_Single_Combined
/*Chai_or_Single_Combined * Option_POSM   */
/ss3 solution;
lsmeans Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined
Store_Promo/stderr pdiff;

run;quit;

/*summary*/
%macro summary_by_var(input_data, by_var);

proc summary data = &input_data  nway;
class &by_var;
var   sales_value_pre_test    sales_value_post_test ;
output out = mean_by_&by_var(drop = _:) mean = /autoname;
run;

proc export data = mean_by_&by_var
outfile = "D:\Project_Files2\CN PMR\百多邦\Feb06_2017\pre_model_summary_dele.xlsx"
dbms = excel replace;
sheet= "&by_var";
run;

%mend summary_by_var;

%let by_var_list =  Option_Position  Option_Product_Range   Option_Shining   Option_POSM  
Chai_or_Single_Combined  Store_Promo;

/* %foreach(v,&mth_id_list, %nrstr(&prefix_01A._&v))*/
%foreach(v, &by_var_list, %nrstr(%summary_by_var(&INPUT_DATA_1,  &v)));



/*further test*/

/*AT MEANS*/
