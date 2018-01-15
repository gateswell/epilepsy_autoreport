use strict;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseXLSX;
use Excel::Writer::XLSX;
use Tk;
use Tk::DateEntry;
use Time::Local;
use Encode;
use Win32;
use utf8;

my $mw = MainWindow -> new;
$mw->geometry("835x325");
$mw->title('抗癫痫药物过敏HLA基因筛查与咨询报告生成系统V1.3.180104');

my ($thisday,$thismon,$year) = (localtime)[3..5];
$thisday = sprintf("%02d", $thisday);
$thismon  = sprintf("%02d", $thismon + 1);
$year = $year + 1900;
my $today = $year.'-'.$thismon.'-'.$thisday;
my $today_alias = $year.$thismon.$thisday;

my $font1 = $mw ->fontCreate(-family => '楷体',-size=>12);
my $font2 = $mw ->fontCreate(-family => 'Times New Roman',-size=>9);

my $frame1 = $mw -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
my $sampleInfo = $frame1-> Label(-text=>"请选择样本信息(.xlsx):",-width=>25,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
my $sampleInfoEntry = $frame1->Entry(-text=>'例如：C:\Users\Administrator\Desktop\样本信息.xlsx',-width=>100,-font=>$font2)->pack(-side=>'left');
my $sampleInfoButton =$frame1->Button(-text=>'...',-width=>2,-command=>\&SampleInput)->pack(-side=>'left');

my $frame2 = $mw -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
my $missionInfo = $frame2-> Label(-text=>"请选择生产任务单(.xlsx):",-width=>25,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
my $missionInfoEntry = $frame2->Entry(-text=>'例如：C:\Users\Administrator\Desktop\生产任务单.xlsx',-width=>100,-font=>$font2)->pack(-side=>'left');
my $missionInfoButton =$frame2->Button(-text=>'...',-width=>2,-command=>\&missionInput)->pack(-side=>'left');

my $frame3 = $mw -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
my $TypingInfo = $frame3-> Label(-text=>"请选择导出结果(.xlsx):",-width=>25,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
my $TypingInfoEntry = $frame3->Entry(-text=>'例如：C:\Users\Administrator\Desktop\导出结果.xlsx',-width=>100,-font=>$font2)->pack(-side=>'left');
my $TypingInfoButton =$frame3->Button(-text=>'...',-width=>2,-command=>\&TypingInput)->pack(-side=>'left');

my $frame10 = $mw -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
my $destdir = $frame10->Label(-text=>"结果另存为:",-width=>15,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
#my $destdirEntry =$frame10->Entry(-text=>'C:\Users\Administrator\Desktop\new.xlsx',-width=>50)->pack(-side=>'left');
my $destdirEntry =$frame10->Entry(-text=>"C\:\\Users\\Administrator\\Desktop\\$today_alias\.xlsx",-width=>50,-font=>$font2)->pack(-side=>'left');
my $destdirButton =$frame10->Button(-text=>'...',-width=>2,-command=>\&destdirInput)->pack(-side=>'left');

my $frame4 = $mw -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
my $detector =$frame4 -> Label(-text=> '请键入检测者：',-width=>15,-anchor =>'w',-font=>$font1)->pack(-side=>'left');
my $detectorEntry = $frame4->Entry(-textvariable=> 'N/A',-width=>50)->pack(-side => 'left');

my $frame5 = $mw -> Frame()->pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
my $checkerInfo = $frame5 -> Label(-text => '请键入审核者：',-width=>15,-anchor => 'w',-font=>$font1)->pack(-side => 'left');
my $checkerInfoEntry = $frame5->Entry(-text => 'N/A',-width=>50)->pack(-side=> 'left');

my $frame6 = $mw -> Frame()->pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
my $signerInfo = $frame6->Label(-text=> '请键入签发者：',-width=>15,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
my $signerInfoEntry =$frame6->Entry(-text=> '卓孝福',-width=>50)->pack(-side=>'left');

#####插入日期的选择

my %idx_for_mon =( '01'=>1, '02'=>2, '03'=>3, '04'=> 4, '05'=> 5, '06'=> 6,'07'=>7, '08'=>8, '09'=>9, '10'=>10, '11'=>11, '12'=>12 );
my $defaultime1 = "$today";
my $defaultime2 = "$today";

my $frame7 = $mw -> Frame()->pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
my $recheckdate= $frame7->Label(-text=>'复核日期：',-width=>15, -anchor=>'w',-font=>$font1)->pack(-side=>'left');
my $recheckdateEntry =$frame7->DateEntry( -textvariable=>\$defaultime1, -width=>15, -parsecmd=>\&parse, -formatcmd=>\&format )->pack(-side=>'left');

my $frame8 = $mw -> Frame()->pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
my $signdate= $frame8->Label(-text=>'签发日期：',-width=>15, -anchor=>'w',-font=>$font1)->pack(-side=>'left');
my $signdateEntry = $frame8 ->DateEntry(-textvariable=>\$defaultime2, -width=>15, -parsecmd=>\&parse, -formatcmd=>\&format )->pack(-side=>'left');

my $frame9 = $mw -> Frame()->pack(-side=>'bottom',-anchor=>'se',-fill=>'x',-ipadx=>5,-ipady=>5);
my $nextstep= $frame9 ->Button(-text=>'下一步',-height=>4,-width=>5,-command=>\&get_all)->pack(-side=>'right');

my $file_initial;

$mw->resizable( 0, 0 );
sub parse {									#读取日期
  my ( $yr, $mon, $day ) = split '-', $_[0];
  return ( $yr, $idx_for_mon{$mon}, $day );
}
sub format {
  my ( $yr, $mon, $day ) = @_;
  return sprintf( "%4d-%02d-%02d", $yr, $mon, $day );
}

sub SampleInput{
	$file_initial = $mw->getOpenFile(-initialdir=>'C:\Users\Administrator\Desktop');
	$sampleInfoEntry ->delete('0.0','end');	#delete(index1, index2)	Deletes items from index1 to index2.
	$sampleInfoEntry->insert("end",$file_initial);	#insert(index, string)	Inserts the text of string at the specified index. This string then becomes available as one of the choices.
	$file_initial =encode("GB2312",$file_initial);
	my $file_initial1 =decode("GB2312",$file_initial);
	#print $fh $file_initial;
}

sub missionInput{
	$file_initial = $mw->getOpenFile(-initialdir=>'C:\Users\Administrator\Desktop');
	$missionInfoEntry ->delete('0.0','end');	#delete(index1, index2)	Deletes items from index1 to index2.
	$missionInfoEntry->insert("end",$file_initial);	#insert(index, string)	Inserts the text of string at the specified index. This string then becomes available as one of the choices.
	$file_initial =encode("GB2312",$file_initial);
	my $file_initial1 =decode("GB2312",$file_initial);
	#print $fh $file_initial;
}

sub TypingInput{
	$file_initial = $mw->getOpenFile(-initialdir=>'C:\Users\Administrator\Desktop');
	$TypingInfoEntry ->delete('0.0','end');	#delete(index1, index2)	Deletes items from index1 to index2.
	$TypingInfoEntry->insert("end",$file_initial);	#insert(index, string)	Inserts the text of string at the specified index. This string then becomes available as one of the choices.
	$file_initial =encode("GB2312",$file_initial);
	my $file_initial1 =decode("GB2312",$file_initial);
	#print $fh $file_initial;
}

sub destdirInput{
	#$file_initial = $mw->getOpenFile(-initialdir=>'C:\Users\Administrator\Desktop\new.xlsx');
	$file_initial = $mw->getOpenFile(-initialdir=>"C\:\\Users\\Administrator\\Desktop\\$today_alias\.xlsx");
	$destdirEntry ->delete('0.0','end');	#delete(index1, index2)	Deletes items from index1 to index2.
	$destdirEntry->insert("end",$file_initial);	#insert(index, string)	Inserts the text of string at the specified index. This string then becomes available as one of the choices.
	$file_initial =encode("GB2312",$file_initial);
	my $file_initial1 =decode("GB2312",$file_initial);
}

my($Sampledir,$missiondir,$Typingdir,$Destdir,$checkername,$detectorname,$signername,$rechedate,$signdate); 
sub get_all{
	$Sampledir =$sampleInfoEntry->get();
	$missiondir = $missionInfoEntry->get();
	$Typingdir = $TypingInfoEntry->get();
	$Destdir = $destdirEntry->get();
	$checkername= $checkerInfoEntry->get();
	$detectorname= $detectorEntry->get();
	$signername = $signerInfoEntry->get();
	$rechedate= $recheckdateEntry->get();
	$signdate = $signdateEntry->get();
	$Sampledir=encode("GB2312",$Sampledir);
	$missiondir=encode("GB2312",$missiondir);
	$Typingdir=encode("GB2312",$Typingdir);
	$Destdir=encode("GB2312",$Destdir);
	$checkername=encode("GB2312",$checkername);
	$detectorname=encode("GB2312",$detectorname);
	$signername=encode("GB2312",$signername);
	#print $fh $Sampledir,"\n",$missiondir,"\n",$Typingdir,"\n",$checkername,"\n",$detectorname,"\n",$signername,"\n",$rechedate,"\n",$signdate,"\n";
	$Sampledir=~ s/\//\\/g;
	$missiondir=~ s/\//\\/g;
	$Typingdir=~ s/\//\\/g;
	$Destdir=~ s/\//\\/g;
	#print $fh $Sampledir,"\n",$missiondir,"\n",$Typingdir,"\n",$checkername,"\n",$detectorname,"\n",$signername,"\n",$rechedate,"\n",$signdate,"\n";

####展示输入的结果，做进一步确认
	$Sampledir=decode("GB2312",$Sampledir);
	$missiondir=decode("GB2312",$missiondir);
	$Typingdir=decode("GB2312",$Typingdir);
	$Destdir=decode("GB2312",$Destdir);
	$checkername=decode("GB2312",$checkername);
	$detectorname=decode("GB2312",$detectorname);
	$signername=decode("GB2312",$signername);
	#print $fh $Sampledir,"\n",$missiondir,"\n",$Typingdir,"\n",$checkername,"\n",$detectorname,"\n",$signername,"\n",$rechedate,"\n",$signdate,"\n";
 
	my $displayAll = $mw->Toplevel;
	$displayAll->geometry("835x325");
	$displayAll->title('抗癫痫药物过敏HLA基因筛查与咨询报告生成系统V1.1');
	my $frame_1 = $displayAll -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
	my $display1 = $frame_1-> Label(-text=>"样本信息:$Sampledir",-width=>500,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
	$mw->messageBox(-message=>'请填入样本信息文件，回到上一步',-type=>'ok') unless $Sampledir;
	my $erro_msge = '例如';
	$erro_msge = encode("GB2312",$erro_msge);
	$erro_msge = decode("GB2312",$erro_msge);
	$mw->messageBox(-message=>'请填入样本信息文件，回到上一步',-type=>'ok') if $Sampledir =~ /$erro_msge/;
	
	my $frame_2 = $displayAll -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
	my $display2 = $frame_2-> Label(-text=>"生产任务单:$missiondir",-width=>500,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
	$mw->messageBox(-message=>'请填入生产任务单，回到上一步',-type=>'ok') unless $missiondir;
	$mw->messageBox(-message=>'请填入生产任务单，回到上一步',-type=>'ok') if $missiondir =~ /$erro_msge/;
	
	my $frame_3 = $displayAll -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
	my $display3 = $frame_3-> Label(-text=>"TypingResult:$Typingdir",-width=>500,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
	$mw->messageBox(-message=>'请填入下机导出结果文件，回到上一步',-type=>'ok') unless $Typingdir;
	$mw->messageBox(-message=>'请填入下机导出结果文件，回到上一步',-type=>'ok') if $Typingdir =~ /$erro_msge/;
	
	my $frame_4 = $displayAll -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
	$Destdir='C:\Users\Administrator\Desktop\new.xlsx' unless $Destdir;
	my $display4 = $frame_4-> Label(-text=>"结果文件存为:$Destdir",-width=>500,-anchor=>'w',-font=>$font1)->pack(-side=>'left');	
	
	my $frame_5 = $displayAll -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
	my $display5 = $frame_5-> Label(-text=>"检测者:$detectorname",-width=>500,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
	
	my $frame_6 = $displayAll -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
	my $display6 = $frame_6-> Label(-text=>"审核者:$checkername",-width=>500,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
	
	my $frame_7 = $displayAll -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
	my $display7 = $frame_7-> Label(-text=>"签发者:$signername",-width=>500,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
	
	my $frame_8 = $displayAll -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
	my $display8 = $frame_8-> Label(-text=>"复核日期:$rechedate",-width=>500,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
	
	my $frame_9 = $displayAll -> Frame()-> pack(-side=>'top',-anchor=>'w',-fill=>'x',-ipady=>5);
	my $display9 = $frame_9-> Label(-text=>"签发日期:$signdate",-width=>500,-anchor=>'w',-font=>$font1)->pack(-side=>'left');
	
	my $frame_10 = $displayAll -> Frame()-> pack(-side=>'top',-anchor=>'se',-fill=>'x',-ipadx=>5,-ipady=>5);
	my $leftdisplaybutton = $frame_10 ->Button(-text=>'上一步',-height=>3,-width=>8,-command=>[$displayAll=>'destroy'])->pack(-side=>'left');
	my $rightdisplaybutton= $frame_10 ->Button(-text=>'确认并生成报告',-height=>3,-width=>12,-command=>\&report_get)->pack(-side=>'right');
}	
	
sub report_get{	
	my %situration = ();
	my %describe_situ =();
	my ($typeA,$typeB,$class,$describe,@TypeA,@TypeB);
	my %ID = ();	#键：医院全称 值：医院编号
	my %alias = ();	#键：简称 值：全称
	my %region = ();	#键：医院全称 值：所属省份
	my %report_sheet = ();	#键：报告单编号 值：样本编码
	my %report_patient = ();	#键：报告单编号 值：患者姓名
	my %lab_report =();	#键：实验室编码 值：报告单编号
	
	my $parser     = Spreadsheet::ParseXLSX->new();
	$Sampledir=encode("GB2312",$Sampledir);
	my $workbook_1 = $parser ->parse("$Sampledir");
	#my $workbook_1 = $parser ->parse('D:\CTAB\epilepsy\test\2017epinfo.xlsx');
	unless ($workbook_1){
		$mw->messageBox(-message=>'没有样本信息，请重新选择！',-type=>'ok');
			#next;
	}
	my ($province,$hospCode,$hospital,$hospAlias,@hospAliases);
	my ($operator,$lab_id,$report_id,$patient);
	my %gender_hash=();	#报告单编号为键，性别为值
	my %age_hash =();	#报告单编号为键，年龄为值
	my %hosptal_hash =();#报告单编号为键，送检医院为值
	my %department_hash =();	#报告单编号为键，送检科室为值
	my %doctor_hash =(); #报告单编号为键，送检医生为值
	my %collect_hash =();#报告单编号为键，采样日期为值
	my %receive_hash =();#报告单编号为键，收样日期为值
	my ($province,$sample_id,$receive_date,$collect_date,$report_id,$name,$gender,$age,$send_hospital,$send_department,$send_doctor);
	for my $worksheet($workbook_1->worksheet(0)){
		my ($row_min,$row_max) =$worksheet->row_range();
		my ($col_min,$col_max) =$worksheet->col_range();
		#print $col_min,"\t",$col_max,"\n";
		for my $row(($row_min+1) .. $row_max){
			my $cell = $worksheet->get_cell($row,$col_min);
			$province = encode ('GB2312', $cell->value());
			my $cell = $worksheet->get_cell($row,$col_min+2);
			$receive_date = encode ('GB2312', $cell->value());
			my $receive_year = substr($receive_date,0,4);
			my $receive_month= substr($receive_date,4,2);
			my $receive_day  = substr($receive_date,6,2);
			my $receive_date_1 = $receive_year.'-'.$receive_month.'-'.$receive_day;
			my $cell = $worksheet ->get_cell($row,$col_min+6);
			$collect_date = $cell? encode ('GB2312', $cell->value()) :'' ; #采样日期
			my $collect_year = substr($collect_date,0,4);
			my $collect_month= substr($collect_date,4,2);
			my $collect_day  = substr($collect_date,6,2);
			my $collect_date_1 = $collect_year.'-'.$collect_month.'-'.$collect_day;
			my $cell = $worksheet->get_cell($row,$col_min+8);
			$sample_id= $cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet ->get_cell($row,$col_min+9);
			$report_id = $cell? encode ('GB2312', $cell->value()) :'';	#报告单编号 CSTB2017EM00022
			my $cell = $worksheet ->get_cell($row,$col_min+10);
			$name = encode ('GB2312', $cell->value());		#患者姓名
			my $cell = $worksheet ->get_cell($row,$col_min+12);
			$gender = $cell? encode ('GB2312', $cell->value()) :'';		#患者性别
			my $cell = $worksheet ->get_cell($row,$col_min+13);
			$age = $cell? encode ('GB2312', $cell->value()) :"";	#有些患者为婴幼儿
			#$age =~ /\d+|(\d+)M/;
			my $cell = $worksheet ->get_cell($row,$col_min+18);
			$send_hospital = $cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet ->get_cell($row,$col_min+19);
			$send_department = $cell? encode ('GB2312', $cell->value()) :'';	#送检科室
			if(exists $alias{$send_hospital}){
				$send_hospital = $alias{$send_hospital};
			}else{
				$send_hospital;
			}
			my $cell = $worksheet ->get_cell($row,$col_min+20);
			$send_doctor = $cell? encode ('GB2312', $cell->value()) :'';
			#print OUT $province,"\t",$sample_id,"\t",$report_id,"\t",$name,"\t",$gender,"\t",$age,"\t",$send_hospital,"\n";
			if($report_id){
				$report_sheet{$report_id} = $sample_id;
				$report_patient{$report_id} = $name;
				$gender_hash{$report_id} = $gender;
				$age_hash{$report_id} = $age;
				$hosptal_hash{$report_id} = $send_hospital;
				$department_hash{$report_id} = $send_department;
				$doctor_hash{$report_id} = $send_doctor;
				$collect_hash{$report_id} = $collect_date_1;
				$receive_hash{$report_id} = $receive_date_1;
				#print OUT $report_sheet{$report_id},"\t",$report_patient{$report_id},"\t",$age_hash{$report_id},"\t",$gender_hash{$report_id},"\t",$hosptal_hash{$report_id},"\t",$doctor_hash{$report_id},"\n";
			}
			#else{ print "缺少报告单编号，请修改！" ;}
			else{$mw->messageBox(-message=>'生产信息中缺少报告单编号，请修改！',-type=>'ok');}
			#print "没有填写采样日期，请修改" unless exists $collect_hash{$report_id};
			#$mw->messageBox(-message=>'生产信息中没有填写采样日期，请注意',-type=>'ok') unless $collect_hash{$report_id};
			#print "没有填写送患者性别，请修改" unless exists $gender_hash{$report_id};
			#$mw->messageBox(-message=>'生产信息中没有填写送患者性别，请注意',-type=>'ok') unless $gender_hash{$report_id};
			#print "没有填写送患者年龄，请修改" unless exists $age_hash{$report_id};
			#$mw->messageBox(-message=>'生产信息中没有填写送患者年龄，请注意',-type=>'ok') unless $age_hash{$report_id};

		}
	}
	my $parser    = Spreadsheet::ParseXLSX->new();
	$missiondir=encode("GB2312",$missiondir);
	my $workbook = $parser ->parse("$missiondir");
	#my $workbook = $parser ->parse('D:\CTAB\epilepsy\test\Bplan.xlsx');
	unless ($workbook){
		print "无$workbook 文件！";
		next;
	}
	my $mission_day;
	my %mission =();
	my @report_id_array=();
	my $report_id_array_var;	#将所有的任务单里面的报告单编号整合到一个变量中，用于后面匹配，若匹配不到则导出结果中存在不正确的任务
	for my $worksheet ($workbook -> worksheets()){
		#my $request_sheet = $workbook -> worksheet();
		my ($row_min, $row_max) = $worksheet -> row_range();
		my ($col_min, $col_max) = $worksheet -> col_range();
		#print $row_min,"\t",$row_max,"\t",$col_min,"\t",$col_max,"\n";
		for my $row ($row_min..$row_max){
			for my $col ( $col_min .. $col_max ){
				my $cell = $worksheet->get_cell( $row, $col );
				next unless $cell;
				my $cell = $worksheet ->get_cell($row_min+6,$col_min+2);
				$operator = encode ('GB2312', $cell-> value());
				my $cell = $worksheet->get_cell($row_min+1,$col_min+5);
				$mission_day = encode ('GB2312', $cell-> value());
				$mission_day=~ /(\d+)\/(\d+)\/(\d+)/;
				my ($missionyear,$missionmon,$missionday)=($3,$1,$2);
				$missionyear=sprintf("%4d", $missionyear);
				$missionmon=sprintf("%02d", $missionmon);
				$missionday=sprintf("%02d", $missionday);
				$mission_day = $missionyear.'-'.$missionmon.'-'.$missionday;
				for my $i (($row_min+8)..$row_max){
					my $cell = $worksheet ->get_cell($i,$col_min+2);
					$lab_id   = encode ('GB2312', $cell-> value());
					my $cell = $worksheet ->get_cell($i,$col_min+3);  
					$report_id   = encode ('GB2312', $cell-> value());
					push @report_id_array,$report_id if $report_id ;
					my $cell = $worksheet ->get_cell($i,$col_min+4);
					$patient   = encode ('GB2312', $cell-> value());
					$lab_report{$lab_id} = $report_id;
					$report_patient{$report_id} = $patient;
					$mission{$report_id} = $mission_day;
				}
			}
		}
	}
	$report_id_array_var = join '',@report_id_array;
	#print $fh $report_id_array_var,"\n";
	$Typingdir=encode("GB2312",$Typingdir);
	#$Typingdir=decode("GB2312",$Typingdir);
	#my $parser    = Spreadsheet::ParseXLSX->new();
	my $parser;
	if($Typingdir=~ /x$/){
		$parser    = Spreadsheet::ParseXLSX->new();
	}
	elsif($Typingdir=~ /s$/){
		$parser    = Spreadsheet::ParseExcel->new();
	}
	else{
		$mw->messageBox(-message=>'导出结果文件格式不对，请返回上一步重新选择.xlsx或.xls格式文件',-type=>'ok');
	}
	my $workbook = $parser->parse("$Typingdir");
	if ( !defined $workbook ) {
		die $parser->error(), ".\n";
	}
	#print "no file \n" unless $workbook;
	$mw->messageBox(-message=>'没有导出结果',-type=>'ok') unless $workbook;
	my ($pre_report_date,$report_date,$result,$sample_type,$num,@samples,$report_alias,$result_report,$result_exper,@sample_types); 
	my ($candidate_1,$candidate_2,$candidate_1_alias,$candidate_2_alias,$type_B_1,$type_B_2,@unique_samples,$year);
	my %HLA_type=();
	my %AorB    =();
	my %alias   =();
	my %typing_A_results=();
	my %typing_B_results=();
	#my %count   =();
	for my $worksheet ($workbook -> worksheets()){
		my ($row_min, $row_max) = $worksheet -> row_range();
		my ($col_min, $col_max) = $worksheet -> col_range();
		for my $row($row_min..$row_max){	#从第1行开始读取，直到没有了样本名为止的行
			my $cell = $worksheet->get_cell($row,0); #读第一列      
			next unless $cell;
			$sample_type = $cell -> value();
			push @sample_types, $sample_type;		#将所有的样本名集中在一起，找到样本名一样的，接下来设为一组
			#print OUT "(",$row,",",0,")","\t",$sample_type,"\n";
			my @tmp = split /\-/,$sample_type;
			$report_alias = $tmp[0];
			$alias{$sample_type} = $report_alias;	#将全称和简称分别作为%alias的键值
			$result_report = 'CSTB'.$report_alias;	#报告单编号
			$mw->messageBox(-message=>'报告单编号缺少或多余，请检查是否选择错误的导出文件，回到上一步',-type=>'ok') if $report_id_array_var !~ /$result_report/;
			$result_exper  =  $report_alias;		#样本编号
			my $cell = $worksheet->get_cell($row,4);		#一列一列读
			next unless $cell;
			$candidate_1 = $cell -> value();
			$candidate_1 =~ /(\w+)\*(\d+)\:(\d+)/;	#$candidate_1为typingresult文件中的样本名如EM00024-A-0701K-A01
			$candidate_1_alias = $2."\:".$3;	#获得HLA-typing的结果用于与对照表匹配
			$AorB{$candidate_1_alias} = $1;	#获得与HLA-typing结果对应的外显子型，如A或B
			my $cell = $worksheet->get_cell($row,5);
			next unless $cell;
			$candidate_2 = $cell -> value();
			$candidate_2 =~ /(\w+)\*(\d+)\:(\d+)/;
			$candidate_2_alias = $2."\:".$3;
			$AorB{$candidate_2_alias} = $1;
			
			if ($AorB{$candidate_1_alias} eq 'A' and $AorB{$candidate_2_alias} eq 'A'){		#2列都为A*
				push @{$typing_A_results{$report_alias}},$candidate_1_alias;
				push @{$typing_A_results{$report_alias}},$candidate_2_alias;
			}
			elsif($AorB{$candidate_1_alias} eq 'B' and $AorB{$candidate_2_alias} eq 'B'){	#2列都为B*
				push @{$typing_B_results{$report_alias}},$candidate_1_alias;
				push @{$typing_B_results{$report_alias}},$candidate_2_alias;
			}
			else {
				next;
			}
			
			$HLA_type{$candidate_1_alias}{$candidate_2_alias} = $report_alias;	#用于确保A-B个为一对
			next unless $HLA_type{$candidate_1_alias}{$candidate_2_alias};
			#next if $HLA_type{$candidate_1_alias}{$candidate_2_alias};
			push @samples,$report_alias;	#将简称合并到一个数组中
			
			#print OUT $HLA_type{$candidate_1_alias}{$candidate_2_alias},"\t",$candidate_1_alias,"\t",$candidate_2_alias,"\t",$AorB{$candidate_1_alias},"\t","@{$typing_A_results{$report_alias}}","\n";
		}
		my %count =();
		foreach $report_alias(@samples){	#去除重复，使每一个样都有4个结果，样本名为键，值为结果组成的数组
			$count{$report_alias} += 1;
			if ($count{$report_alias} == 1){
				push @unique_samples, $report_alias;
			}
			else {next;}
		}
		#print "@unique_samples\n";
	}	
		my %level   =();
		my %miaoshu =();
		my %miaoshu_result =(); #每个样本名为键，存在描述的结果为值数组
		my %level_result =();	#每个样本名为键，配型的结果为值数组
		my %miaoshu_level = (); #以不同的型对应的级别为键，描述为值，最后输出对应的描述
		my %PosorNeg =();
		my %yinyang = ();
		my %jingshi =();
		my $B_type = '15:0215:0138:0215:1135:01';
		my $B_type2= '15:0115:11';
		my $B_type3= '38:0235:01';
		my $B_type4= '15:0138:0215:1135:01';
		my $A_positive_type = '24:0231:0102:01';
		
		my $level1 = ' 低风险 ';
		my $level2 = ' 提醒注意 ';
		my $level3 = ' 警示 ';
		my $level4 = ' 严重警示 ';
		#my $level5 = '警示';	#若出现2个A阳性(24:02|31:01|02:01)的时候，仍然报警示
		
		my $describe1 = " 禁止使用卡马西平。尽量避免使用其它芳香族抗癫痫药物，如无其它有效抗癫痫药物时，需征得患者或家属同意，减少起始剂量使用一段时间 ";
		my $describe2 = ' 使用芳香族抗癫痫药物，可能出现轻型过敏 ';
		my $describe3 = " 慎用或避免使用芳香族抗癫痫药物。如无其它有效抗癫痫药物时，需征得患者或家属同意，减少起始剂量使用一段时间 ";
		my $describe4 = ' 可以使用芳香族抗癫痫药物，但不排除其它因素引起的过敏 ';
		#my $describe5 = '无。';
		
		my $positive  = ' 阳性 ';
		my $negative  = ' 阴性 ';
		
	foreach my $name(@unique_samples){
		my ($HLA_A_typing,$HLA_B_typing);
		my ($i,$j)=0;
		#print OUT $name,"\t","@{$typing_A_results{$name}}","\t","@{$typing_B_results{$name}}","\n";
		
		my %PosorNeg_value =();	
		my (%HLA_A_pos,%HLA_B_pos);
		foreach my $HLA_A_allele (@{$typing_A_results{$name}}){		
			if ($A_positive_type =~ /$HLA_A_allele/){
				$PosorNeg_value{$name} +=1;
				push @{$HLA_A_pos{$name}},$HLA_A_allele;
			}
			else{next;}
		}
		foreach my $HLA_B_allele (@{$typing_B_results{$name}}){		
			if ($B_type =~ /$HLA_B_allele/){
				$PosorNeg_value{$name} +=3;
				push @{$HLA_B_pos{$name}},$HLA_B_allele;
			}
			else{next;}
		}
		if ($PosorNeg_value{$name} ==1){
			$PosorNeg{$name} = $positive;
			$yinyang{$name} = "1.HLA-A\*"."${$HLA_A_pos{$name}}[0]"."基因筛查结果为";
			#$yinyang{$name} = "1.HLA-A\*24\:02基因筛查结果为 $PosorNeg{$name}。";
			#$yinyang{$name} = "1.HLA-A\*24\:02基因筛查结果为 ";
		}
		elsif ($PosorNeg_value{$name}==2){
			$PosorNeg{$name} = $positive;
			$yinyang{$name} = "1.HLA-A\*"."${$HLA_A_pos{$name}}[0]"."基因筛查结果为" if ${$HLA_A_pos{$name}}[0] eq ${$HLA_A_pos{$name}}[1];
			$yinyang{$name} = "1.HLA-A\*"."${$HLA_A_pos{$name}}[0]".'和'."HLA-A\*"."${$HLA_A_pos{$name}}[1]"."基因筛查结果为" if ${$HLA_A_pos{$name}}[0] ne ${$HLA_A_pos{$name}}[1];
		}
		elsif($PosorNeg_value{$name} ==3) {
			$PosorNeg{$name} = $positive;
			#$yinyang{$name} = "1.HLA-A\*15\:02基因筛查结果为 $PosorNeg{$name}。";
			$yinyang{$name} = "1.HLA-B\*"."${$HLA_B_pos{$name}}[0]"."基因筛查结果为";
		}
		elsif($PosorNeg_value{$name} ==6){
			$PosorNeg{$name} = $positive;
			$yinyang{$name} = "1.HLA-B\*"."${$HLA_B_pos{$name}}[0]"."基因筛查结果为" if ${$HLA_B_pos{$name}}[0] eq ${$HLA_B_pos{$name}}[1];
			$yinyang{$name} = "1.HLA-B\*"."${$HLA_B_pos{$name}}[0]".'和'."HLA-B\*"."${$HLA_B_pos{$name}}[1]"."基因筛查结果为" if ${$HLA_B_pos{$name}}[0] ne ${$HLA_B_pos{$name}}[1];
		}
		elsif($PosorNeg_value{$name} ==4){
			$PosorNeg{$name} = $positive;
			#$yinyang{$name}="1.HLA-A和HLA-B基因筛查结果为 $PosorNeg{$name}。";
			$yinyang{$name} = "1.HLA-A\*"."${$HLA_A_pos{$name}}[0]".'和'."HLA-B\*"."${$HLA_B_pos{$name}}[0]"."基因筛查结果为";
		}
		elsif($PosorNeg_value{$name} ==5){
			$PosorNeg{$name} = $positive;
			$yinyang{$name} = "1.HLA-A\*"."${$HLA_A_pos{$name}}[0]".'和'."HLA-B\*"."${$HLA_B_pos{$name}}[0]"."基因筛查结果为" if ${$HLA_A_pos{$name}}[0] eq ${$HLA_A_pos{$name}}[1];
			$yinyang{$name} = "1.HLA-A\*"."${$HLA_A_pos{$name}}[0]".','."HLA-A\*"."${$HLA_A_pos{$name}}[1]".'和'."HLA-B\*"."${$HLA_B_pos{$name}}[0]"."基因筛查结果为" if ${$HLA_A_pos{$name}}[0] ne ${$HLA_A_pos{$name}}[1];
		}
		elsif($PosorNeg_value{$name} ==7){
			$PosorNeg{$name} = $positive;
			$yinyang{$name} = "1.HLA-A\*"."${$HLA_A_pos{$name}}[0]".'和'."HLA-B\*"."${$HLA_B_pos{$name}}[0]"."基因筛查结果为" if ${$HLA_B_pos{$name}}[0] eq ${$HLA_B_pos{$name}}[1];
			$yinyang{$name} = "1.HLA-A\*"."${$HLA_A_pos{$name}}[0]".','."HLA-B\*"."${$HLA_B_pos{$name}}[0]".'和'."HLA-B\*"."${$HLA_B_pos{$name}}[1]"."基因筛查结果为" if ${$HLA_B_pos{$name}}[0] ne ${$HLA_B_pos{$name}}[1];
		}
		elsif($PosorNeg_value{$name} ==8){
			$PosorNeg{$name} = $positive;
			$yinyang{$name} = "1.HLA-A\*"."${$HLA_A_pos{$name}}[0]".'和'."HLA-B\*"."${$HLA_B_pos{$name}}[0]"."基因筛查结果为" if ${$HLA_A_pos{$name}}[0] eq ${$HLA_A_pos{$name}}[1] and ${$HLA_B_pos{$name}}[0] eq ${$HLA_B_pos{$name}}[1];
			$yinyang{$name} = "1.HLA-A\*"."${$HLA_A_pos{$name}}[0]".','."HLA-B\*"."${$HLA_B_pos{$name}}[0]".'和'."HLA-B\*"."${$HLA_B_pos{$name}}[1]"."基因筛查结果为" if ${$HLA_A_pos{$name}}[0] eq ${$HLA_A_pos{$name}}[1] and ${$HLA_B_pos{$name}}[0] ne ${$HLA_B_pos{$name}}[1];
			$yinyang{$name} = "1.HLA-A\*"."${$HLA_A_pos{$name}}[0]".','."HLA-B\*"."${$HLA_A_pos{$name}}[1]".'和'."HLA-B\*"."${$HLA_B_pos{$name}}[0]"."基因筛查结果为" if ${$HLA_A_pos{$name}}[0] ne ${$HLA_A_pos{$name}}[1] and ${$HLA_B_pos{$name}}[0] eq ${$HLA_B_pos{$name}}[1];
			$yinyang{$name} = "1.HLA-A\*"."${$HLA_A_pos{$name}}[0]".','."HLA-A\*"."${$HLA_A_pos{$name}}[1]".','."HLA-B\*"."${$HLA_B_pos{$name}}[0]".'和'."HLA-B\*"."${$HLA_B_pos{$name}}[1]"."基因筛查结果为" if ${$HLA_A_pos{$name}}[0] ne ${$HLA_A_pos{$name}}[1] and ${$HLA_B_pos{$name}}[0] ne ${$HLA_B_pos{$name}}[1];
		}
		else {
			$PosorNeg{$name} = $negative;
			#$yinyang{$name} = "1.HLA-A和HLA-B基因筛查结果为 $PosorNeg{$name}。";
			$yinyang{$name} = "1.HLA-A和HLA-B位点的基因筛查结果均为";
		}
		#print OUT $PosorNeg{$name},"\n";
		
		for $i($i <= $#{$typing_A_results{$name}},$i++) {
			for $j($j <= $#{$typing_B_results{$name}},$j++){
				if (${$typing_A_results{$name}}[$i] eq '24:02' and $B_type =~ /${$typing_B_results{$name}}[$j]/){
					$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $level4;
					$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $describe1;
					$miaoshu_level{$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]}}=$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				}
				elsif(${$typing_A_results{$name}}[$i] eq '24:02' and $B_type !=~ /${$typing_B_results{$name}}[$j]/){
					$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $level3;
					$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $describe3;
					$miaoshu_level{$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]}}=$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				}
				elsif(${$typing_A_results{$name}}[$i] eq '31:01' and $B_type2 =~ /${$typing_B_results{$name}}[$j]/){
					$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $level3;
					$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $describe3;
					$miaoshu_level{$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]}}=$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				}
				elsif(${$typing_A_results{$name}}[$i] eq '31:01' and $B_type3 =~ /${$typing_B_results{$name}}[$j]/){
					$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $level2;
					$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $describe2;
					$miaoshu_level{$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]}}=$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				}
				elsif(${$typing_A_results{$name}}[$i] eq '31:01' and $B_type3 eq '15:02'){
					$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $level4;
					$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $describe1;
					$miaoshu_level{$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]}}=$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				}
				elsif(${$typing_A_results{$name}}[$i] eq '31:01' and $B_type !=~ /${$typing_B_results{$name}}[$j]/){
					$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $level2;
					$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $describe2;
					$miaoshu_level{$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]}}=$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				}
				elsif(${$typing_A_results{$name}}[$i] eq '02:01' and $B_type4=~ /${$typing_B_results{$name}}[$j]/){
					$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $level3;
					$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $describe3;
					$miaoshu_level{$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]}}=$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				}
				elsif(${$typing_A_results{$name}}[$i] eq '02:01' and $B_type !=~ /${$typing_B_results{$name}}[$j]/){
					$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $level3;
					$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $describe3;
					$miaoshu_level{$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]}}=$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				}
				elsif(${$typing_B_results{$name}}[$j] eq '15:02'){
					$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $level4;
					$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $describe1;
					$miaoshu_level{$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]}}=$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				}
				elsif(${$typing_B_results{$name}}[$j] eq '15:01' or ${$typing_B_results{$name}}[$j] eq '15:11'){
					$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $level3;
					$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $describe3;
					$miaoshu_level{$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]}}=$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				}
				elsif(${$typing_B_results{$name}}[$j] eq '38:02' or ${$typing_B_results{$name}}[$j] eq '35:01'){
					$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $level2;
					$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $describe2;
					$miaoshu_level{$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]}}=$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				}
				else{
					$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $level1;
					$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]} = $describe4;
					$miaoshu_level{$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]}}=$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				}
				push @{$level_result{$name}},$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				push @{$miaoshu_result{$name}},$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]};
				#print OUT $name,"\t",${$typing_A_results{$name}}[$i],"\t",${$typing_B_results{$name}}[$j],"\t",$level{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]},"\t",$miaoshu{${$typing_A_results{$name}}[$i]}{${$typing_B_results{$name}}[$j]},"\n";
				
				
			}
		}
		
		#}
		#print OUT "@{$level_result{$name}}","\n";	# $level_result{$name} 为级别组成的数组
		my %value =();
		$value{$level1} =1;
		$value{$level2} =2;
		$value{$level3} =3;
		$value{$level4} =4;
		#$value{$level5} =5;
		my @level_all	=();
		push @level_all,$level1;
		push @level_all,$level2;
		push @level_all,$level3;
		push @level_all,$level4;
		#push @level_all,$level5;
	
		my %LevelValue =();
		my %max_level_value=();
		
		foreach my $a (@{$level_result{$name}}){
			push @{$LevelValue{$name}},$value{$a};
		}
		$max_level_value{$name} = &max(@{$LevelValue{$name}});	#最大的值
		foreach my $b(@level_all){
			if ($value{$b} == $max_level_value{$name}) {
				$jingshi{$name} =$b;
			}
			else{
				next;
			}
		}
		#print OUT $jingshi{$name},"\n";	#每个样本的最终警示等级 <- 重要信息
		#print OUT "@{$miaoshu_result{$name}}\n";
		
		#print OUT $miaoshu_level{$jingshi{$name}},"\n";	#每个样本最终警示等级对应的描述 <-重要信息
	}	
	sub max{
		my $max = shift @_;
		foreach (@_){
			if ($_ > $max){
				$max =$_;
			}
		}
		$max;
	}
	
	#########################################################生成报告,每个样本生成一个sheet
	
	#my $workbook = Excel::Writer::XLSX->new( "D\:\\CTAB\\epilepsy\\test\\$dateXXX.xlsx" );	#在目的文件夹下生成目的文件，后期可根据用户需求随时修改
	my $workbook = Excel::Writer::XLSX->new( "$Destdir" );
	my $format1 = $workbook -> add_format();
	$format1 ->set_bold();
	$format1 ->set_color('black');
	$format1 ->set_size(15);
	$format1 ->set_font('宋体');
	$format1 ->set_align( 'center' );
	$format1 ->set_align('vcenter');
	
	my $format2 = $workbook -> add_format();
	$format2 ->set_bold();
	$format2 ->set_color('black');
	$format2 ->set_size(9);
	$format2 ->set_font('Times New Roman');
	$format2 ->set_align( 'right' );
	$format2 ->set_align('vcenter');
	
	my $format3 = $workbook -> add_format();
	$format3 ->set_color('black');
	$format3 ->set_size(9);
	$format3 ->set_font('宋体');
	$format3 ->set_align( 'left' );
	$format3 ->set_align('vcenter');
	
	my $format3_1 = $workbook -> add_format(color=>'black',size=>9,font=>'宋体',align=>'left');
	my $format3_2 = $workbook -> add_format(color=>'black',size=>9,font_script=>1);
	
	my $format4 = $workbook -> add_format();
	$format4 ->set_color('black');
	$format4 ->set_size(9);
	$format4 ->set_font('宋体');
	$format4 ->set_align('center');
	$format4 ->set_align('vcenter');
	
	my $format5 = $workbook -> add_format();
	$format5 ->set_bold();
	$format5 ->set_color('black');
	$format5 ->set_size(12);
	$format5 ->set_font('宋体');
	$format5 ->set_align('left');
	$format5 ->set_align('vcenter');
	
	my $format5_1 = $workbook -> add_format(bold=>1,color=>'black',size=>12,font=>'宋体',align=>'left');
	my $format5_2 = $workbook -> add_format(bold=>1,color=>'black',size=>12,font=>'宋体',align=>'left',underline=>1);
	
	my $format6 = $workbook -> add_format();
	$format6 ->set_bold();
	$format6 ->set_border(2);
	$format6 ->set_color('black');
	$format6 ->set_size(12);
	$format6 ->set_font('宋体');
	#$format6 ->set_font('Times New Roman');
	$format6 ->set_align('center');
	$format6 ->set_align('vcenter');
	
	my $format7 = $workbook -> add_format();
	$format7 ->set_bold();
	$format7 ->set_border(2);
	$format7 ->set_color('black');
	$format7 ->set_size(10.5);
	$format7 ->set_font('Times New Roman');
	$format7 ->set_align('center');
	$format7 ->set_align('vcenter');
	
	my $format8 = $workbook -> add_format();
	$format8 ->set_color('black');
	$format8 ->set_border(2);
	$format8 ->set_size(12);
	$format8 ->set_font('Times New Roman');
	$format8 ->set_align('center');
	$format8 ->set_align('vcenter');
	
	my $format8_1 = $workbook -> add_format(bold=>1,size=>12,font=>'Times New Roman',bg_color=>'yellow',color=>'black');
	my $format8_2 = $workbook -> add_format(size=>12,font=>'Times New Roman',color=>'black');
	
	my $format8_3 = $workbook -> add_format();
	$format8_3 ->set_color('black');
	$format8_3 ->set_pattern();
	$format8_3 ->set_bold();
	$format8_3 ->set_bg_color('yellow');
	$format8_3 ->set_bottom(2);
	$format8_3 ->set_size(12);
	$format8_3 ->set_font('Times New Roman');
	$format8_3 ->set_align('right');
	$format8_3 ->set_align('vcenter');
	
	
	my $format8_4 = $workbook -> add_format();
	$format8_4 ->set_color('black');
	$format8_4 ->set_bottom(2);
	$format8_4 ->set_size(12);
	$format8_4 ->set_font('Times New Roman');
	$format8_4 ->set_align('right');
	$format8_4 ->set_align('vcenter');
	
	my $format8_5 = $workbook -> add_format();
	$format8_5 ->set_color('black');
	$format8_5 ->set_pattern();
	$format8_5 ->set_bold();
	$format8_5 ->set_bg_color('yellow');
	$format8_5 ->set_bottom(2);
	$format8_5 ->set_size(12);
	$format8_5 ->set_font('Times New Roman');
	$format8_5 ->set_align('left');
	$format8_5 ->set_align('vcenter');
	
	
	my $format8_6 = $workbook -> add_format();
	$format8_6 ->set_color('black');
	$format8_6 ->set_bottom(2);
	$format8_6 ->set_size(12);
	$format8_6 ->set_font('Times New Roman');
	$format8_6 ->set_align('left');
	$format8_6 ->set_align('vcenter');
	
	my $format9 = $workbook -> add_format();
	$format9 ->set_bold();
	$format9 ->set_color('black');
	$format9 ->set_size(12);
	$format9 ->set_font('宋体');
	$format9 ->set_align('left');
	$format9 ->set_align('vcenter');
	$format9 ->set_underline(1);
	$format9 ->set_text_wrap(1);
	
	my $format9_1 = $workbook -> add_format(bold => 1,color=>'black',size=>12,font=>'宋体',align=>'left');
	my $format9_2 = $workbook -> add_format(bold => 1,color=>'black',size=>12,font=>'宋体',align=>'left',underline=>1);
	
	my $format10 = $workbook -> add_format();
	$format10 ->set_color('black');
	$format10 ->set_size(9);
	$format10 ->set_font('Times New Roman');
	$format10 ->set_align('left');
	$format10 ->set_align('vcenter');
	$format10 ->set_text_wrap(1);
	
	my $format11 = $workbook -> add_format();
	$format11 ->set_bold();
	$format11 ->set_color('black');
	$format11 ->set_size(16);
	$format11 ->set_font('Times New Roman');
	$format11 ->set_align( 'center' );
	$format11 ->set_align('vcenter');
	
	my $format12 = $workbook -> add_format();
	$format12 ->set_bold();
	$format12 ->set_color('black');
	$format12 ->set_size(10.5);
	#$format12 ->set_font('宋体');
	$format12 ->set_font('Times New Roman');
	$format12 ->set_align('left');
	$format12 ->set_align('vcenter');
	
	my $format13 = $workbook -> add_format();
	$format13 ->set_bold();
	$format13 ->set_top(1.25);
	$format13 ->set_top_color('#739cc3');
	$format13 ->set_color('black');
	$format13 ->set_size(10.5);
	$format13 ->set_font('宋体');
	$format13 ->set_align('left');
	$format13 ->set_align('vcenter');
	
	my $format13_1 = $workbook -> add_format();
	$format13_1 ->set_bold();
	$format13_1 ->set_top(1.25);
	$format13_1 ->set_top_color('#739cc3');
	$format13_1 ->set_color('black');
	$format13_1 ->set_size(10.5);
	$format13_1 ->set_font('Times New Roman');
	$format13_1 ->set_align('center');
	$format13_1 ->set_align('vcenter');

	my $format14 = $workbook -> add_format();
	$format14 ->set_bold();
	$format14 ->set_bottom(1.25);
	$format14 ->set_bottom_color('#739cc3');
	$format14 ->set_color('black');
	$format14 ->set_size(10.5);
	$format14 ->set_font('宋体');
	$format14 ->set_align('left');
	$format14 ->set_align('vcenter');
	
	my $format15 = $workbook -> add_format();
	$format15 ->set_bold();
	$format15 ->set_bottom(1.5);
	$format15 ->set_bottom_color('#739cc3');
	$format15 ->set_color('black');
	$format15 ->set_size(10.5);
	#$format15 ->set_font('宋体');
	$format15 ->set_font('Times New Roman');
	$format15 ->set_align('center');
	$format15 ->set_align('vcenter');
	
	my $format16 = $workbook -> add_format();
	$format16 ->set_bold();
	$format16 ->set_top(1.25);
	$format16 ->set_top_color('#739cc3');
	$format16 ->set_color('black');
	$format16 ->set_size(10.5);
	$format16 ->set_font('宋体');
	$format16 ->set_align('left');
	$format16 ->set_align('vcenter');
	
	my $format17 = $workbook -> add_format();
	$format17 ->set_bold();
	$format17 ->set_bottom(1.25);
	$format17 ->set_bottom_color('#739cc3');
	$format17 ->set_color('black');
	$format17 ->set_size(10.5);
	$format17 ->set_font('宋体');
	$format17 ->set_align('left');
	$format17 ->set_align('vcenter');
	
	my $format18 = $workbook -> add_format();
	$format18 ->set_color('black');
	$format18 ->set_size(9);
	$format18 ->set_font('宋体');
	$format18 ->set_align( 'left' );
	$format18 ->set_align('vcenter');
	$format18 ->set_font_script(1);
	
	my $format19 = $workbook -> add_format();
	$format19 ->set_bold();
	$format19 ->set_color('black');
	$format19 ->set_size(12);
	$format19 ->set_font('黑体');
	$format19 ->set_align('left');
	$format19 ->set_align('vcenter');
	
	
	my $format19_1 = $workbook -> add_format(bold=>1,size=>12,color=>'black',font=>'黑体',align=>'left');
	my $format19_2 = $workbook -> add_format(bold=>1,size=>12,color=>'black',font=>'宋体',align=>'left');
	
	my $format20 = $workbook -> add_format();	#上边框带颜色（不是黑色）
	$format20 ->set_top(1.25);
	$format20 ->set_top_color('#739cc3');
	
	my $format21 = $workbook -> add_format();	#下边框带颜色（不是黑色）
	$format21 ->set_bottom(1.25);
	$format21 ->set_bottom_color('#739cc3');
	
	my $format22 = $workbook -> add_format();	#上边框带颜色
	$format22 ->set_top(2);
	$format22 ->set_top_color('black');
	
	my $format23 = $workbook -> add_format();	#下边框带颜色
	$format23 ->set_bottom(2);
	$format23 ->set_bottom_color('black');
	
	my $format24 = $workbook -> add_format();	#右边和下边框带颜色
	$format24 ->set_bottom(2);
	$format24 ->set_right(2);
	$format24 ->set_bottom_color('black');
	$format24 ->set_right_color('black');
	
	my $format25 = $workbook -> add_format();	#下边框带颜色
	$format25 ->set_bottom(2);
	$format25 ->set_bottom_color('black');
	$format25 ->set_color('black');
	$format25 ->set_size(12);
	$format25 ->set_font('Times New Roman');
	$format25 ->set_align('center');
	$format25 ->set_align('vcenter');
	
	foreach my $name(@unique_samples){	#按样本数产生sheet数
		my $sheetname_full = 'CSTB'.$name;	#报告单编号
		my $sheetname      = $name;	#样本编号,如2017EM00024 
		my $worksheet = $workbook -> add_worksheet($sheetname_full);
		#$worksheet->set_margins_LR(2);
		$worksheet->center_horizontally();
		$worksheet->set_column(0,0,7.5);
		$worksheet->set_column(1,1,17);
		$worksheet->set_column(2,2,3.63);
		$worksheet->set_column(3,3,19.63);
		#$worksheet->set_column(4,4,8.75);
		$worksheet->set_column(4,4,9.25);
		$worksheet->set_column(5,5,4.88);
		$worksheet->set_column(6,6,0.40);
		$worksheet->set_column(7,7,4.88);
		#$worksheet->set_column(8,8,8.75);
		$worksheet->set_column(4,4,9.25);
		#$worksheet->set_column(9,9,7.5);
		$worksheet->set_column(9,9,7);
		$worksheet->set_margin_left(0.65);
		$worksheet->set_margin_right(0.65);
		#my @rows = (26.25,27,20,14.25,18,21,21,25,21,16,16,25,21,24,40,13.5,18.75,27.75,21,21,15,15,15,15);
		my @rows = (13.5,26.25,9.75,27,27,20,10.5,20.25,28.5,28.5,28.5,34.5,21,21,21,30,28.5,28.5,40,4.5,18.75,12,27.75,27.75,28.5,28.5,21,15,15,15,15);
		for my $i(0..$#rows){   
			$worksheet->set_row($i,$rows[$i]);
		}
		$worksheet ->merge_range('A2:J2','深圳荻硕贝肯精准医学有限公司',$format1);
		$worksheet ->merge_range('A4:J4','抗癫痫药物过敏HLA基因筛查与咨询报告',$format11);
		$worksheet ->merge_range('E6:J6',"报告单编号：$sheetname_full",$format2);
		$worksheet ->merge_range('A8:E8','地址：深圳市大鹏新区葵涌镇金业大道140号（深圳生命科学产业园）A17栋',$format3);
		$worksheet ->merge_range('G8:J8','邮政编码：518116',$format4);
	
		my $patientname = $report_patient{$sheetname_full};
		$patientname = decode('GB2312',$patientname);
		$patientname = '患者姓名：'.$patientname;
		$worksheet ->merge_range('A9:B9',$patientname,$format13);
		
		$worksheet -> write('C9','',$format20);
	
		my $gender_opt = $gender_hash{$sheetname_full};
		$gender_opt = decode('GB2312',$gender_opt);
		$gender_opt = '性别：'.$gender_opt;
		$worksheet -> write('D9',$gender_opt,$format13);
	
		#$worksheet ->merge_range('E9:F9',"年龄：$age_hash{$sheetname_full}岁",$format13) unless $age_hash{$sheetname_full} =~ /(\d+)M/;
		if($age_hash{$sheetname_full} =~ /^(\d+)M$/){
			$age_hash{$sheetname_full} =~ s/M//;
			$worksheet ->merge_range('E9:F9',"年龄：$age_hash{$sheetname_full}个月",$format13);
		}
		elsif($age_hash{$sheetname_full} =~ /^(\d+)Y(\d+)M/){
			$worksheet ->merge_range('E9:F9',"年龄：$1岁$2个月",$format13);
		}
		elsif($age_hash{$sheetname_full} =~ /^(\d+)Y$/){
			$age_hash{$sheetname_full} =~ s/Y//;
			$worksheet ->merge_range('E9:F9',"年龄：$age_hash{$sheetname_full}岁",$format13);
		}
		elsif($age_hash{$sheetname_full} =~ /^(\d+)$/){
			$worksheet ->merge_range('E9:F9',"年龄：$age_hash{$sheetname_full}岁",$format13);
		}
		elsif($age_hash{$sheetname_full} =~ /^(\d+)M(\d+)(D|d)$/){
			$worksheet ->merge_range('E9:F9',"年龄：$1个月$2天",$format13);
		}
		elsif($age_hash{$sheetname_full} =~ /^(\d+)(D|d)$/){
			$age_hash{$sheetname_full} =~ s/(D|d)//;
			$worksheet ->merge_range('E9:F9',"年龄：$age_hash{$sheetname_full}天",$format13);
		}
		$worksheet ->merge_range('G9:J9','ID信息：',$format13_1);
	
		my $hospital_opt = $hosptal_hash{$sheetname_full};
		$hospital_opt = decode('GB2312',$hospital_opt);
		#if(exists $alias{$hospital_opt}){
		#	$hospital_opt = '送检医院：'.$alias{$hospital_opt};
		#}else{
		#	$hospital_opt = '送检医院：'.$hospital_opt;
		#}
		$hospital_opt = '送检医院：'.$hospital_opt;
		$worksheet ->merge_range('A10:D10',$hospital_opt,$format12);
		my $department_opt = $department_hash{$sheetname_full};
		$department_opt= decode('GB2312',$department_opt);
		$department_opt = '送检科室：'.$department_opt;
		$worksheet ->merge_range('E10:H10',$department_opt,$format12);
		my $doctor_opt = $doctor_hash{$sheetname_full};
		$doctor_opt = decode('GB2312',$doctor_opt);
		$doctor_opt = '送检医师：'.$doctor_opt;
		$worksheet ->merge_range('I10:J10',$doctor_opt,$format12);
		
		$worksheet ->merge_range('A11:B11',"采样日期：$collect_hash{$sheetname_full}",$format14);
		$worksheet ->write('C11','',$format21);
		$worksheet ->write('D11',"收样日期：$receive_hash{$sheetname_full}",$format14);
		$worksheet ->merge_range('E11:J11',"样本编号：$report_sheet{$sheetname_full}",$format15);
		$worksheet ->merge_range_type('rich_string','A12:D12',$format19_1,'HLA—A、B位点高分辨基因',$format19_2,'筛查',$format19_1,'结果：',$format19);
		$worksheet ->merge_range('B13:C13','样本编号',$format6);
		$worksheet ->write('D13','位点',$format6);
		$worksheet ->merge_range('E13:I13','基因型结果',$format6);
		#$worksheet ->merge_range('A10:A11',$report_sheet{$sheetname},$format7);
		$worksheet ->merge_range('B14:C15',$sheetname,$format7);
		$worksheet ->write('D14','HLA-A',$format8);
		$worksheet ->write('E14','',$format23);
		$worksheet ->write('D15','HLA-B',$format8);
		$worksheet ->write('E15','',$format23);
		#my $HLA_A = "@{$typing_A_results{$name}}";
		#$HLA_A =~ s/\s/\, /;
		my $HLA_A_1 = ${$typing_A_results{$name}}[0];
		my $HLA_A_2 = ${$typing_A_results{$name}}[1];
		if($A_positive_type =~ /$HLA_A_1/){
			$worksheet ->write('F14',$HLA_A_1,$format8_3);	#加粗加亮的模式
		}
		else{
			$worksheet ->write('F14',$HLA_A_1,$format8_4);
		}
		if($A_positive_type =~ /$HLA_A_2/){
			$worksheet ->write('H14',$HLA_A_2,$format8_5);	#加粗加亮的模式
		}
		else{
			$worksheet ->write('H14',$HLA_A_2,$format8_6);
		}
		$worksheet ->write('G14',',',$format25);

		my $HLA_B_1 = ${$typing_B_results{$name}}[0];
		my $HLA_B_2 = ${$typing_B_results{$name}}[1];
		if($B_type =~ /$HLA_B_1/){
			$worksheet ->write('F15',$HLA_B_1,$format8_3);	#加粗加亮的模式
		}
		else{
			$worksheet ->write('F15',$HLA_B_1,$format8_4);
		}
		if($B_type =~ /$HLA_B_2/){
			$worksheet ->write('H15',$HLA_B_2,$format8_5);	#加粗加亮的模式
		}
		else{
			$worksheet ->write('H15',$HLA_B_2,$format8_6);
		}
		$worksheet ->write('G15',',',$format25);

		$worksheet ->write('I14','',$format24);
		$worksheet ->write('I15','',$format24);
		#$worksheet ->write('C14',$HLA_A,$format8);
		#$worksheet ->write('C15',$HLA_B,$format8);
		$worksheet ->merge_range('A16:B16','结果说明：',$format19);
		#$worksheet ->merge_range('A17:D17',"1.HLA-A\*24\:02基因筛查结果为 $PosorNeg{$name}",$format5);	#
		$worksheet ->merge_range_type('rich_string','A17:J17',$format5_1,"$yinyang{$name}",$format5_2,"$PosorNeg{$name}",$format5_1,'。',$format5);
		$worksheet ->merge_range_type('rich_string','A18:J18',$format9_1,"2.抗癫痫药物预警级别为", $format9_2,"$jingshi{$name}",$format9_1,"。",$format5);
		#$worksheet ->write_rich_string($format9_1,"3.建议\:",$format9_2,"$miaoshu_level{$jingshi{$name}}");
		$worksheet ->merge_range_type('rich_string','A19:J19',$format9_1,"3.建议\:",$format9_2,"$miaoshu_level{$jingshi{$name}}",$format9_1,"。",$format9);
		$worksheet ->merge_range_type('rich_string','A21:J21',$format3_1,'注：本报告的结果说明基于文献报道分析，研究基于中国大规模皮肤反应数据',$format3_2,'[1]',$format3_1,'，仅供临床医生参考。',$format3);
		$worksheet ->merge_range('A23:J23','1. Shi, Y. W., Min, F. L., Zhou, D., Qin, B., Wang, J., Hu, F. Y., ... & Zhou, L. M. (2017). HLA-A* 24: 02 as a common risk factor for antiepileptic drug–induced cutaneous adverse reactions. Neurology, 10-1212.',$format10);
		$worksheet ->merge_range('A25:C25',"检测者：$detectorname",$format16);
		#$worksheet ->write('C25','',$format20);
		$worksheet ->merge_range('D25:F25',"  审核者：$checkername",$format16);
		#$worksheet ->merge_range('E25:F25','',$format20);
		$worksheet ->merge_range('G25:J25',"签发者：$signername",$format16);
		$worksheet ->merge_range('A26:B26',"检测日期：$mission{$sheetname_full}",$format17);
		$worksheet ->write('C26','',$format21);
		$worksheet ->merge_range('D26:E26',"  复核日期：$rechedate",$format17);
		$worksheet ->write('F26','',$format21);
		$worksheet ->merge_range('G26:J26',"签发日期：$signdate",$format17);
		$worksheet ->merge_range('A28:J28','1. 实验方法为PCR-SSO，采用PCR-SBT科研试剂盒复核。',$format3);
		$worksheet ->merge_range('A29:J29','2. 本结果只对本次送检样品负责。如对结果有疑义，请在收到结果后7个工作日内与我们联系。',$format3);
		$worksheet ->merge_range('A30:J30','3. 本报告用于生物学数据比对、分析，非临床检测报告。',$format3);
		$worksheet ->merge_range('A31:J31','4. 相关疾病的诊断治疗请咨询临床医生。',$format3);	
		#print OUT $report_patient{$sheetname},"\t",$gender_hash{$sheetname},"\t",$jingshi{$sheetname},"\n";
	}
	$mw->messageBox(-message=>'报告生成完毕',-type=>'ok');
	exit;
}

MainLoop;