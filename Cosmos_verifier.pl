use strict;
use Term::ProgressBar;
use constant MAX => 2030;
use Spreadsheet::ParseExcel::SaveParser;
use Spreadsheet::ParseExcel::Simple;
my $progress = Term::ProgressBar->new({count => 100});




my(@brand,@aliases,@category,@twitter,@facebook,@topic,@in_brand,@in_category,@b_id);
my $xls = Spreadsheet::ParseExcel::Simple->read('test.xls');
  foreach my $sheet ($xls->sheets) {

     while ($sheet->has_data) {
         my @data = $sheet->next_row;
         push(@brand,$data[0]); push(@aliases,$data[1]); push(@category,$data[2]);push(@twitter,$data[3]);push(@facebook,$data[4]);push(@topic,$data[5]),push(@b_id,$data[6]);
     }
 }

my $xls = Spreadsheet::ParseExcel::Simple->read('Report.xls');
  foreach my $sheet ($xls->sheets) {

     while ($sheet->has_data) {
         my @data = $sheet->next_row;
         push(@in_brand,$data[0]);push(@in_category,$data[1]);

     }
 }

my $sheet=0;
my $parser   = new Spreadsheet::ParseExcel::SaveParser;
my $template = $parser->Parse('Report.xls');
my $format   = $template->{Worksheet}[$sheet]
                            ->{FormatNo};

my $lns=scalar(@in_brand);
my $row=1;
for(my $j=1;$j<=$#in_brand;$j++){

	# print"***$in_brand[$j]***\n";
	my $cal=(($j+1)/$lns)*100;
	$progress->update($cal);
	
	my $flg=0;
	
	for(my $i=0;$i<=$#brand;$i++){

		if($brand[$i] ne ''){

			$brand[$i]=~s/\s+/\\s\*/igs; #$in_brand[$i]=~s/\s+/\\s\*/igs;
			$brand[$i]=~s/\|//igs; $aliases[$i]=~s/\|//igs;
			# $in_brand[$j]=~s/ //igs;
			#print"Brand ***$in_brand[$j]***$brand[$i]***\n";<>;
			if($in_brand[$j]=~m/^$brand[$i]$/is){
				$flg=1;
				  # print"Brand Matched****|$in_brand[$j]|***$brand[$i]***\n";
				 $template->AddCell(0, $row,   9, "YES",     $format);
				 $template->AddCell(0, $row,   15, $b_id[$i],     $format);
				 $template->AddCell(0, $row,   16, $category[$i],     $format);
				 my @cats=split(",",$category[$i]);

					# $category[$i]=~s/\s+/\\s\*/igs;
					
					my $flgs=0;
					foreach my $cat(@cats){
						# print"***$in_category[$j]<>$category[$i]****\n";
						if($in_category[$j]=~m/^$cat$/is){

							$template->AddCell(0, $row,   10, "YES",     $format);
							&checkSocial($i);
							last;
						}
						else{
							$template->AddCell(0, $row,   10, "NO",     $format);
							&checkSocial($i);
						}
					}


				next;
			}
			else{
				my @alias_arr=split(",",$aliases[$i]);
				foreach my $alias(@alias_arr){
					$in_brand[$j]=~s/\s+/\\s\*/igs;
					# $in_brand[$j]=~s/ //igs;
					$alias=~s/^\s+//igs;
					# $in_brand[$j]=~s/\s+$//igs;
					# print"ELSE****$alias***|$in_brand[$j]|***\n";
					if($alias=~m/^$in_brand[$j]$/is){

						$flg=1;
						$template->AddCell(0, $row,   15, $b_id[$i],     $format);
						$template->AddCell(0, $row,   16, $category[$i],     $format);
						# print"Brand Matched in ELSEIF****$alias***|$in_brand[$j]|***\n";
						$template->AddCell(0, $row,   9, "YES",     $format);
						$category[$i]=~s/\s+/\\s\*/igs;
						if($in_category[$j]=~m/\b$category[$i]\b/is){
							$template->AddCell(0, $row,   10, "YES",     $format);
							&checkSocial($i);

						}else{
							$template->AddCell(0, $row,   10, "NO",     $format);
							&checkSocial($i);
						}

					}
				}	
		 	}
	}

	}
	if($flg == 0){
			$template->AddCell(0, $row,   9, "NO",     $format);
			$template->AddCell(0, $row,   10, "NO",     $format);
			$template->AddCell(0, $row,   11, "NO",     $format);
			$template->AddCell(0, $row,   12, "NO",     $format);
			$template->AddCell(0, $row,   13, "NO",     $format);	}
	
	$row++;
}
 $template->SaveAs('Report.xls');

sub checkSocial(){

	my $i=shift;


	if($twitter[$i] ne ''){ $template->AddCell(0, $row,   11, "YES",     $format); } else{ $template->AddCell(0, $row,   11, "NO",     $format); }
	
	if($facebook[$i] ne ''){ $template->AddCell(0, $row,   12, "YES",     $format); } else{ $template->AddCell(0, $row,   12, "NO",     $format); }

	if($topic[$i] != 'NULL'){ $template->AddCell(0, $row,   13, "YES",     $format); } else{ $template->AddCell(0, $row,   13, "NO",     $format); }
	
	
}
