use strict;
# use warnings;
# use Data::Dumper;
use MongoDB;
use MongoDB::Collection;
use Spreadsheet::WriteExcel;
use Spreadsheet::ParseExcel::SaveParser;
use Spreadsheet::ParseExcel::Simple;

my $conn = new MongoDB::Connection( "query_timeout" => "-1");
my $db   = $conn->BrandTest; 
my $coll = $db->brands;

my $all = $coll->find();


my (@BrandId,@BrandName,@Source,@SOurceCount,@Category,@Categorycount,@InCosmos);
while(my $dts = $all->next)
{

	push(@BrandId,$dts->{BrandId});
	push(@BrandName,$dts->{BrandName});
	push(@Source,$dts->{Source});
	push(@SOurceCount,$dts->{SourceCount});
	push(@Category,$dts->{Category});
	push(@Categorycount,$dts->{CategoryCount});
	push(@InCosmos,$dts->{InCosmos});


}
my $count=$#BrandId;
my(@in_brand,@in_category,@in_source,@in_cosmos);
my $xls = Spreadsheet::ParseExcel::Simple->read('Report.xls');
  foreach my $sheet ($xls->sheets) {

     while ($sheet->has_data) {
         my @data = $sheet->next_row;
         push(@in_brand,$data[0]);push(@in_category,$data[1]);push(@in_source,$data[2]);push(@in_cosmos,$data[3]);

     }
 }

my $sheet=0;
my $parser   = new Spreadsheet::ParseExcel::SaveParser;
my $template = $parser->Parse('Report.xls');
my $format   = $template->{Worksheet}[$sheet]
                            ->{FormatNo};


my $row=1;


for(my $k=1;$k<=$#in_brand;$k++){
	# sleep(2);
	my $bName=$in_brand[$k];
	my $category=$in_category[$k];
	my $source=$in_source[$k];
	my $incosmos=$in_cosmos[$k];
		print"***$bName***\n";
		my $catFlg=0;
		$bName=~s/\s+/ZZZ/igs;
		$bName=~s/\W//igs;
		$bName=~s/ZZZ/\\s\*/igs;

		for(my $i=0;$i<=$#BrandName;$i++){
			
			if(defined($BrandName[$i])&& ($BrandName[$i]=~m/^$bName$/is)){
				print"***$BrandName[$i]***$bName***\n";
				$bName=~s/\\s\*/ /igs;
				print"******$BrandName[$i]***\tBrand Matched\n";
				if($category eq ''){

					print"Category notfound\n";
					$template->AddCell(0, $row, 7, "$Category[$i]",  $format);$template->AddCell(0, $row, 4, "$BrandId[$i]",  $format);
					$template->AddCell(0, $row, 5, "$SOurceCount[$i]",  $format);$template->AddCell(0, $row, 8, "$Source[$i]",  $format);
					next;
				}				

				$catFlg=1;

				my @moreCategory=split(";",$category);
				foreach my $subcategory(@moreCategory){
					$subcategory=~s/\s+/\\s\*/igs;
					$subcategory=~s/\)//igs;$subcategory=~s/\(//igs;					
					if($Category[$i]=~m/$subcategory/is){

						$subcategory=~s/\\s\*/ /igs;
						print"Category Matched\n";
						
						# print".................Matching--$Source[$i]<>$source\n";
						if($source){
							my @moreSource=split(";",$source);
							foreach my $subsource(@moreSource){
								$subsource=~s/\s+/\\s\*/igs;
								if($Source[$i]!~m/$subsource/is){
									$subsource=~s/\\s\*/ /igs;
									my @SrcCount=split(",",$Source[$i]);
									# print scalar(@SrcCount);
									if(scalar(@SrcCount) > 0){
										# print"inside sourcecount\n";
										push(@SrcCount,"$subsource");
										my $srcUp=join(",",@SrcCount);
										my $source_count=scalar(@SrcCount);
										print"UPDATE1\n";
										$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Source" => "$srcUp","SOurceCount" => "$source_count"}});
										$template->AddCell(0, $row, 4, "Source Updated",  $format);$template->AddCell(0, $row, 5, "$source_count",  $format);
										$template->AddCell(0, $row, 8, "$srcUp",  $format);
										$Source[$i].=",$subsource";

									}
									else{
										print"UPDATE2\n";
										$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Source" => "$subsource"}});
										$template->AddCell(0, $row, 4, "Source Updated",  $format);
									}
									
								
								}
								else{

									print"Matched All-Dupe wit Master\n";
									$template->AddCell(0, $row, 4, "Duped",  $format);$template->AddCell(0, $row, 5, scalar(@moreSource),  $format);
									$template->AddCell(0, $row, 6, scalar(@moreCategory),  $format);
								}
							}
						}
						else{ 
							
							$template->AddCell(0, $row, 6, $Categorycount[$i],  $format); 
							$template->AddCell(0, $row, 4, "Duped",  $format); 
						}
					}

					else{

						my @catCount=split(",",$Category[$i]);
						print scalar(@catCount);
						print"TEST_COUNT<>:";
						$subcategory=~s/\\s\*/ /igs;
						if(scalar(@catCount) > 0){

							push(@catCount,"$subcategory");
							@catCount=uniq(@catCount);
							my $catup=join(",",@catCount);
							my $category_count=scalar(@catCount);
							print"$category_count\n";
							$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Category" => "$catup","CategoryCount" => "$category_count"}});
							$template->AddCell(0, $row, 4, "Category Updated",  $format);$template->AddCell(0, $row, 6, "$category_count",  $format);
							$template->AddCell(0, $row, 7, "$catup",  $format);
							$Category[$i].=",$subcategory";						
							my @moreSource=split(";",$source);
							if($source){
								foreach my $subsource(@moreSource){
									$subsource=~s/\s+/\\s\*/igs;
									if($Source[$i]!~m/$subsource/is){
										$subsource=~s/\\s\*/ /igs;
										my @SrcCount=split(",",$Source[$i]);
										print scalar(@SrcCount);
										if(scalar(@SrcCount) > 0){
											# print"inside sourcecount\n";
											push(@SrcCount,"$subsource");
											my $srcUp=join(",",@SrcCount);
											my $source_count=scalar(@SrcCount);
											print"UPDATE4\n";
											$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Source" => "$srcUp","SOurceCount" => "$source_count"}});
											$template->AddCell(0, $row, 4, "Source Updated",  $format);$template->AddCell(0, $row, 5, "$source_count",  $format);
											$template->AddCell(0, $row, 8, "$srcUp",  $format);
											$Source[$i].=",$subsource";
											
										}
										else{
											# print"Source Not Matched\n";
											print"UPDATE5\n";
											$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Source" => "$source"}});
											$template->AddCell(0, $row, 4, "Source Updated",  $format);
											$Source[$i].=",$subsource";
											
										}
									}
								}
							}
							# else{ $template->AddCell(0, $row, 6, scalar(@moreCategory),  $format);  }

						}
						else{
							print"UPDATE6\n";
							$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Category" => "$subcategory"}});
							$template->AddCell(0, $row, 4, "Category Updated",  $format);
							my @moreSource=split(";",$source);
							if($source){
								foreach my $subsource(@moreSource){
									$subsource=~s/\s+/\\s\*/igs;
									if($Source[$i]!~m/$subsource/is){
										$subsource=~s/\\s\*/ /igs;
										my @SrcCount=split(",",$Source[$i]);
										# print scalar(@SrcCount);
										if(scalar(@SrcCount) > 0){
											# print"inside sourcecount\n";
											push(@SrcCount,"$subsource");
											my $srcUp=join(",",@SrcCount);
											my $source_count=scalar(@SrcCount);
											print"UPDATE7\n";
											$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Source" => "$srcUp","SOurceCount" => "$source_count"}});
											$template->AddCell(0, $row, 4, "Source Updated",  $format);$template->AddCell(0, $row, 5, "$source_count",  $format);
											$template->AddCell(0, $row, 8, "$srcUp",  $format);
											$Source[$i].=",$subsource";
											
										}
										else{
											print"UPDATE8\n";
											$db->brands->update({"BrandId" => "$BrandId[$i]" }, {'$set' => {"Source" => "$source"}});
											$template->AddCell(0, $row, 4, "Source Updated",  $format);
											$Source[$i].=",$subsource";
											
										}
									}
								}
							}
							else{ $template->AddCell(0, $row, 6, scalar(@moreCategory),  $format);  }
						}
					}
				}
			}

		
		}
		$bName=~s/\\s\*/ /igs;

		# $worksheet->write($r,0,"$bName");$worksheet->write($r,1,"$category");$worksheet->write($r,2,"$source");$worksheet->write($r,3,"$incosmos");
		if($category ne ''){
			if($catFlg == 0){
				print"Not Found \n inserting a record...\n";
				$count++;
				my $id="BR".$count;
				# print"Last rec-$id\nincremented..\n";
				$id++;
				print"INSERT1\n";
				my @catCount=split(",",$category);
				my @sorCount=split(",",$source);
				$db->brands->insert({"BrandId" => "$id","BrandName" => "$bName","Source" => "$source","SOurceCount" => "1","Category" => "$category","CategoryCount" => scalar(@catCount),"InCosmos" => "$incosmos"});
				$template->AddCell(0, $row, 4, "Added",  $format);$template->AddCell(0, $row, 5, scalar(@sorCount),  $format);$template->AddCell(0, $row, 6, scalar(@catCount),  $format);
				print"Inserted\n";

			}
		}
		# else{

		# 	$template->AddCell(0, $row, 4, "NotFound",  $format);
		# }
		$row++;
	

}
 $template->SaveAs('Report.xls');

sub uniq {
    return keys %{{ map { $_ => 1 } @_ }};
}