#!/usr/bin/perl

# needs convert from imagemagick in the path

use warnings;
use strict;
use Data::Dumper;
my %Passed_Values;
my $ImageDir='imagedocs/';
my $basedir='/data/vhost/services.disruptiveproactivity.com/docs/pdfexcel/'.$ImageDir;
#my $basedir='./'; 
use CGI qw/param/;

{
	foreach my $p (param()) {
		$Passed_Values{$p}= param($p);
	}

	if (defined $Passed_Values{'action'}) {

		# do somethign
		if (defined $Passed_Values{'sleep'}) {
			sleep($Passed_Values{'sleep'});
		}

		if ($Passed_Values{'action'} eq 'step1') {
			&create_PDF_images();
		} elsif ($Passed_Values{'action'} eq 'step2') {
			&chop_and_make_excel()
		}
	} else {
		&ask_for_pdf();
	}


}



sub create_PDF_images { # create a page with the docs in it.
	use LWP::Simple;
	use File::Temp;
	use Digest::MD5 qw/md5_hex/;
	#return &ask_for_pdf unless ($Passed_Values{'url'}=~ m#\.pdf$#i); #
	my $pdf= get ($Passed_Values{"url"} ) || return &ask_for_pdf('fetch failed');
	my $hashname= md5_hex($pdf);
	my $name= "$basedir/$hashname" ;
	mkdir $name || die "can't mkdir $name: $!";
	open (F, ">$name/original.pdf");
	print F $pdf;
	close(F);

	chdir($name);
	unless (-e "$name/output.jpg" or -e "$name/output-0.jpg") {
		`/usr/local/bin/pdftohtml -xml -stdout  $name/original.pdf | /usr/local/bin/xmllint --recover - > $name/output.xml`;
		my $head= `head $name/output.xml`;
		my ($width)= $head=~ m#width="(\d+)"#;
		my ($height)= $head=~ m#height="(\d+)"#;
		#warn "/usr/local/bin/convert -size {$width}x{$height} $name/original.pdf $name/output.jpg";
		`/usr/local/bin/convert -size "${width}x${height}" $name/original.pdf $name/output.jpg`;
	}

	# do this again in case the images were generated and this was a reload
	my $head= `head $name/output.xml`;
	my ($width)= $head=~ m#width="(\d+)"#;
	my ($height)= $head=~ m#height="(\d+)"#;
		# onmouseover, add reference to the previous box

print <<EOhtml;
Content-Type: text/html\n\n
<html>
<head>
	<script type="text/javascript" src="/pdfexcel/js/prototype.js" language="javascript"></script>
	<script type="text/javascript" src="/pdfexcel/js/scriptaculous/scriptaculous.js?load=builder,dragdrop" language="javascript"></script>
	<script type="text/javascript" src="/pdfexcel/js/cropper/cropper.js" language="javascript"></script>

	<script>

        function cache_box( fields, coords ) {
                \$( 'image_'+fields+'_x1' ).value = coords.x1;
                \$( 'image_'+fields+'_y1' ).value = coords.y1;
                \$( 'image_'+fields+'_x2' ).value = coords.x2;
                \$( 'image_'+fields+'_y2' ).value = coords.y2;
                \$( 'image_'+fields+'_boxdrawn' ).value = "1";
        };

	var current_image=1;

	</script>
	<style>
		.submitbox {
			position: fixed;	
			background-color: white;
			float: right;
			top: 5px;
			right: 300px;
			width: 20em;
			z-index: 3;
		}
	</style>
</head>
<body>
<form method="GET" action="pdf2xls.pl">
<div class="submitbox">
	Submit for processing: <input type="submit" value="Submit" />
</div>
<p><br /></p><p><a href="./pdf2xls.pl?action=step1;sleep=25;url=$Passed_Values{'url'}">if nothing here, click to reload since processing may take a while on large pdfs</a></p>
<h2>Draw boxes over the tables</h2>

	<input type="hidden" name="action" value="step2" />
	<input type="hidden" name="this_image" value="" id="this_image" />
	<input type="hidden" name="code" value="$hashname" />
	<input type="hidden" name="url" value="$Passed_Values{'url'}" />
	var image_count=0;
EOhtml
	opendir (D, "$name");
	my @images;
	while (my $file =readdir(D)) {
		if ($file=~ m#jpg#) { push @images, $file; }
	}
	closedir(D);
	my $thisid=0;
	foreach my $image (sort @images ) {
		# figure out how big image is, add 
		print <<EOhtml;
	<input type="hidden" id="image_${thisid}_boxdrawn" name="image_${thisid}_boxdrawn" value="" />
	<input type="hidden" id="image_${thisid}_x1" name="image_${thisid}_x1" value="" />
	<input type="hidden" id="image_${thisid}_y1" name="image_${thisid}_y1" value="" />
	<input type="hidden" id="image_${thisid}_x2" name="image_${thisid}_x2" value="" />
	<input type="hidden" id="image_${thisid}_y2" name="image_${thisid}_y2" value="" />
	<img src="$ImageDir/$hashname/$image" style="border: 2px solid black;" alt=" " width="$width" height="$height" onmouseover="current_image='$image';" id="image_$thisid"  /><br />
<script type="text/javascript">
                new Cropper.Img( 'image_$thisid', { onEndCrop: cache_box_$thisid });
		function cache_box_$thisid(coords){ cache_box("$thisid", coords) };
		image_count++;
</script>

EOhtml
		$thisid++;
	}

	$thisid--; # since we incrememnt at the end of the run
print <<EOscript;
	<input type="hidden" name="image_count" value="$thisid" />
</form>
</body>
</html>
EOscript
}


sub chop_and_make_excel {
	# we get x/y box information for each image that it was done for.
	#	if not, remember the previous ones we wa
	# foreach page in order
	# 	store x/y size info if we have it for that page
	#	use the stored values
	#if ($Passed_Values{"image_${thisid}_boxdrawn"}) {

	#}


		# anything that starts or ends in the box is included.
	use XML::Simple;
	use JSON qw/to_json/;
	use Spreadsheet::WriteExcel;

	&ask_for_pdf if $Passed_Values{'code'}=~ m#[^0-9A-Z]#i; 

	my $xml_file= XMLin("$basedir/$Passed_Values{'code'}/output.xml", Cache => 'Storable', ForceArray => 1);
	my ($x1, $x2, $y1, $y2);

	#use Data::Dumper;
	#print "Content-Type: text/plain\n\n".  Dumper $xml_file;
	my $max_page = 1;
	if (ref($xml_file->{'page'}) eq 'ARRAY') {
		$max_page= $xml_file->{'page'}->[-1]->{'number'} -1;
	}

	my $width_factor= 20; #min width of column before they get merged together. 20px
	my $colinfo;
	my $workbook  = Spreadsheet::WriteExcel->new("$basedir/$Passed_Values{'code'}/output.xls");
#print "Content-Type: text/plain\n\n";
#print  Dumper $xml_file;
	my %colmapping;
	my $rowcount;
	foreach my $pageno (0 .. $max_page) { 
		# get bounding box (or cached values from last pass) or skip
		if (defined $Passed_Values{"image_${pageno}_boxdrawn"}  and $Passed_Values{"image_${pageno}_boxdrawn"}  ne '' ) {
			$x1= $Passed_Values{"image_${pageno}_x1"};
			$y1= $Passed_Values{"image_${pageno}_y1"};
			$x2= $Passed_Values{"image_${pageno}_x2"};
			$y2= $Passed_Values{"image_${pageno}_y2"};

#warn "page $pageno";
			my $last_top;
			foreach my $textblock (@{$xml_file->{'page'}->[$pageno]->{'text'}}) {
				# next if left outside bounding box,
				next if $textblock->{'left'} > $x2;
				next if $textblock->{'left'} < $x1;
				# next if top is outside the bounding box
				next if $textblock->{'top'} < $y1;
				next if $textblock->{'top'} > $y2;

				if (abs($textblock->{'top'} - $last_top) >= 10) {
					$rowcount++;
				}
				#$rowcount++ if not defined $colinfo->{'top'}->{$this_top_offset};
				my $leftindex= int($textblock->{'left'} / $width_factor);

				my $this_top_offset=$colinfo->{'top'}/9;
				$colinfo->{'top'}->{$this_top_offset}++;
				$colinfo->{'left'}->{$leftindex}++;
				$colinfo->{'left_row'}->{$leftindex}->{$rowcount}++;
				$colinfo->{'count'}++;
			}

			# produce map of leftindex mapping to spreadsheet columns
			my $excel_col=0;
			%colmapping=();
			foreach my $index (sort {$a <=> $b} keys %{$colinfo->{'left'}}) {
				#warn $index;
				#warn $colinfo->{'left'};
				$colmapping{$index}= $excel_col++;
			}

			# have now rebuild data structures given new bounding box. 
		}

#warn Dumper \%colmapping;
		$rowcount=0; # reset it before reuse

		my $worksheet = $workbook->add_worksheet("page_$pageno");


		my $last_top=-1;
		# now do all rows
		foreach my $textblock (@{$xml_file->{'page'}->[$pageno]->{'text'}}) {
			# next if top is outside the bounding box
			next if $textblock->{'left'} > $x2;
			next if $textblock->{'left'} < $x1;
			# next if left outside bounding box,
			next if $textblock->{'top'} < $y1;
			next if $textblock->{'top'} > $y2;

			# add to worksheet
			if (10 < abs($last_top - $textblock->{'top'})) {
				$rowcount++; 
				$last_top= $textblock->{'top'};
			}

			my $column=$colmapping{int($textblock->{'left'} / $width_factor)};
			my $content='';
			my $format=undef; 
			if (defined $textblock->{'content'}) { 
				$content= $textblock->{'content'};
			} else {
				$content= undef;

				my %skipkeys;
				foreach my $k (qw/width left top content height font/) {
					$skipkeys{$k}=1;
				}
				$format = $workbook->add_format(); # Add a format
				$format->set_align('center');
				foreach my $key (keys %$textblock) {
					next if defined $skipkeys{$key};
					$content= $textblock->{$key}; #->{'value'}

					if ($key =~ m#^b%#i) { $format->set_bold(); }
					if ($key =~ m#^i%#i) { $format->set_italic(); }

					if (ref($content)) {
						$content= $textblock->{$key}->{$content} 
						# need an alt approach if we find 3 levels of hierarchy: <font><b><i>
					}
				}
				# TODO  need to find the unusual key in the hash, and drill down to find the value of it, possibly multiple levels
				#    possibly set a format on the string to carry some formatting onwards (headers as bold etc)

			}
			if (defined $content) { $worksheet->write($rowcount, $column, $content, $format); }
			$worksheet->write_comment($rowcount, $column, to_json($textblock, {utf8 => 1}));

		}

	}
	$workbook->close();
	print "Location: $ImageDir/$Passed_Values{'code'}/output.xls\n\n";

}

sub ask_for_pdf {
	my $warning= shift;

	if (defined $warning) {
		$warning= "<div class=\"bigredbox\">$warning</div>";
	}
print "Content-Type: text/html\n\n";
print <<EOhtml;

<h1>PDF to Excel</h1>

<form method="get" action="./pdf2xls.pl">
	<input type="hidden" name="action" value="step1" />
	URL of <strong>PDF</strong>: <input type="input" name="url" /> <input type="submit" value="Process" />
	<br />
	<input type="checkbox" name="output_excel" />Straight to excel?
</form>

testing links:
<ul>
	<li>run <a href="pdf2xls.pl?action=step1;url=http://www.northamptonshire.gov.uk/en/councilservices/Council/transparency/Documents/PDF%20Documents/Monthly%20spend%20-%20Health%20and%20Adult%20Social%20Services%20May%202010.pdf">http://www.northamptonshire.gov.uk/en/councilservices/Council/transparency/Documents/PDF%20Documents/Monthly%20spend%20-%20Health%20and%20Adult%20Social%20Services%20May%202010.pdf</a></li>
	<li>run <a
	href="pdf2xls.pl?action=step1;url=http://www.fareham.gov.uk/pdf/finance/payments500.pdf">http://www.fareham.gov.uk/pdf/finance/payments500.pdf</a>
	<li><a
	href="pdf2xls.pl?action=step1;url=http://www.manchester.gov.uk/download/13272/manchester_joint_strategic_needs_assessment_2008-2013_ward_factsheet_withington">http://www.manchester.gov.uk/download/13272/manchester_joint_strategic_needs_assessment_2008-2013_ward_factsheet_withington</a></li>

</ul>

For programmers:
	<li>json of the raw parse is dumped into all the excel comment fields where there's stuff in the text - you can just ignore the actual content
	<li>if you're careful, you can, possibly, try getting with a full size bounding box for page 0 and it might work (for some PDFs), which means you can automate this in 2 get requests. We'll work on this simple. The code you get back is the MD5 hex checksum of the pdf you ask it for.
	<li>It requires the pdftohtml to be in /usr/local/bin (the one with the -xml option - pdf<i>2</i>html wont do), and the following perl modules:
		<ul><li>use CGI qw/param/;
        	<li>use LWP::Simple;
        	<li>use File::Temp;
        	<li>use Digest::MD5 qw/md5_hex/;
        	<li>use XML::Simple;
        	<li>use JSON qw/to_json/;
        	<li>use Spreadsheet::WriteExcel;</li></ul></li>
	<li>Please don't hit this script with automated scrapers yet. At some point soon, I'll stick it on github so you can run it yourself once I'm happy it's actually substantially ok and passes the minimum stuff you need it to pass (at the moment, it has a couple of major missing bits).</li>
EOhtml


}
