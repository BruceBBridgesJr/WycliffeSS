###############################################################################################
# Program: ProcessWycliffePDF.pl
# This program will process a pdf version of a Social Security Statement 
# correcting errors introduced by the OCR scanning process.  
# All pdf files in the current directory are processed one file at a time
# Step 1 Convert the pdf file to a tiff file because Tesseract-OCR will not accept a pdf file.
# Step 2 Process the tiff file with Tesseract-OCR to produce a plain text file.
# Step 3 Eliminate all lines from the text file except those for earnings data.
# Step 4 Correct errors in earning data using pattern matching
# Step 5 Sort earning ascending by year
# Step 6 Convert the corrected text file to a Microsoft Excel document
###############################################################################################
# Enforce declaration of variables before use
use strict;
use warnings;
# use version 5.010 for new keywords such as "say"
use 5.010;
# Include module for current working directory function
use Cwd;
my $directory = getdcwd();
# Change forward slash to backslash for windows. Using # as delimiter
$directory =~ s#/#\\#g;
# Append a trailing backslash
$directory = $directory . "\\";
# Variables global to main and subs
my $fh_input;
my $fh_output;
my $base_file_name;
my $tiff_file;
my $input_file;
my $input_line;
my @line_elements;
my $ss_name;
my $name_found = 0;
my $text_from_tesseract;
my $text_from_eliminate_lines;
my $text_from_create_para;
my $text_from_process_para;
my $text_from_format_lines;
my $sorted_output;


###############################################################################################
# Open directory containing the program and any input pdf files
###############################################################################################
opendir my $dh, $directory or die "Cannot open directory $directory";

###############################################################################################
# Process all pdf files in the directory
###############################################################################################
foreach my $file_in_dir (readdir $dh) {
	next unless $file_in_dir =~ /\.pdf/;
	my @fn_parts = split /(\.)/, $file_in_dir;
	$base_file_name = $fn_parts[0];
	$input_file = "$fn_parts[0].$fn_parts[2]";
	say "\nProcessing file $input_file";
	# Step 1 Convert the pdf file to a tiff file because Tesseract-OCR will not accept a pdf file.
	# Assign a value to the variable for the output file
	$tiff_file = "$base_file_name".'.tif';
	pdf_to_tiff();
	# Step 2 Process the tiff file with Tesseract-OCT to produce a plain text file.
	tiff_to_text();
	# Step 3 Eliminate all lines except those containing earnings data. 
	# Assign a value to the variable for the outut file
	$text_from_eliminate_lines = "$base_file_name".'_2_eliminate_lines.txt';
	eliminate_lines();
	# Step 4 Correct errors in the earning data using pattern matching
	process_lines();
	format_lines();
	# Step 5 Sort the file on the first field (earning year)
	say "\tSorting earnings by date";
	# Declare a variable for the outut file name of the sort
	$sorted_output = "$ss_name".'earnings.csv';
	system "sort /+1 < \"$directory$text_from_format_lines\" > \"$directory$sorted_output\"";  
	system "del \"$directory$text_from_format_lines\""; # delete sort input file
	# Step 6 Import the corrected csv file into Microsoft Excel
	say "\n\tImporting the comma delimited text file into Excel";
	system "excel.exe /e \"$directory$sorted_output\"";
}
closedir $dh;
say "\n\tProcessing complete for file $input_file";
exit; # end of program

sub pdf_to_tiff {
###############################################################################################
#Step 1 - Convert the PDF file to TIFF so it can be processed by Tesseract-OCR
###############################################################################################
say "\n\tConverting PDF image file to a TIFF image file";
system "pdftotiffcmd.exe -i \"$directory$input_file\" -o \"$directory$tiff_file\" -b 1 -c 0 -m -r 800 -q";
}

sub tiff_to_text {
###############################################################################################
#Step 2 - Process the TIFF file with Tesseract-OCR to produce a text file
###############################################################################################
say "\n\tConverting the TIFF image file to a text file";
# Assign a value to the variable for the outut file
#$text_from_tesseract = "$base_file_name".'_1'; # tesseract adds a .txt suffix to the specified file name
system "tesseract.exe \"$directory$tiff_file\"  \"$directory$base_file_name\" -l eng 2> NUL";
# Delete the input file
system "del \"$directory$tiff_file\"";
}

sub eliminate_lines {
###############################################################################################
#Step 3 - Eliminate all lines except those containing earnings data.
#Earnings date line start with a four digit year followed by a space.
###############################################################################################
say "\n\tEliminating all lines except those containing earnings data";
$text_from_tesseract = $base_file_name . '.txt';
open ($fh_input, '<', "$directory$text_from_tesseract") 
	or die "1 Cannot open input file $directory$text_from_tesseract: $! / $^E\n";
open ($fh_output,  '>', "$directory$text_from_eliminate_lines") 
	or die "2 Cannot open output file $directory$text_from_eliminate_lines: $! / $^E\n";
while ($input_line = <$fh_input>) {
	if ($input_line =~ m/^\d{4} /) {
		# line starts with a four digit year. Output this line
		$input_line =~ s/,//g; # Eliminate commas within amounts
		$input_line =~ s/ +$//; # Eliminate spaces at end of line
		print $fh_output "\n$input_line";
	} elsif (!$name_found) {
		# Capture name for use in output file name
		if ($input_line =~ m/January/) {
				$name_found = 1;
				@line_elements = split 'January', $input_line, 2;
				# The first arrary element contains the name preceding the month name
				$ss_name = $line_elements[0];
				$ss_name =~ s/\.//g;
		} elsif ($input_line =~ m/February/) {
				$name_found = 1;
				@line_elements = split 'February', $input_line, 2;
				$ss_name = $line_elements[0];
				$ss_name =~ s/\.//g;
		} elsif ($input_line =~ m/March/) {
				$name_found = 1;
				@line_elements = split 'March', $input_line, 2;
				$ss_name = $line_elements[0];
				$ss_name =~ s/\.//g;
		} elsif ($input_line =~ m/April/) {
				$name_found = 1;
				@line_elements = split 'April', $input_line, 2;
				$ss_name = $line_elements[0];
				$ss_name =~ s/\.//g;
		} elsif ($input_line =~ m/May/) {
				$name_found = 1;
				@line_elements = split 'May', $input_line, 2;
				$ss_name = $line_elements[0];
				$ss_name =~ s/\.//g;
		} elsif ($input_line =~ m/June/) {
				$name_found = 1;
				@line_elements = split 'June', $input_line, 2;
				$ss_name = $line_elements[0];
				$ss_name =~ s/\.//g;
		} elsif ($input_line =~ m/July/) {
				$name_found = 1;
				@line_elements = split 'July', $input_line, 2;
				$ss_name = $line_elements[0];
				$ss_name =~ s/\.//g;
		} elsif ($input_line =~ m/August/) {
				$name_found = 1;
				@line_elements = split 'August', $input_line, 2;
				$ss_name = $line_elements[0];
				$ss_name =~ s/\.//g;
		} elsif ($input_line =~ m/September/) {
				$name_found = 1;
				@line_elements = split 'September', $input_line, 2;
				$ss_name = $line_elements[0];
				$ss_name =~ s/\.//g;
		} elsif ($input_line =~ m/October/) {
				$name_found = 1;
				@line_elements = split 'October', $input_line, 2;
				$ss_name = $line_elements[0];
				$ss_name =~ s/\.//g;
		} elsif ($input_line =~ m/November/) {
				$name_found = 1;
				@line_elements = split 'November', $input_line, 2;
				$ss_name = $line_elements[0];
				$ss_name =~ s/\.//g;
		} elsif ($input_line =~ m/December/) {
				$name_found = 1;
				@line_elements = split 'December', $input_line, 2;
				$ss_name = $line_elements[0];
				$ss_name =~ s/\.//g;
		} else {
				if (!$name_found) {
						$ss_name = "Social Security"
				} # end if !$name_found
		} # end if $input_line =~
	} # end elsif !$name_found
} # end while
close $fh_input;
#system "del \"$directory$text_from_tesseract\""; # delete the input file
close $fh_output;
} # end of subroutine eliminate_lines

sub process_lines {
##############################################################################################
#Step 4.1 - Process the input file one line at a time
###############################################################################################
say "\n\tProcessing the text file to remove OCR errors";
# Assign a value to the variable for the outut file
$text_from_process_para = "$base_file_name".'_3_process_para.txt';
open ($fh_input, '<', "$directory$text_from_eliminate_lines") 
		or die "2 Can't open input file $directory$text_from_eliminate_lines: $! / $^E\n";
open ($fh_output,  '>', "$directory$text_from_process_para") 
		or die "3 Can't open output file $directory$text_from_process_para: $! / $^E\n";
while ($input_line = <$fh_input>) {
		###############################################################################################
		# Eliminate the last year of earnings which is incomplete.
		# That is a four digit year followed by a space, followed by one or more alpha characters
		################################################################################################
		$input_line =~ s/ \d{4} [A-Za-z\s]+/\n/g;

		###############################################################################################
		#Eliminate space at the beginning of a paragraph
		################################################################################################
		$input_line =~ s/\A +//g;

		###############################################################################################
		# Write the line to output if it starts with a year
		###############################################################################################
		if ($input_line =~ m/^\d{4}/) {
				print $fh_output "$input_line";
		} # end if
} # end while
close $fh_input;
system "del \"$directory$text_from_eliminate_lines\""; # delete the input file
close $fh_output;
} # end of subroutine process_lines

sub format_lines {
###############################################################################################
# Step 4.2 - Format earning lines into a csv format that can be input to Excel
###############################################################################################
say "\n\tFormatting csv lines for Excel\n";
# Declare a variable for the outut file name in the sub-routine
$text_from_format_lines = "$base_file_name".'_4_pdf.txt';
open INPUT, '<', "$text_from_process_para" or die "3 Can't open input file $directory$text_from_process_para: $!\n";
open OUTPUT,  '>', "$directory$text_from_format_lines" or die "4 Can't open output file $directory$text_from_format_lines: $!\n";
#undef $/; # undefine the newline character
while ($input_line = <INPUT>) {
	if ($input_line =~ m/^\R/) { 
		# This line only contains a CR/LF, skip it #
	} else {
		###############################################################################################
		#Correct 57 that S/B 57.
		#################################################################################################
		$input_line =~ s/5L/57/g;

		###############################################################################################
		# Format earnings lines
		#################################################################################################
		chomp ($input_line);
		if ($input_line =~ m/\d{4} +\d+ +\d+/) {
			$input_line =~ s/(\d{4}) +(\d+) +\d+/$1,$2\n/g;
		} elsif ($input_line =~ m/\d{4} +\d+ +\d+ \d{4} +\d+ +\d+/) {
			$input_line =~ s/(\d{4}) +(\d+) +\d+ (\d{4}) +(\d+) +\d+/$1,$2\n$3,$4\n/g;
		} elsif ($input_line =~ m/\d{4} +\d+/) {
			$input_line =~ s/(\d{4}) +(\d+)/$1,$2\n/g;
		} 
		
		###############################################################################################
		#Eliminate all spaces
		#################################################################################################
		$input_line =~ s/\ +//g; # Eliminate spaces in line
		if ($input_line =~ m/^\d+/) {
			print OUTPUT "$input_line";
		}	
	}
}
close INPUT;
system "del \"$directory$text_from_process_para\""; # delete the input file
close OUTPUT;
} # end of subroutine format_lines

