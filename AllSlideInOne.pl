#!/usr/bin/env perl

use strict;
use Carp;
use Getopt::Long;
use FindBin qw($Bin);
use Win32::PowerpointJoin;
#use utf8;
use Encode;
#binmode STDIN, ':utf8'; 
#binmode STDOUT, ':utf8'; 
 
my ($HELP, $CONFIG, $OUTPUT, $START_WITH);


my $outfile = './config.txt';#输出config.txt文件
open CONFIGFILE, ">:encoding(gbk)",$outfile or die $!;
(my $d1,my $f1)= &Win32::PowerpointJoin::getAllFile();
#$f1=decode("gb2312",$f1);
foreach my $str (split /;/ , $f1 ){
	if($str ne ""){
		my $ss=$d1."/".$str."\n";
		$ss=~ s{\/}{\\}g;
		$ss="file=".$ss;
		$ss=decode("gbk",$ss);
		print CONFIGFILE $ss;
		#print $ss;
	}
}
close CONFIGFILE; 

$CONFIG="config.txt";
$OUTPUT=undef;
$START_WITH="start.pptx";
if($START_WITH) {
    $START_WITH = "$Bin\\$START_WITH";
}

&Win32::PowerpointJoin::merge($CONFIG, { start_with => $START_WITH, output => $OUTPUT});