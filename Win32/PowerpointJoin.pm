#!/usr/bin/env perl
package Win32::PowerpointJoin;

use strict;
use Carp;
use Data::Dumper;
use autodie qw(open close);
use Win32::PowerPoint;
use Win32::OLE::Const 'Microsoft PowerPoint';

use Cwd;

our $AUTOLOAD;
my $start="start.pptx";
sub merge {
    my ($config_file, $start_with, $output) = @_;
    my $config = &parse_config($config_file);
    &process_charts($config, $start_with, $output);
}

sub parse_config {
    my $file = shift;
    my %config;
    open my $IN, '<', $file;
    my @lines = <$IN>;
    close $IN;
    my $current_file;
    for(my $i=0; $i<@lines; $i++) {
        my $line = $lines[$i];
        chomp $line;
        next if $line =~ /^$/ || $line =~ /^[#;]/;
        if ($line =~ /^file=(.+)$/) {
            $current_file = $1;
            $config{$current_file}{slides} = 0;#if configFile does not contain slides info,set to copy all slides.
        } elsif ( $line =~ /^slides=(.+)$/) {
            $config{$current_file}{slides} = $1;
        } else {
            die "Unrecognized line in $file $line";
        }
    }
    return \%config;#return the File&Slides list 
}


# $ole
# $ole->Presentations->Add
# $ppt = $ole->Presentations->Open(filename)
# check file exists
# $ppt->Slides

sub process_charts {
    my ($config, $opts) = @_;   #$config is the File&Slides list
 
    my $start_with = $opts->{start_with} || undef;
    my $output     = $opts->{output}     || undef;
    if ($start_with && $output) {
        die "Only specify a start-with ppt or an output ppt, but not both";
    }

    my $save_file = $start_with || $output;
    $save_file =~ s{\/}{\\}g;#change F:/ to F:\\

    my %config = %$config;  #$config is the File&Slides list
    my $pp = Win32::PowerPoint->new;

    # Get a count of the slides
    foreach my $key ( sort keys %config) {# to get count of every pptFile's slides
        my $file = $key;
        #$file =~ s{\\}{\\\\}g;
        $file =~ s{\/}{\\\\}g;  #change c:\1 to c:\\1 but why?
        print "Opening file $file\n";
        if (! -r $file) {
            die "Can't find file $file\n";
        }
        my $ppt_insert = $pp->application->Presentations->Open($file);
        my $n_slides = $ppt_insert->Slides->Count;
        $ppt_insert->Close;
        $config{$key}{count} = $n_slides;
        print "$file ($n_slides slides)\n";
    }

    my $ppt;
    if ($start_with) {
        $ppt = $pp->application->Presentations->Open($save_file);
    } else {
        $pp->new_presentation;
        $ppt = $pp->presentation;
    }

    foreach my $key ( sort keys %config ) {
        my $file = $key;
        $file =~ s{\\}{\\\\}g;
        $file =~ s{\/}{\\\\}g;

        my $count = $config{$key}{count}; 

        print "File=$file ($count slides)\n";
        # app.Slides.InsertFromFile("filename", "index", "slideStart", "slideend")
        
        my $last = $ppt->Slides->Count;
        if (! exists $config{$key}{slides} || $config{$key}{slides} eq 0 ) {
            print "   > inserting all (1-$count)\n";
            my $n_inserted = $ppt->Slides->InsertFromFile($file, $last, 1, $count);
            if ( ! $n_inserted ) {
                print "Error: No slides were inserted: " . Win32::OLE->LastError() . "\n";
                print "File:  $file\n";
                print "Count: $count\n";
                print "Range: all\n";
                exit;
            }
        } else {
            my $slides = $config{$key}{slides};
            $slides =~ s/\s+//g;
            my @ranges = split(/[;,]/, $slides);
            foreach my $range (@ranges) {
                my $last1 = $ppt->Slides->Count;

                my ($start, $end); 
                if ($range =~ /^(\d+)-(\d+)$/) {
                    ($start, $end) = ($1, $2);
                } elsif ($range =~ /^\d+$/) {
                    ($start, $end) = ($range, $range);
                } else {
                    print "Error: Invalid range\n";
                    print "File:  $file\n";
                    print "Count: $count\n";
                    print "Range: $range\n";
                    exit;
                }
                if ($start > $end || $start > $count || $end > $count) {
                    print "Error: Invalid range\n";
                    print "File:  $file\n";
                    print "Count: $count\n";
                    print "Start: $start\n";
                    print "End:   $end\n";
                    exit;
                }
                my $n_inserted = $ppt->Slides->InsertFromFile($file, $last1, $start, $end);#insert slides from ppt file to ppt from slide start to end 
                print "   > inserted $n_inserted slides (range $range)\n";
                if ( ! $n_inserted ) {
                    print "Error: No slides were inserted: " . Win32::OLE->LastError() . "\n";
                    print "File:  $file\n";
                    print "Count: $count\n";
                    print "Range: $range\n";
                    print "Start: $start\n";
                    print "End:   $end\n";
                    exit;
                }
            } 
        }
    }

    print "Saving to $save_file\n";
    my $result = $ppt->SaveAs( $save_file );
    sleep 2;
    $pp->close_presentation;
    undef $pp;

    if (! -f $save_file) {
        print "$save_file did not get saved\n";
    }
}

sub dump_ole {
    my ($name, $obj) = @_;
    print "Properties: $name\n--------------------\n";
    my @k = sort keys %{$obj};
    foreach my $key (sort keys %{$obj}) {
        my $value;
        eval { $value = $obj->{$key} };
        $value = "***Exception: $@";
        $value = "<undef>" unless defined $value;
        $value = '[' . Win32::OLE->QueryObjectType($value) . ']'
            if UNIVERSAL::isa($value, 'Win32::OLE');

        $value = '(' . join(',', @$value) . ')' if ref($value) eq 'ARRAY';
        printf "%s %s %s\n", $key, '.' x (40-length($key)), $value;
    }
    print "\nMethods: $name\n----------------------\n";

    my $typeinfo = $obj->GetTypeInfo();
    my $attr = $typeinfo->_GetTypeAttr();
    my @functions;
    for (my $i = 0; $i< $attr->{cFuncs}; $i++) {
        my $desc = $typeinfo->_GetFuncDesc($i);
        # the call conversion of method was detailed in %$desc
        my $funcname = @{$typeinfo->_GetNames($desc->{memid}, 1)}[0];
        push(@functions, $funcname);
    }
    print join("\n  ->", sort @functions);
    print "\n\n";
}

sub AUTOLOAD {
    my $obj = shift;
    $AUTOLOAD =~ s/^.*:://;
    my $meth = $AUTOLOAD;
    $AUTOLOAD = "SUPER::" . $AUTOLOAD;
    my $retval = $obj->AUTOLOAD(@_);
    unless (defined($retval) || $AUTOLOAD eq 'DESTROY') {
        my $err = Win32::OLE::LastError();
        croak(sprintf("$meth returnled OLE error 0x%08x", $err))
            if ($err);
        return $retval;
    }
}

sub getAllFile{
	my $file; 
	my @dir;
	my $fileList;
	my $dir = Cwd->getcwd; 
	opendir (DIR, $dir) or die "can’t open the directory!"; 
	@dir = readdir DIR; 
	foreach $file (@dir) {
		 if ( $file =~ /\.pptx/ and not $file=~/$start/){
		 	 $fileList.=";".$file;
		 } 
	}
	return ($dir,$fileList);
}



1;
