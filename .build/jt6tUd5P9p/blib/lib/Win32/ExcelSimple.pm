package Win32::ExcelSimple;
use warnings;
use strict;
use Try::Tiny;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use Win32::OLE::NLS qw(:LOCALE :DATE);
use Spreadsheet::Read;   #use cr2cell, cell2cr
# ABSTRACT: a wrap for excel
=head1 NAME
:
Win32::ExcelSimple - new Win32::ExcelSimple!
Please Note!!! this module is based on CELL address!!!!
you can use cr2cell, or cell2cr funcs to translate address easily. 
=head1 VERSION

Version 0.03

=cut

use Exporter;
our @ISA       = qw( Exporter );
our @EXPORT    = qw( cell2cr cr2cell );

our $VERSION = '0.03';
sub new {
	my ($class_name, $file_name) = @_;
    
		my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
    	  // Win32::OLE->new( 'Excel.Application', 'Quit' );
    	$Excel->{DisplayAlerts} = 0;
  		$Win32::OLE::Warn = 2;   
  		my $book = $Excel->Workbooks->Open($file_name);
  		my $self = {
		  'excel_handle'   =>  $Excel,
		  'book_handle'    =>  $book,
	  	   };
	bless $self, $class_name;
	return $self;

}

sub open_sheet{
	die ref $_[0]->{'book_handle'}->Worksheets($_[1]);
	my $t = {
		     sheet_h  => $_[0]->{'book_handle'}->Worksheets($_[1])
		    };
	bless $t, 'Win32::ExcelSimple::Sheet';
}

sub save_excel{

	my $self = shift;
	$self->{ 'book_handle' }->save();
	return 0;

} 
sub saveas_excel{
	my $self = shift;
	my $name = shift;
	$self->{ 'book_handle' }->saveas($name);
	return 0;
}

sub close_excel{
	my $self = shift;
 	$self->{ 'excel_handle' }->Workbooks->close;

}


sub DESTROY{

	my $self = shift;
	$self->close_excel();
#	print "save all and exit!!!\n";


}
package Win32::ExcelSimple::Sheet;
no strict 'subs';

sub read{
	my ($sheet_h, $x1,$y1, $x2, $y2) = @_;
	my $address = cell2cr($x1,$y1) . ':' . cell2cr($x2, $y2);
	my $data = $$sheet_h->Range($address)->{Value};
	return $data;
}


sub get_last_row{
	my $sheet_h = shift;
    return $sheet_h->{'sheet_h'}->UsedRange->Find({What=>"*",
    			SearchDirection=> xlPrevious,
    			SearchOrder=> xlByRows})->{Row};

}

sub get_last_col{
	my $sheet_h = shift;
	return $sheet_h->{'sheet_h'}->UsedRange->Find({What=>"*", 
                  SearchDirection=> xlPrevious,
                  SearchOrder=> xlByColumns})->{Column};
	  }

sub read_cell{
	my ($sheet_h, $row, $col) = @_;
}
sub write_cell{
}
sub write_row{
	my ($sheet_h, $x1,$y1, $data) = @_;
	my  $x2 = $x1;
	my  $y2 =  $y1 + (scalar @{$data});
	my $address = cell2cr($x1,$y1) . ':' . cell2cr($x2, $y2);
	$sheet_h->Range($address)->{Value} = $data;
}

sub write{
	my ($sheet_h, $x, $y, $data) = @_;
	         if( ref $$data[0] != ref []){
				 $sheet_h->write_row($x,$y,$data);
			 }
			 else{
				 for(my $i = 0; $i < (scalar @{$data}); $i ++){
						 $sheet_h->write_row($x, $y, $data->[$i]);
							 $x ++;
				 }
             }
}

sub cell_walk{
    my ($sheet_h, $x1, $y1, $x2, $y2, $callback, $callback_data) = @_;
	for ( my $row = $x1 ; $row <= $x2 ; $row++ ) {
    	for ( my $col = $y1 ; $col <= $y2 ; $col++ ) {
			  $callback->($sheet_h->Cells( $row, $col ), $callback_data);
		}
	}

}

sub whole_walk{
    my ($self, $callback) = @_;
	my $x = [1,1];

my $y = [$self->get_last_row(), $self->get_last_col()];

$self->cell_walk($x, $y, $callback);

}
=head1 SYNOPSIS

Quick summary of what the module does.

Perhaps a little code snippet.

    use Win32::ExcelSimple;

    my $foo = Win32::ExcelSimple->new();
    ...

=head1 EXPORT

A list of functions that can be exported.  You can delete this section
if you don't export anything, such as for a purely object-oriented module.

=head1 SUBROUTINES/METHODS

=head2 transpose_array


=cut





=head1 AUTHOR

Andy Xiao, C<< <andy.xiao at gmail.com> >>

=head1 BUGS

Please report any bugs or feature requests to C<bug-win32-excelsimple at rt.cpan.org>, or through
the web interface at L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=Win32-ExcelSimple>.  I will be notified, and then you'll
automatically be notified of progress on your bug as I make changes.




=head1 SUPPORT

You can find documentation for this module with the perldoc command.

    perldoc Win32::ExcelSimple


You can also look for information at:

=over 4

=item * RT: CPAN's request tracker (report bugs here)

L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=Win32-ExcelSimple>

=item * AnnoCPAN: Annotated CPAN documentation

L<http://annocpan.org/dist/Win32-ExcelSimple>

=item * CPAN Ratings

L<http://cpanratings.perl.org/d/Win32-ExcelSimple>

=item * Search CPAN

L<http://search.cpan.org/dist/Win32-ExcelSimple/>

=back


=head1 ACKNOWLEDGEMENTS


=head1 LICENSE AND COPYRIGHT

Copyright 2011 Andy Xiao.

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.


=cut
1; # End of Win32::ExcelSimple
