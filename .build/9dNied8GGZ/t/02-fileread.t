use strict;
use Test::More;
use Test::Deep;
use Test::Harness;
use Win32::ExcelSimple;

my $es = Win32::ExcelSimple->new('test.xls');
my $sheet_h = $es->open_sheet('Report');

is($sheet_h->get_last_col(), 5, "read last col");
is($sheet_h->get_last_row(), 7, "read last row");
is($es->read_cell(2,1),  1, "read data from cell B1");
#is($es->read('B1:C2'),  ([1,2],[1,2]), "read data from a rectangle");
done_testing;

