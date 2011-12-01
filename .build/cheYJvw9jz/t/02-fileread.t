use strict;
use Test::More;
use Test::Deep;
use Test::Harness;
use Win32::ExcelSimple qw/transpose_array cell2aa aa2cell/;

my $es = Win32::ExcelSimple->new(test.xls);



is($es->read_cell('B1'),  1, "read data from cell B1");
is($es->read_row('A1',4), (undef, 1,2,3), "read data from row A1");
is($es->read_col('B1',4), (1,1,1,1),      "read data from col B1");
is($es->read_rect('B1:C2'),  ([1,2],[1,2]), "read data from a rectangle");
}
done_testing;

