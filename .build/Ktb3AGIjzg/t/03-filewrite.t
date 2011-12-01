use strict;
use Test::More;
use Test::Deep;
use Test::Harness;
use Win32::ExcelSimple;

my $es = Win32::ExcelSimple->new(test.xls);



is($es->write_cell('B1'),  1, "write data from cell B1");
is($es->write_row('A1'), (undef, 1,2,3), "write data to row A1");
is($es->write_col('B1'), (1,1,1,1),      "write data to col B1");
is($es->write('B1'),  ([1,2],[1,2]), "write data to a rectangle");
}
done_testing;

