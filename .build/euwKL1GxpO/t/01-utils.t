use strict;
use Data::Dumper;
use Test::More;
use Test::Deep;
use Test::Harness;
use Win32::ExcelSimple; 




is(aa2cell('AB3'),  [28, 3], "convert AB3 to [28, 3]");
is(cell2aa([2,34]), 'B34', "convert (2,34) to B34");
done_testing;

