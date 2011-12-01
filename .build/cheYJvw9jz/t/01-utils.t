use strict;
use Data::Dumper;
use Test::More;
use Test::Deep;
use Test::Harness;
use Win32::ExcelSimple; 




is(Win32::ExcelSimple::cr2cell('AB3'),  [28, 3], "convert AB3 to [28, 3]");
is(Win32::ExcelSimple::cell2cr([2,34]), 'B34', "convert (2,34) to B34");
done_testing;

